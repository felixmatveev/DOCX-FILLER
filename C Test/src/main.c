#include <gtk/gtk.h>
#include <string.h>
#include "docx.h"

typedef struct {
    GtkApplication *app;
    GtkWidget *window;
    GtkWidget *preview_view;
    GtkTextBuffer *preview_buffer;
    GtkWidget *fields_box;      /* vertical box holding one row per field */
    GtkWidget *fields_scroller;
    GtkWidget *no_doc_label;    /* shown in the fields panel when nothing is loaded */
    GtkWidget *save_button;
    GtkWidget *header_subtitle;

    DocxDocument *doc;          /* currently loaded document, or NULL */
    GHashTable *values;         /* field name -> current text value (owned) */
    char *loaded_path;
} AppState;

static void refresh_preview(AppState *state) {
    docx_render_preview(state->doc, state->preview_buffer, state->values);
}

static void on_entry_changed(GtkEntry *entry, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    const char *field_name = (const char *)g_object_get_data(G_OBJECT(entry), "field-name");
    const char *text = gtk_entry_get_text(entry);
    g_hash_table_replace(state->values, g_strdup(field_name), g_strdup(text));
    refresh_preview(state);
}

static void clear_fields_panel(AppState *state) {
    GList *children = gtk_container_get_children(GTK_CONTAINER(state->fields_box));
    for (GList *l = children; l; l = l->next) {
        gtk_widget_destroy(GTK_WIDGET(l->data));
    }
    g_list_free(children);
}

static void build_fields_panel(AppState *state) {
    clear_fields_panel(state);

    guint count = docx_field_count(state->doc);
    if (count == 0) {
        GtkWidget *label = gtk_label_new("This template has no {{fields}} to fill in.");
        gtk_widget_set_margin_top(label, 12);
        gtk_label_set_line_wrap(GTK_LABEL(label), TRUE);
        gtk_style_context_add_class(gtk_widget_get_style_context(label), "dim-label");
        gtk_box_pack_start(GTK_BOX(state->fields_box), label, FALSE, FALSE, 0);
        gtk_widget_show(label);
        return;
    }

    for (guint i = 0; i < count; i++) {
        const char *name = docx_field_name(state->doc, i);

        GtkWidget *row = gtk_box_new(GTK_ORIENTATION_VERTICAL, 4);
        gtk_widget_set_margin_start(row, 12);
        gtk_widget_set_margin_end(row, 12);
        gtk_widget_set_margin_top(row, 6);

        GtkWidget *label = gtk_label_new(name);
        gtk_label_set_xalign(GTK_LABEL(label), 0.0);
        gtk_style_context_add_class(gtk_widget_get_style_context(label), "field-label");

        GtkWidget *entry = gtk_entry_new();
        gtk_entry_set_placeholder_text(GTK_ENTRY(entry), "Enter value...");
        g_object_set_data_full(G_OBJECT(entry), "field-name", g_strdup(name), g_free);
        g_signal_connect(entry, "changed", G_CALLBACK(on_entry_changed), state);

        gtk_box_pack_start(GTK_BOX(row), label, FALSE, FALSE, 0);
        gtk_box_pack_start(GTK_BOX(row), entry, FALSE, FALSE, 0);

        gtk_box_pack_start(GTK_BOX(state->fields_box), row, FALSE, FALSE, 0);
        gtk_widget_show_all(row);
    }
}

static void show_error_dialog(AppState *state, const char *message) {
    GtkWidget *dialog = gtk_message_dialog_new(GTK_WINDOW(state->window),
                                                GTK_DIALOG_MODAL,
                                                GTK_MESSAGE_ERROR,
                                                GTK_BUTTONS_OK,
                                                "%s", message);
    gtk_dialog_run(GTK_DIALOG(dialog));
    gtk_widget_destroy(dialog);
}

static void reset_document_state(AppState *state) {
    if (state->doc) {
        docx_free(state->doc);
        state->doc = NULL;
    }
    g_hash_table_remove_all(state->values);
    g_free(state->loaded_path);
    state->loaded_path = NULL;
}

static void load_template(AppState *state, const char *path) {
    GError *error = NULL;
    DocxDocument *new_doc = docx_load(path, &error);
    if (!new_doc) {
        char *msg = g_strdup_printf("Couldn't open this file:\n%s", error ? error->message : "unknown error");
        show_error_dialog(state, msg);
        g_free(msg);
        if (error) g_error_free(error);
        return;
    }

    reset_document_state(state);
    state->doc = new_doc;
    state->loaded_path = g_strdup(path);

    build_fields_panel(state);
    refresh_preview(state);

    char *base = g_path_get_basename(path);
    char *subtitle = g_strdup_printf("%s", base);
    gtk_label_set_text(GTK_LABEL(state->header_subtitle), subtitle);
    g_free(subtitle);
    g_free(base);

    gtk_widget_set_sensitive(state->save_button, TRUE);
}

static void on_open_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;

    GtkWidget *dialog = gtk_file_chooser_dialog_new("Open Word Template",
                                                      GTK_WINDOW(state->window),
                                                      GTK_FILE_CHOOSER_ACTION_OPEN,
                                                      "_Cancel", GTK_RESPONSE_CANCEL,
                                                      "_Open", GTK_RESPONSE_ACCEPT,
                                                      NULL);

    GtkFileFilter *filter = gtk_file_filter_new();
    gtk_file_filter_set_name(filter, "Word documents (*.docx)");
    gtk_file_filter_add_pattern(filter, "*.docx");
    gtk_file_chooser_add_filter(GTK_FILE_CHOOSER(dialog), filter);

    if (gtk_dialog_run(GTK_DIALOG(dialog)) == GTK_RESPONSE_ACCEPT) {
        char *path = gtk_file_chooser_get_filename(GTK_FILE_CHOOSER(dialog));
        load_template(state, path);
        g_free(path);
    }
    gtk_widget_destroy(dialog);
}

static void on_save_clicked(GtkButton *button, gpointer user_data) {
    AppState *state = (AppState *)user_data;
    if (!state->doc) return;

    GtkWidget *dialog = gtk_file_chooser_dialog_new("Save Filled Document",
                                                      GTK_WINDOW(state->window),
                                                      GTK_FILE_CHOOSER_ACTION_SAVE,
                                                      "_Cancel", GTK_RESPONSE_CANCEL,
                                                      "_Save", GTK_RESPONSE_ACCEPT,
                                                      NULL);
    gtk_file_chooser_set_do_overwrite_confirmation(GTK_FILE_CHOOSER(dialog), TRUE);

    GtkFileFilter *filter = gtk_file_filter_new();
    gtk_file_filter_set_name(filter, "Word documents (*.docx)");
    gtk_file_filter_add_pattern(filter, "*.docx");
    gtk_file_chooser_add_filter(GTK_FILE_CHOOSER(dialog), filter);

    if (state->loaded_path) {
        char *base = g_path_get_basename(state->loaded_path);
        char *no_ext = g_strdup(base);
        char *dot = strrchr(no_ext, '.');
        if (dot) *dot = '\0';
        char *suggested = g_strdup_printf("%s-filled.docx", no_ext);
        gtk_file_chooser_set_current_name(GTK_FILE_CHOOSER(dialog), suggested);
        g_free(suggested);
        g_free(no_ext);
        g_free(base);
    } else {
        gtk_file_chooser_set_current_name(GTK_FILE_CHOOSER(dialog), "filled.docx");
    }

    if (gtk_dialog_run(GTK_DIALOG(dialog)) == GTK_RESPONSE_ACCEPT) {
        char *path = gtk_file_chooser_get_filename(GTK_FILE_CHOOSER(dialog));

        char *final_path = path;
        gboolean owns_final_path = FALSE;
        if (!g_str_has_suffix(path, ".docx")) {
            final_path = g_strdup_printf("%s.docx", path);
            owns_final_path = TRUE;
        }

        GError *error = NULL;
        if (docx_save(state->doc, final_path, state->values, &error)) {
            char *msg = g_strdup_printf("Saved filled document to:\n%s", final_path);
            GtkWidget *ok = gtk_message_dialog_new(GTK_WINDOW(state->window), GTK_DIALOG_MODAL,
                                                     GTK_MESSAGE_INFO, GTK_BUTTONS_OK, "%s", msg);
            gtk_dialog_run(GTK_DIALOG(ok));
            gtk_widget_destroy(ok);
            g_free(msg);
        } else {
            char *msg = g_strdup_printf("Couldn't save the document:\n%s", error ? error->message : "unknown error");
            show_error_dialog(state, msg);
            g_free(msg);
            if (error) g_error_free(error);
        }

        if (owns_final_path) g_free(final_path);
        g_free(path);
    }
    gtk_widget_destroy(dialog);
}

static void apply_css(void) {
    GtkCssProvider *provider = gtk_css_provider_new();
    const char *css =
        ".field-label { font-weight: 600; font-size: 90%; opacity: 0.8; }\n"
        "textview { padding: 16px; }\n"
        "textview text { background-color: #ffffff; }\n";
    gtk_css_provider_load_from_data(provider, css, -1, NULL);
    gtk_style_context_add_provider_for_screen(gdk_screen_get_default(),
                                               GTK_STYLE_PROVIDER(provider),
                                               GTK_STYLE_PROVIDER_PRIORITY_APPLICATION);
    g_object_unref(provider);
}

static void on_app_activate(GtkApplication *app, gpointer user_data) {
    const char *startup_path = (const char *)user_data;
    AppState *state = g_new0(AppState, 1);
    state->app = app;
    state->values = g_hash_table_new_full(g_str_hash, g_str_equal, g_free, g_free);

    apply_css();

    state->window = gtk_application_window_new(app);
    gtk_window_set_default_size(GTK_WINDOW(state->window), 980, 640);

    GtkWidget *header = gtk_header_bar_new();
    gtk_header_bar_set_show_close_button(GTK_HEADER_BAR(header), TRUE);
    gtk_header_bar_set_title(GTK_HEADER_BAR(header), "DocxFill");
    state->header_subtitle = gtk_label_new(NULL);
    gtk_header_bar_set_subtitle(GTK_HEADER_BAR(header), "No template loaded");
    gtk_window_set_titlebar(GTK_WINDOW(state->window), header);

    GtkWidget *open_button = gtk_button_new_with_label("Open Template…");
    g_signal_connect(open_button, "clicked", G_CALLBACK(on_open_clicked), state);
    gtk_header_bar_pack_start(GTK_HEADER_BAR(header), open_button);

    state->save_button = gtk_button_new_with_label("Save As…");
    gtk_style_context_add_class(gtk_widget_get_style_context(state->save_button), "suggested-action");
    gtk_widget_set_sensitive(state->save_button, FALSE);
    g_signal_connect(state->save_button, "clicked", G_CALLBACK(on_save_clicked), state);
    gtk_header_bar_pack_end(GTK_HEADER_BAR(header), state->save_button);

    GtkWidget *paned = gtk_paned_new(GTK_ORIENTATION_HORIZONTAL);

    /* --- left: preview --- */
    GtkWidget *preview_scroller = gtk_scrolled_window_new(NULL, NULL);
    gtk_widget_set_hexpand(preview_scroller, TRUE);
    gtk_widget_set_vexpand(preview_scroller, TRUE);

    state->preview_view = gtk_text_view_new();
    gtk_text_view_set_editable(GTK_TEXT_VIEW(state->preview_view), FALSE);
    gtk_text_view_set_cursor_visible(GTK_TEXT_VIEW(state->preview_view), FALSE);
    gtk_text_view_set_wrap_mode(GTK_TEXT_VIEW(state->preview_view), GTK_WRAP_WORD);
    gtk_text_view_set_left_margin(GTK_TEXT_VIEW(state->preview_view), 8);
    gtk_text_view_set_right_margin(GTK_TEXT_VIEW(state->preview_view), 8);
    state->preview_buffer = gtk_text_view_get_buffer(GTK_TEXT_VIEW(state->preview_view));
    gtk_text_buffer_set_text(state->preview_buffer, "Open a .docx template to see it here.", -1);

    gtk_container_add(GTK_CONTAINER(preview_scroller), state->preview_view);

    GtkWidget *preview_frame = gtk_frame_new(NULL);
    gtk_container_add(GTK_CONTAINER(preview_frame), preview_scroller);
    gtk_widget_set_margin_start(preview_frame, 8);
    gtk_widget_set_margin_top(preview_frame, 8);
    gtk_widget_set_margin_bottom(preview_frame, 8);

    /* --- right: fields panel --- */
    GtkWidget *right_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 0);
    gtk_widget_set_size_request(right_box, 280, -1);

    GtkWidget *fields_title = gtk_label_new(NULL);
    gtk_label_set_markup(GTK_LABEL(fields_title), "<b>Fields</b>");
    gtk_label_set_xalign(GTK_LABEL(fields_title), 0.0);
    gtk_widget_set_margin_start(fields_title, 12);
    gtk_widget_set_margin_top(fields_title, 12);
    gtk_widget_set_margin_bottom(fields_title, 4);
    gtk_box_pack_start(GTK_BOX(right_box), fields_title, FALSE, FALSE, 0);

    state->fields_scroller = gtk_scrolled_window_new(NULL, NULL);
    gtk_widget_set_vexpand(state->fields_scroller, TRUE);
    gtk_scrolled_window_set_policy(GTK_SCROLLED_WINDOW(state->fields_scroller),
                                    GTK_POLICY_NEVER, GTK_POLICY_AUTOMATIC);

    state->fields_box = gtk_box_new(GTK_ORIENTATION_VERTICAL, 0);
    gtk_container_add(GTK_CONTAINER(state->fields_scroller), state->fields_box);
    gtk_box_pack_start(GTK_BOX(right_box), state->fields_scroller, TRUE, TRUE, 0);

    state->no_doc_label = gtk_label_new("Open a .docx template with\n{{field}} placeholders to begin.");
    gtk_label_set_justify(GTK_LABEL(state->no_doc_label), GTK_JUSTIFY_CENTER);
    gtk_style_context_add_class(gtk_widget_get_style_context(state->no_doc_label), "dim-label");
    gtk_widget_set_margin_top(state->no_doc_label, 24);
    gtk_widget_set_margin_start(state->no_doc_label, 12);
    gtk_widget_set_margin_end(state->no_doc_label, 12);
    gtk_box_pack_start(GTK_BOX(state->fields_box), state->no_doc_label, FALSE, FALSE, 0);

    gtk_widget_set_margin_end(right_box, 8);
    gtk_widget_set_margin_top(right_box, 8);
    gtk_widget_set_margin_bottom(right_box, 8);

    gtk_paned_pack1(GTK_PANED(paned), preview_frame, TRUE, FALSE);
    gtk_paned_pack2(GTK_PANED(paned), right_box, FALSE, FALSE);
    gtk_paned_set_position(GTK_PANED(paned), 660);

    gtk_container_add(GTK_CONTAINER(state->window), paned);

    gtk_widget_show_all(state->window);

    g_object_set_data_full(G_OBJECT(state->window), "app-state", state, (GDestroyNotify)NULL);

    if (startup_path && *startup_path) {
        load_template(state, startup_path);
    }
}

int main(int argc, char **argv) {
    GtkApplication *app = gtk_application_new("org.example.docxfill", G_APPLICATION_FLAGS_NONE);
    const char *startup_path = (argc > 1) ? argv[1] : NULL;
    g_signal_connect(app, "activate", G_CALLBACK(on_app_activate), (gpointer)startup_path);
    /* Ignore GApplication's own file-list argv handling; we just want a
     * plain "open this path on launch" convenience for the local build. */
    int status = g_application_run(G_APPLICATION(app), 1, argv);
    g_object_unref(app);
    return status;
}
