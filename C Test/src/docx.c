#include "docx.h"

#include <zip.h>
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/xpath.h>
#include <libxml/xpathInternals.h>
#include <string.h>
#include <stdlib.h>

#define WORDML_NS (const xmlChar *)"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
#define DOCUMENT_XML_PATH "word/document.xml"

/* ---- internal model ------------------------------------------------- */

typedef struct {
    char *text;
    gboolean bold;
    gboolean italic;
    gboolean underline;
} Run;

typedef struct {
    xmlNodePtr node;     /* the <w:p> element in the live xmlDoc */
    GArray *runs;        /* array of Run, in document order */
    GString *full_text;  /* concatenation of all run text in this paragraph */
    gboolean has_field;
} Paragraph;

struct DocxDocument {
    char *src_path;
    xmlDocPtr xmldoc;         /* parsed word/document.xml, owned */
    GArray *paragraphs;       /* array of Paragraph, in document order */
    GPtrArray *field_names;   /* unique field names, owns strings */
};

/* ---- small helpers ---------------------------------------------------- */

static void run_clear(gpointer p) {
    Run *r = (Run *)p;
    g_free(r->text);
}

static void paragraph_clear(gpointer p) {
    Paragraph *para = (Paragraph *)p;
    if (para->runs) g_array_free(para->runs, TRUE);
    if (para->full_text) g_string_free(para->full_text, TRUE);
}

/* Read an entry fully from a zip archive into a NUL-terminated buffer.
 * Caller owns the returned buffer (g_free). Sets *out_len to the byte
 * length (excluding the extra terminating NUL). Returns NULL on error. */
static char *zip_read_entry(struct zip *za, const char *entry_name, gsize *out_len, GError **error) {
    struct zip_stat st;
    zip_stat_init(&st);
    if (zip_stat(za, entry_name, 0, &st) != 0) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_NOENT,
                    "This does not look like a valid .docx file (missing %s)", entry_name);
        return NULL;
    }

    struct zip_file *zf = zip_fopen(za, entry_name, 0);
    if (!zf) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "Could not open %s inside the .docx archive", entry_name);
        return NULL;
    }

    char *buf = g_malloc(st.size + 1);
    zip_int64_t read = zip_fread(zf, buf, st.size);
    zip_fclose(zf);

    if (read < 0 || (zip_uint64_t)read != st.size) {
        g_free(buf);
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "Failed reading %s from the .docx archive", entry_name);
        return NULL;
    }

    buf[st.size] = '\0';
    if (out_len) *out_len = st.size;
    return buf;
}

/* Look at a <w:rPr> node (may be NULL) and fill in bold/italic/underline. */
static void read_run_props(xmlNodePtr rPr, gboolean *bold, gboolean *italic, gboolean *underline) {
    *bold = FALSE;
    *italic = FALSE;
    *underline = FALSE;
    if (!rPr) return;

    for (xmlNodePtr child = rPr->children; child; child = child->next) {
        if (child->type != XML_ELEMENT_NODE) continue;
        if (xmlStrcmp(child->name, (const xmlChar *)"b") == 0) {
            xmlChar *val = xmlGetNsProp(child, (const xmlChar *)"val", WORDML_NS);
            *bold = (!val || xmlStrcmp(val, (const xmlChar *)"false") != 0) && (!val || xmlStrcmp(val, (const xmlChar *)"0") != 0);
            if (val) xmlFree(val);
        } else if (xmlStrcmp(child->name, (const xmlChar *)"i") == 0) {
            xmlChar *val = xmlGetNsProp(child, (const xmlChar *)"val", WORDML_NS);
            *italic = (!val || xmlStrcmp(val, (const xmlChar *)"false") != 0) && (!val || xmlStrcmp(val, (const xmlChar *)"0") != 0);
            if (val) xmlFree(val);
        } else if (xmlStrcmp(child->name, (const xmlChar *)"u") == 0) {
            xmlChar *val = xmlGetNsProp(child, (const xmlChar *)"val", WORDML_NS);
            *underline = (!val || xmlStrcmp(val, (const xmlChar *)"none") != 0);
            if (val) xmlFree(val);
        }
    }
}

/* Concatenate the text of all <w:t> children of a <w:r> run node. */
static char *read_run_text(xmlNodePtr runNode) {
    GString *s = g_string_new(NULL);
    for (xmlNodePtr child = runNode->children; child; child = child->next) {
        if (child->type != XML_ELEMENT_NODE) continue;
        if (xmlStrcmp(child->name, (const xmlChar *)"t") == 0) {
            xmlChar *txt = xmlNodeGetContent(child);
            if (txt) {
                g_string_append(s, (const char *)txt);
                xmlFree(txt);
            }
        } else if (xmlStrcmp(child->name, (const xmlChar *)"tab") == 0) {
            g_string_append_c(s, '\t');
        } else if (xmlStrcmp(child->name, (const xmlChar *)"br") == 0) {
            g_string_append_c(s, '\n');
        }
    }
    return g_string_free(s, FALSE);
}

/* Find all {{field}} occurrences in text and register the (trimmed) field
 * names into field_names (deduplicated). Returns TRUE if at least one
 * placeholder was found in this text. */
static gboolean register_fields_in_text(const char *text, GPtrArray *field_names) {
    gboolean found_any = FALSE;
    const char *p = text;
    while ((p = strstr(p, "{{")) != NULL) {
        const char *close = strstr(p + 2, "}}");
        if (!close) break;
        const char *name_start = p + 2;
        gsize len = (gsize)(close - name_start);

        char *raw = g_strndup(name_start, len);
        char *trimmed = g_strstrip(raw);

        if (*trimmed != '\0') {
            found_any = TRUE;
            gboolean already = FALSE;
            for (guint i = 0; i < field_names->len; i++) {
                if (g_strcmp0((const char *)g_ptr_array_index(field_names, i), trimmed) == 0) {
                    already = TRUE;
                    break;
                }
            }
            if (!already) {
                g_ptr_array_add(field_names, g_strdup(trimmed));
            }
        }
        g_free(raw);
        p = close + 2;
    }
    return found_any;
}

/* Build a substituted copy of `text`, replacing every {{field}} with the
 * looked-up value from `values` (or "" if absent). */
static char *substitute_text(const char *text, GHashTable *values) {
    GString *out = g_string_new(NULL);
    const char *p = text;
    while (*p) {
        if (p[0] == '{' && p[1] == '{') {
            const char *close = strstr(p + 2, "}}");
            if (close) {
                char *raw = g_strndup(p + 2, (gsize)(close - (p + 2)));
                char *name = g_strstrip(raw);
                const char *val = values ? (const char *)g_hash_table_lookup(values, name) : NULL;
                g_string_append(out, val ? val : "");
                g_free(raw);
                p = close + 2;
                continue;
            }
        }
        g_string_append_c(out, *p);
        p++;
    }
    return g_string_free(out, FALSE);
}

/* ---- loading ------------------------------------------------------------ */

DocxDocument *docx_load(const char *path, GError **error) {
    int zerr = 0;
    struct zip *za = zip_open(path, ZIP_RDONLY, &zerr);
    if (!za) {
        zip_error_t ze;
        zip_error_init_with_code(&ze, zerr);
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "Could not open '%s' as a .docx (zip) file: %s", path, zip_error_strerror(&ze));
        zip_error_fini(&ze);
        return NULL;
    }

    gsize xml_len = 0;
    char *xml_data = zip_read_entry(za, DOCUMENT_XML_PATH, &xml_len, error);
    zip_close(za);
    if (!xml_data) return NULL;

    xmlDocPtr xmldoc = xmlReadMemory(xml_data, (int)xml_len, DOCUMENT_XML_PATH, NULL, XML_PARSE_NOBLANKS);
    g_free(xml_data);
    if (!xmldoc) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "The document.xml inside '%s' could not be parsed as XML", path);
        return NULL;
    }

    xmlXPathContextPtr xpctx = xmlXPathNewContext(xmldoc);
    xmlXPathRegisterNs(xpctx, (const xmlChar *)"w", WORDML_NS);

    xmlXPathObjectPtr paraResult = xmlXPathEvalExpression((const xmlChar *)"//w:p", xpctx);

    DocxDocument *doc = g_new0(DocxDocument, 1);
    doc->src_path = g_strdup(path);
    doc->xmldoc = xmldoc;
    doc->paragraphs = g_array_new(FALSE, TRUE, sizeof(Paragraph));
    g_array_set_clear_func(doc->paragraphs, paragraph_clear);
    doc->field_names = g_ptr_array_new_with_free_func(g_free);

    if (paraResult && paraResult->nodesetval) {
        for (int i = 0; i < paraResult->nodesetval->nodeNr; i++) {
            xmlNodePtr pNode = paraResult->nodesetval->nodeTab[i];

            Paragraph para;
            para.node = pNode;
            para.runs = g_array_new(FALSE, TRUE, sizeof(Run));
            g_array_set_clear_func(para.runs, run_clear);
            para.full_text = g_string_new(NULL);
            para.has_field = FALSE;

            for (xmlNodePtr child = pNode->children; child; child = child->next) {
                if (child->type != XML_ELEMENT_NODE) continue;
                if (xmlStrcmp(child->name, (const xmlChar *)"r") != 0) continue;

                xmlNodePtr rPr = NULL;
                for (xmlNodePtr rc = child->children; rc; rc = rc->next) {
                    if (rc->type == XML_ELEMENT_NODE && xmlStrcmp(rc->name, (const xmlChar *)"rPr") == 0) {
                        rPr = rc;
                        break;
                    }
                }

                Run run;
                run.text = read_run_text(child);
                read_run_props(rPr, &run.bold, &run.italic, &run.underline);
                g_array_append_val(para.runs, run);
                g_string_append(para.full_text, run.text);
            }

            para.has_field = register_fields_in_text(para.full_text->str, doc->field_names);
            g_array_append_val(doc->paragraphs, para);
        }
    }

    if (paraResult) xmlXPathFreeObject(paraResult);
    xmlXPathFreeContext(xpctx);

    return doc;
}

void docx_free(DocxDocument *doc) {
    if (!doc) return;
    g_free(doc->src_path);
    if (doc->xmldoc) xmlFreeDoc(doc->xmldoc);
    if (doc->paragraphs) g_array_free(doc->paragraphs, TRUE);
    if (doc->field_names) g_ptr_array_free(doc->field_names, TRUE);
    g_free(doc);
}

guint docx_field_count(DocxDocument *doc) {
    return doc && doc->field_names ? doc->field_names->len : 0;
}

const char *docx_field_name(DocxDocument *doc, guint index) {
    if (!doc || !doc->field_names || index >= doc->field_names->len) return NULL;
    return (const char *)g_ptr_array_index(doc->field_names, index);
}

/* ---- preview rendering --------------------------------------------------- */

void docx_render_preview(DocxDocument *doc, GtkTextBuffer *buffer, GHashTable *values) {
    gtk_text_buffer_set_text(buffer, "", 0);

    GtkTextTagTable *tags = gtk_text_buffer_get_tag_table(buffer);
    GtkTextTag *bold_tag = gtk_text_tag_table_lookup(tags, "docx-bold");
    if (!bold_tag) bold_tag = gtk_text_buffer_create_tag(buffer, "docx-bold", "weight", PANGO_WEIGHT_BOLD, NULL);
    GtkTextTag *italic_tag = gtk_text_tag_table_lookup(tags, "docx-italic");
    if (!italic_tag) italic_tag = gtk_text_buffer_create_tag(buffer, "docx-italic", "style", PANGO_STYLE_ITALIC, NULL);
    GtkTextTag *underline_tag = gtk_text_tag_table_lookup(tags, "docx-underline");
    if (!underline_tag) underline_tag = gtk_text_buffer_create_tag(buffer, "docx-underline", "underline", PANGO_UNDERLINE_SINGLE, NULL);
    GtkTextTag *filled_tag = gtk_text_tag_table_lookup(tags, "docx-filled");
    if (!filled_tag) filled_tag = gtk_text_buffer_create_tag(buffer, "docx-filled", "foreground", "#1a5fb4", "weight", PANGO_WEIGHT_BOLD, NULL);
    GtkTextTag *placeholder_tag = gtk_text_tag_table_lookup(tags, "docx-placeholder");
    if (!placeholder_tag) placeholder_tag = gtk_text_buffer_create_tag(buffer, "docx-placeholder",
                                                                        "foreground", "#c64600",
                                                                        "background", "#fff3e6",
                                                                        "style", PANGO_STYLE_ITALIC, NULL);

    GtkTextIter end;

    if (!doc) return;

    for (guint pi = 0; pi < doc->paragraphs->len; pi++) {
        Paragraph *para = &g_array_index(doc->paragraphs, Paragraph, pi);

        if (!para->has_field) {
            /* Fast path: render run by run, preserving per-run formatting. */
            for (guint ri = 0; ri < para->runs->len; ri++) {
                Run *r = &g_array_index(para->runs, Run, ri);
                if (!r->text || *r->text == '\0') continue;

                GPtrArray *applied = g_ptr_array_new();
                if (r->bold) g_ptr_array_add(applied, bold_tag);
                if (r->italic) g_ptr_array_add(applied, italic_tag);
                if (r->underline) g_ptr_array_add(applied, underline_tag);

                gtk_text_buffer_get_end_iter(buffer, &end);
                if (applied->len == 0) {
                    gtk_text_buffer_insert(buffer, &end, r->text, -1);
                } else {
                    gint start_offset;
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    start_offset = gtk_text_iter_get_offset(&end);
                    gtk_text_buffer_insert(buffer, &end, r->text, -1);
                    GtkTextIter start_iter;
                    gtk_text_buffer_get_iter_at_offset(buffer, &start_iter, start_offset);
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    for (guint ti = 0; ti < applied->len; ti++) {
                        gtk_text_buffer_apply_tag(buffer, g_ptr_array_index(applied, ti), &start_iter, &end);
                    }
                }
                g_ptr_array_free(applied, TRUE);
            }
        } else {
            /* Paragraph contains one or more {{fields}}: walk the full text,
             * splitting it into literal chunks and field chunks so we can
             * style filled-in values differently from still-empty ones. */
            const char *text = para->full_text->str;
            const char *p = text;
            while (*p) {
                const char *open = strstr(p, "{{");
                if (!open) {
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    gtk_text_buffer_insert(buffer, &end, p, -1);
                    break;
                }
                if (open > p) {
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    gtk_text_buffer_insert(buffer, &end, p, (int)(open - p));
                }
                const char *close = strstr(open + 2, "}}");
                if (!close) {
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    gtk_text_buffer_insert(buffer, &end, open, -1);
                    break;
                }
                char *raw = g_strndup(open + 2, (gsize)(close - (open + 2)));
                char *name = g_strstrip(raw);
                const char *val = values ? (const char *)g_hash_table_lookup(values, name) : NULL;

                gint start_offset;
                gtk_text_buffer_get_end_iter(buffer, &end);
                start_offset = gtk_text_iter_get_offset(&end);

                if (val && *val) {
                    gtk_text_buffer_insert(buffer, &end, val, -1);
                    GtkTextIter s;
                    gtk_text_buffer_get_iter_at_offset(buffer, &s, start_offset);
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    gtk_text_buffer_apply_tag(buffer, filled_tag, &s, &end);
                } else {
                    char *shown = g_strdup_printf("{{%s}}", name);
                    gtk_text_buffer_insert(buffer, &end, shown, -1);
                    GtkTextIter s;
                    gtk_text_buffer_get_iter_at_offset(buffer, &s, start_offset);
                    gtk_text_buffer_get_end_iter(buffer, &end);
                    gtk_text_buffer_apply_tag(buffer, placeholder_tag, &s, &end);
                    g_free(shown);
                }
                g_free(raw);
                p = close + 2;
            }
        }

        gtk_text_buffer_get_end_iter(buffer, &end);
        gtk_text_buffer_insert(buffer, &end, "\n", -1);
    }
}

/* ---- saving --------------------------------------------------------------- */

gboolean docx_save(DocxDocument *doc, const char *out_path, GHashTable *values, GError **error) {
    if (!doc) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED, "No document loaded");
        return FALSE;
    }

    /* For every paragraph that contains a field, collapse its runs into a
     * single run whose text has the placeholders substituted, reusing the
     * formatting (rPr) of the paragraph's first run if it has one. This
     * keeps the fix localized to paragraphs that actually need it. */
    for (guint pi = 0; pi < doc->paragraphs->len; pi++) {
        Paragraph *para = &g_array_index(doc->paragraphs, Paragraph, pi);
        if (!para->has_field) continue;

        char *substituted = substitute_text(para->full_text->str, values);

        /* Grab rPr from the first existing <w:r>, if any, before we start
         * removing nodes. */
        xmlNodePtr templateRPr = NULL;
        xmlNodePtr firstRun = NULL;
        for (xmlNodePtr child = para->node->children; child; child = child->next) {
            if (child->type == XML_ELEMENT_NODE && xmlStrcmp(child->name, (const xmlChar *)"r") == 0) {
                firstRun = child;
                break;
            }
        }
        if (firstRun) {
            for (xmlNodePtr rc = firstRun->children; rc; rc = rc->next) {
                if (rc->type == XML_ELEMENT_NODE && xmlStrcmp(rc->name, (const xmlChar *)"rPr") == 0) {
                    templateRPr = xmlCopyNode(rc, 1);
                    break;
                }
            }
        }

        /* Remove all existing <w:r> children of this paragraph. */
        xmlNodePtr child = para->node->children;
        while (child) {
            xmlNodePtr next = child->next;
            if (child->type == XML_ELEMENT_NODE && xmlStrcmp(child->name, (const xmlChar *)"r") == 0) {
                xmlUnlinkNode(child);
                xmlFreeNode(child);
            }
            child = next;
        }

        /* Build a fresh <w:r><w:rPr>...</w:rPr><w:t xml:space="preserve">...</w:t></w:r> */
        xmlNsPtr wns = xmlSearchNsByHref(doc->xmldoc, para->node, WORDML_NS);
        xmlNodePtr newRun = xmlNewNode(wns, (const xmlChar *)"r");
        if (templateRPr) {
            xmlAddChild(newRun, templateRPr);
        }
        xmlNodePtr tNode = xmlNewTextChild(newRun, wns, (const xmlChar *)"t", NULL);
        xmlNodeSetContent(tNode, (const xmlChar *)"");
        xmlChar *escaped = xmlEncodeSpecialChars(doc->xmldoc, (const xmlChar *)substituted);
        xmlNodeSetContent(tNode, escaped);
        xmlFree(escaped);

        /* xml:space="preserve" so leading/trailing spaces in the value
         * survive round-tripping through Word. The "xml" prefix is the
         * predefined XML namespace, not the wordprocessing one. */
        static const xmlChar *XML_NS_HREF = (const xmlChar *)"http://www.w3.org/XML/1998/namespace";
        xmlNsPtr xmlNs = xmlSearchNsByHref(doc->xmldoc, tNode, XML_NS_HREF);
        if (!xmlNs) {
            xmlNodePtr root = xmlDocGetRootElement(doc->xmldoc);
            xmlNs = xmlNewNs(root, XML_NS_HREF, (const xmlChar *)"xml");
        }
        if (xmlNs) {
            xmlSetNsProp(tNode, xmlNs, (const xmlChar *)"space", (const xmlChar *)"preserve");
        }

        xmlAddChild(para->node, newRun);
        g_free(substituted);
    }

    /* Serialize the modified document.xml */
    xmlChar *xmlbuf = NULL;
    int xmlbuf_len = 0;
    xmlDocDumpFormatMemory(doc->xmldoc, &xmlbuf, &xmlbuf_len, 0);
    if (!xmlbuf) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED, "Failed to serialize the document XML");
        return FALSE;
    }

    /* Copy the original .docx to out_path byte-for-byte, then replace just
     * the word/document.xml entry in the copy. This preserves every other
     * part of the archive (styles, media, headers, footers, ...) exactly. */
    if (g_strcmp0(doc->src_path, out_path) != 0) {
        GError *copy_err = NULL;
        char *contents = NULL;
        gsize len = 0;
        if (!g_file_get_contents(doc->src_path, &contents, &len, &copy_err)) {
            xmlFree(xmlbuf);
            g_propagate_error(error, copy_err);
            return FALSE;
        }
        gboolean ok = g_file_set_contents(out_path, contents, len, &copy_err);
        g_free(contents);
        if (!ok) {
            xmlFree(xmlbuf);
            g_propagate_error(error, copy_err);
            return FALSE;
        }
    }

    int zerr = 0;
    struct zip *za = zip_open(out_path, 0, &zerr);
    if (!za) {
        zip_error_t ze;
        zip_error_init_with_code(&ze, zerr);
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "Could not reopen '%s' to write the filled document: %s", out_path, zip_error_strerror(&ze));
        zip_error_fini(&ze);
        xmlFree(xmlbuf);
        return FALSE;
    }

    zip_source_t *src = zip_source_buffer(za, xmlbuf, (zip_uint64_t)xmlbuf_len, 0);
    if (!src) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED, "Failed to prepare updated document.xml for writing");
        zip_discard(za);
        xmlFree(xmlbuf);
        return FALSE;
    }

    zip_int64_t idx = zip_name_locate(za, DOCUMENT_XML_PATH, 0);
    zip_int64_t result;
    if (idx >= 0) {
        result = zip_file_replace(za, idx, src, 0);
    } else {
        result = zip_file_add(za, DOCUMENT_XML_PATH, src, ZIP_FL_OVERWRITE);
    }

    if (result < 0) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED,
                    "Failed to write updated document.xml into the archive: %s", zip_strerror(za));
        zip_source_free(src);
        zip_discard(za);
        xmlFree(xmlbuf);
        return FALSE;
    }

    if (zip_close(za) != 0) {
        g_set_error(error, G_FILE_ERROR, G_FILE_ERROR_FAILED, "Failed to finalize the .docx archive");
        xmlFree(xmlbuf);
        return FALSE;
    }

    xmlFree(xmlbuf);
    return TRUE;
}
