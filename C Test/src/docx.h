#ifndef DOCX_H
#define DOCX_H

#include <gtk/gtk.h>

/*
 * docx.h - Minimal .docx template engine.
 *
 * Loads a .docx file (which is a zip archive containing OOXML),
 * parses word/document.xml into a simple paragraph/run model,
 * detects {{field}} style placeholders, can render a formatted
 * preview into a GtkTextBuffer, and can save a new .docx with
 * placeholders substituted by user-supplied values.
 *
 * Formatting fidelity is intentionally limited to what is needed
 * for a template-fill tool: paragraph breaks, bold, italic and
 * underline runs are preserved and rendered. Images, tables layout,
 * headers/footers etc. are left untouched in the saved file but are
 * not specially rendered in the preview.
 */

typedef struct DocxDocument DocxDocument;

/* Load a .docx file. Returns NULL and sets *error on failure. */
DocxDocument *docx_load(const char *path, GError **error);

/* Free all resources associated with a loaded document. */
void docx_free(DocxDocument *doc);

/* Number of distinct {{field}} placeholders found in the document. */
guint docx_field_count(DocxDocument *doc);

/* Name of the field at the given index (0-based), no {{ }} delimiters. */
const char *docx_field_name(DocxDocument *doc, guint index);

/*
 * Render the document into the given (already created, empty or not)
 * GtkTextBuffer. `values` maps field name (char*) -> current value
 * (char*), as typed by the user so far. Fields with no value yet (or
 * an empty value) are shown using their {{field}} placeholder text
 * styled distinctly so the user can see what still needs filling in.
 * Safe to call repeatedly (e.g. on every keystroke) to refresh a live
 * preview; it clears the buffer first.
 */
void docx_render_preview(DocxDocument *doc, GtkTextBuffer *buffer, GHashTable *values);

/*
 * Save a new .docx file at out_path, identical to the original except
 * that every {{field}} placeholder is replaced with the corresponding
 * value from `values` (missing fields are replaced with an empty
 * string). Returns TRUE on success, FALSE and sets *error on failure.
 */
gboolean docx_save(DocxDocument *doc, const char *out_path, GHashTable *values, GError **error);

#endif /* DOCX_H */
