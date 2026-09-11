# DocxFill

A small GTK3 (GNOME) desktop app, written in C, for filling out Word
document templates. Open a `.docx` file that contains `{{field_name}}`
placeholders, see a formatted live preview, fill in each field in the
side panel, and save a new `.docx` with the placeholders replaced.

## Features

- Opens any `.docx` (it's a zip archive of OOXML under the hood) and
  parses `word/document.xml`.
- Detects every `{{field}}` placeholder in the document and builds a
  form field for each one automatically — no manual configuration.
- Live preview pane renders paragraphs with **bold**, *italic* and
  <u>underline</u> preserved. Unfilled placeholders are shown
  highlighted in orange; filled-in values are shown in blue as you type.
- "Save As…" writes a new `.docx`, byte-identical to the original
  except that filled paragraphs have their placeholder text replaced
  with your values. All other parts of the file (styles, images,
  headers/footers, etc.) are copied through untouched.

## Making a template

In Word (or LibreOffice Writer, Google Docs, etc.), just type the
placeholder text directly wherever you want a fillable value, e.g.:

```
Invoice for {{client_name}}
Amount due: {{amount}}
```

Save/export as `.docx`. Every distinct `{{...}}` becomes a field in
the app's side panel. Field names can contain letters, numbers,
spaces and underscores.

Note: Word sometimes splits a single sentence across multiple
internal "runs" (e.g. due to autocorrect or spell-check), which can
occasionally split a `{{placeholder}}` across two runs. If a
placeholder isn't detected, try disabling autocorrect for that
sentence or retyping it in one go, then re-saving.

## Build

### Linux

Dependencies (Debian/Ubuntu package names):

```bash
sudo apt-get install build-essential pkg-config \
    libgtk-3-dev libxml2-dev libzip-dev
```

Then:

```bash
make
./docxfill
```

You can also pass a `.docx` path on the command line to open it
immediately on startup:

```bash
./docxfill path/to/template.docx
```

### Windows (MSYS2)

1. Install MSYS2 from https://www.msys2.org/ and run the installer.
2. Open the **"MSYS2 MinGW64"** terminal (Start menu — not the plain
   MSYS2 terminal).
3. Update packages:
   ```bash
   pacman -Syu
   ```
   (it may ask you to reopen the terminal partway through; do that
   and run `pacman -Syu` again)
4. Install the toolchain and dependencies:
   ```bash
   pacman -S mingw-w64-x86_64-gcc \
             mingw-w64-x86_64-pkg-config \
             mingw-w64-x86_64-make \
             mingw-w64-x86_64-gtk3 \
             mingw-w64-x86_64-libxml2 \
             mingw-w64-x86_64-libzip
   ```
5. Build (from inside the MinGW64 shell, in the project folder):
   ```bash
   mingw32-make
   ```
6. Run:
   ```bash
   ./docxfill.exe
   ```

To hand the `.exe` to someone without MSYS2 installed, bundle its
DLLs alongside it (run from the MinGW64 shell, in the build folder):

```bash
ldd docxfill.exe | grep mingw64 | awk '{print $3}' | xargs -I{} cp {} .
cp -r /mingw64/share/glib-2.0 .
cp -r /mingw64/lib/gdk-pixbuf-2.0 .
mkdir -p share/icons && cp -r /mingw64/share/icons/Adwaita share/icons/
```
Then zip the folder (exe + DLLs + copied share/ folders) as a
portable distribution. Test on a clean Windows machine before
sharing, since exact runtime file needs can vary by GTK3 version.


## Project layout

```
Makefile
src/
  main.c   - GTK UI: window, preview pane, dynamic fields panel, open/save dialogs
  docx.c   - .docx engine: zip+XML parsing, field detection, preview rendering, saving
  docx.h   - public interface to docx.c, documented
```

## Design notes / limitations

- Rendering fidelity targets template-style documents: paragraph
  breaks, bold/italic/underline runs, tabs and line breaks are
  rendered. Tables are read as plain paragraphs (cell text appears in
  reading order); images are not shown in the preview but survive
  unchanged in the saved output. Full WYSIWYG (fonts, colors, tables
  layout, images) is out of scope for a lightweight template filler.
- For paragraphs that contain a field, all runs in that paragraph are
  collapsed into a single run for the save (using the formatting of
  the paragraph's first run) so that placeholder substitution is safe
  even when Word has split the sentence across several runs. Paragraphs
  with no fields are left completely untouched.
