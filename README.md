# DocxFill (Python / Tkinter)

Same app as the C/GTK version, rewritten in Python: open a `.docx`
template containing `{{field}}` placeholders, see a formatted live
preview, fill in each field in the side panel, and save a new `.docx`
with the placeholders replaced.

## Features

- Opens any `.docx` and parses it with `python-docx`.
- Detects every `{{field}}` placeholder in the document — including
  inside headers and footers — and builds a form field for each one
  automatically.
- Live preview pane renders paragraphs with **bold**, *italic* and
  underline preserved per run. Unfilled placeholders are shown
  highlighted in orange; filled-in values are shown in blue as you type.
  Header/footer fields are still fillable in the side panel but are
  intentionally left out of the preview itself, since headers/footers
  repeat on every page and aren't part of the document's reading flow.
- "Save As…" writes a new `.docx` with the placeholder text replaced
  by your values, everywhere they appear — body, headers, and footers
  alike. Paragraphs with no fields are left completely untouched,
  including their original formatting.
- Optionally load a "prefill" phrasebook file offering a pick-list of
  canned values for one or more fields, so repeated boilerplate text
  doesn't need retyping every time. See **Prefill files** below.

## Making a template

Just type the placeholder text directly in Word (or LibreOffice
Writer, Google Docs, etc.) wherever you want a fillable value:

```
Invoice for {{client_name}}
Amount due: {{amount}}
```

Save/export as `.docx`. Every distinct `{{...}}` becomes a field in
the app's side panel. Field names can contain letters, numbers,
spaces and underscores.

Note: Word sometimes splits a single sentence across multiple
internal "runs" (e.g. due to autocorrect or spell-check), which can
occasionally split a `{{placeholder}}` across two runs so it isn't
detected. If a placeholder doesn't show up as a field, try retyping
that sentence in one go and re-saving.

## Prefill files

Click "Load Prefill…" to load a plain text or Markdown file offering
a pick-list of canned values for one or more fields. Any field the
file has options for gets a dropdown in the side panel instead of a
plain text box — you can pick one of the listed options or still type
your own custom text, same as before.

File format:

```
{{field name}}
[prefill option 1]
[prefill option 2]
[prefill option 3 with {{another field}} inside]

{{another field}}
[option A]
[option B]
```

- A line that's *exactly* `{{field name}}` starts that field's list of
  options; every following `[...]` line is one option, until the next
  `{{...}}` header or end of file.
- An option is everything between the first `[` and the last `]` on
  its line, so brackets can safely appear inside the option text.
- Blank lines, and any line that's neither a `{{field}}` header nor a
  `[option]` line, are ignored — so the file can be annotated with
  ordinary prose or Markdown headings as organizational comments.
- Each option is a single line; multi-line options aren't supported.

**Nested fields and the two-pass fill.** An option's text can itself
contain `{{another field}}`. When you save:

1. **Pass 1** fills every `{{field}}` that's actually in the
   document, using either what you typed or the prefill option you
   picked (which may still contain `{{another field}}` unexpanded at
   this point).
2. **Pass 2** resolves any `{{field}}` that pass 1 just introduced.
   - If that nested field is *also* a real field elsewhere in the
     document, it already has its own entry in the side panel and
     pass 2 reuses that same value.
   - If it's *not* a real document field — it only exists inside
     prefill text — it shows up in the panel under "Additional fields
     (used inside prefill text)" so you can give it a value too.
3. **There is no pass 3.** Anything still shaped like `{{...}}` after
   pass 2 (e.g. a typed value that itself happens to contain `{{...}}`
   with no matching field) is left as literal text in the output.

The live preview shows this same two-pass resolution as you fill in
fields, so what you see is what gets saved.

## Install & run

Requires Python 3.8+. `tkinter` ships with most Python installs; on
some Linux distros it's a separate package.

```bash
# Linux (Debian/Ubuntu) — only needed if `python3 -c "import tkinter"` fails
sudo apt-get install python3-tk

pip install -r requirements.txt
python3 docxfill.py
```

You can also pass a `.docx` path on the command line to open it
immediately on startup:

```bash
python3 docxfill.py path/to/template.docx
```

### Windows

Python from python.org bundles Tkinter already, so no extra install
step is needed there. From a regular command prompt:

```cmd
pip install -r requirements.txt
python docxfill.py
```

To distribute a double-clickable executable instead of requiring
Python on the target machine, package it with PyInstaller:

```cmd
pip install pyinstaller
pyinstaller --onefile --windowed --name DocxFill docxfill.py
```

`--windowed` suppresses the console window (the equivalent of
`-mwindows` in the C/GTK version). The resulting `DocxFill.exe` will
be in the `dist/` folder; it bundles Python, Tkinter and python-docx
so it runs standalone.

## Project layout

```
docxfill.py        - Tkinter UI: window, preview pane, dynamic fields panel, open/save dialogs
docx_engine.py      - .docx engine: load, detect fields, two-pass substitution, save
prefill_store.py    - loads a {{field}} / [option] prefill phrasebook file
requirements.txt    - python-docx dependency
```

## Design notes / limitations

- Rendering fidelity targets template-style documents: paragraph
  breaks and bold/italic/underline runs are rendered. Tables are read
  (their paragraphs are scanned for fields, recursively into nested
  tables) but rendered as plain paragraph text in the preview, after
  the document's direct paragraphs rather than perfectly interleaved
  with them. Images are not shown in the preview but are left
  untouched in the saved output.
- For paragraphs that contain a field, all runs in that paragraph are
  collapsed into a single run on save (using the formatting of the
  paragraph's first run), so that placeholder substitution is safe
  even when Word has split the sentence across several runs.
  Paragraphs with no fields are left completely untouched.
- Field detection, the live preview, and saving all resolve
  `{{field}}` against the paragraph's full text rather than one run
  at a time, so a placeholder that Word has split across multiple
  runs (very common — e.g. `{{CLIENT NAME}}` frequently ends up as
  three separate runs: `{{`, `CLIENT NAME`, `}}`) is still detected,
  previewed, and substituted correctly. The preview additionally maps
  each character back to its originating run so bold/italic/underline
  are still shown correctly even across a split placeholder.
- Header and footer fields are collected from every section's header
  and footer (default, first-page, and even-page variants). A section
  that doesn't define its own header/footer and instead links back to
  the previous section's is only processed once, so its paragraphs
  aren't scanned and substituted multiple times. These fields are
  filled on save exactly like body fields, but are deliberately kept
  out of `DocxTemplate.paragraphs` (the preview's data source) and
  only live in `DocxTemplate.header_footer_paragraphs`.
- A placeholder split across two *paragraphs* (rather than runs
  within the same paragraph) won't be detected, since a paragraph
  break always ends the current field search.

## Zip-bomb protection

A `.docx` is a zip archive, and a template file may come from outside
your organization. Before anything is decompressed, `DocxTemplate`
inspects the archive's central directory (which lists every entry's
compressed and uncompressed size without decompressing anything) and
refuses to open the file if:

- any single part claims to be over 100 MB uncompressed,
- the whole archive would expand past 300 MB uncompressed,
- any part has a compression ratio over 200:1 (a strong zip-bomb
  signature — real Word document parts don't compress anywhere near
  that well), or
- the archive has more than 5000 internal entries (guards against the
  many-tiny-files variant, e.g. the classic "42.zip").

These limits are generous for genuine templates (including ones with
a few embedded images/logos) while catching archives engineered to
exhaust memory or CPU. A rejected file raises `DocxZipSafetyError`
(a subclass of `ValueError`), which the GUI already catches and shows
as an error dialog — no special handling needed by callers beyond the
existing `try`/`except` around `DocxTemplate(path)`.

The limits are plain module-level constants at the top of
`docx_engine.py` (`MAX_UNCOMPRESSED_ENTRY`, `MAX_UNCOMPRESSED_TOTAL`,
`MAX_COMPRESSION_RATIO`, `MAX_ENTRY_COUNT`) if you need to tune them
for unusually large templates.
