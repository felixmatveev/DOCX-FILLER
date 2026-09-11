"""
docx_engine.py - Minimal .docx template engine built on python-docx.

Loads a .docx file, walks its paragraphs (including inside tables, and
inside headers/footers), detects {{field}} style placeholders, and can
substitute them with user-supplied values when saving a new .docx.

A .docx is a zip archive, and a template may come from someone outside
your organization (a client, a subcontractor). Before any decompression
happens, the archive's declared sizes and compression ratios are
checked against sane limits for what a Word document part should look
like, to guard against zip bombs - small archives engineered to expand
to an unreasonable size and exhaust memory/CPU/disk on whoever opens
them. See DocxZipSafetyError and _check_zip_safety() below.

Header/footer fields are detected and filled on save just like body
fields, but are intentionally excluded from the live preview (headers
and footers repeat on every page in Word and aren't part of the
"reading flow" a template-fill preview needs to show).

Formatting fidelity is intentionally limited to what's useful for a
template-fill tool: bold, italic and underline are read per run for
preview rendering. Paragraphs that contain a field are collapsed to a
single run on save (reusing the first run's formatting), since Word
often splits a single sentence's text across several runs and naively
substituting inside one run can silently miss a placeholder that spans
a run boundary.
"""

import re
import zipfile

from docx import Document

FIELD_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# --- zip-bomb safety limits -------------------------------------------------
#
# A legitimate .docx's individual parts (word/document.xml, headers,
# footers, styles.xml, etc.) are typically tens to a few hundred KB even
# for a long document; embedded images are the main thing that can
# legitimately push an archive larger. These limits are generous enough
# not to false-positive on real templates (with or without a few
# embedded images/logos) while still catching archives engineered to
# expand far beyond anything a Word document part should be.

MAX_UNCOMPRESSED_ENTRY = 100 * 1024 * 1024   # 100 MB for any single entry
MAX_UNCOMPRESSED_TOTAL = 300 * 1024 * 1024   # 300 MB for the whole archive
MAX_COMPRESSION_RATIO = 200                   # uncompressed / compressed
MAX_ENTRY_COUNT = 5000                        # guards against many-small-files bombs


class DocxZipSafetyError(ValueError):
    """Raised when a .docx archive's declared sizes/ratios look like a
    zip bomb rather than a plausible Word document, before any part of
    it is decompressed."""


def _check_zip_safety(path):
    """Inspect a .docx's zip central directory - which lists each entry's
    compressed and uncompressed size without decompressing anything -
    and reject the archive if it looks engineered to expand to an
    unreasonable size. Raises DocxZipSafetyError on a suspicious
    archive, or DocxZipSafetyError wrapping the original error if the
    file isn't a readable zip at all."""
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()

            if len(infos) > MAX_ENTRY_COUNT:
                raise DocxZipSafetyError(
                    f"This file has an implausible number of internal parts "
                    f"({len(infos)}) for a .docx and was not opened."
                )

            total_uncompressed = 0
            for info in infos:
                total_uncompressed += info.file_size

                if info.file_size > MAX_UNCOMPRESSED_ENTRY:
                    raise DocxZipSafetyError(
                        f"'{info.filename}' inside this file claims to be "
                        f"{info.file_size / (1024*1024):.0f} MB uncompressed, "
                        f"which is implausible for a .docx part. Refusing to open it."
                    )

                if info.compress_size > 0:
                    ratio = info.file_size / info.compress_size
                    if ratio > MAX_COMPRESSION_RATIO:
                        raise DocxZipSafetyError(
                            f"'{info.filename}' inside this file has a suspicious "
                            f"compression ratio ({ratio:.0f}:1), which is a common "
                            f"signature of a zip bomb. Refusing to open it."
                        )

            if total_uncompressed > MAX_UNCOMPRESSED_TOTAL:
                raise DocxZipSafetyError(
                    f"This file would expand to "
                    f"{total_uncompressed / (1024*1024):.0f} MB uncompressed, "
                    f"which is implausible for a .docx. Refusing to open it."
                )
    except DocxZipSafetyError:
        raise
    except zipfile.BadZipFile as exc:
        raise DocxZipSafetyError(
            "This file isn't a valid .docx (not a readable zip archive)."
        ) from exc


def _iter_block_paragraphs(parent):
    """Yield every paragraph under `parent` (a Document, a table cell, or
    a header/footer), including paragraphs nested inside tables
    (recursively, for nested tables). Order matches python-docx's own
    top-to-bottom paragraph order within each container, but paragraphs
    inside tables are visited after the container's direct paragraphs
    rather than being perfectly interleaved with surrounding text."""
    for para in parent.paragraphs:
        yield para
    for table in getattr(parent, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_block_paragraphs(cell)


def _iter_header_footer_paragraphs(document):
    """Yield every paragraph inside every section's headers and footers
    (default, first-page, and even-page variants), including inside any
    tables they contain. Sections that link a header/footer back to the
    previous section (the common case when you haven't customized
    per-section headers) share the same underlying part, so each
    distinct part is only visited once to avoid processing - and
    double-substituting - the same paragraphs multiple times."""
    seen_part_ids = set()
    for section in document.sections:
        for attr in (
            "header", "footer",
            "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ):
            hf = getattr(section, attr, None)
            if hf is None:
                continue
            if getattr(hf, "is_linked_to_previous", False):
                continue
            part_id = id(hf.part)
            if part_id in seen_part_ids:
                continue
            seen_part_ids.add(part_id)
            yield from _iter_block_paragraphs(hf)


class _ParaInfo:
    """Cached info about one paragraph: the python-docx Paragraph object,
    its concatenated run text, and whether it contains a field."""

    __slots__ = ("para", "full_text", "has_field")

    def __init__(self, para):
        self.para = para
        self.full_text = "".join(run.text or "" for run in para.runs)
        self.has_field = bool(FIELD_RE.search(self.full_text))


class DocxTemplate:
    def __init__(self, path):
        self.path = path

        _check_zip_safety(path)

        self.document = Document(path)

        # Body paragraphs (incl. tables): used for both the live preview
        # and saving.
        self.paragraphs = [_ParaInfo(p) for p in _iter_block_paragraphs(self.document)]

        # Header/footer paragraphs: fields inside these are detected and
        # filled on save, but deliberately left out of self.paragraphs so
        # the preview pane doesn't show repeating header/footer content.
        self.header_footer_paragraphs = [
            _ParaInfo(p) for p in _iter_header_footer_paragraphs(self.document)
        ]

        self.field_names = []
        seen = set()
        for info in self.paragraphs + self.header_footer_paragraphs:
            if not info.has_field:
                continue
            for m in FIELD_RE.finditer(info.full_text):
                name = m.group(1).strip()
                if name and name not in seen:
                    seen.add(name)
                    self.field_names.append(name)

    @staticmethod
    def substitute(text, values):
        """Replace every {{field}} in `text` with values.get(field, '')."""

        def repl(m):
            name = m.group(1).strip()
            val = values.get(name, "")
            return val if val is not None else ""

        return FIELD_RE.sub(repl, text)

    def internal_fields_for(self, prefill_store):
        """Given a PrefillStore, return the field names that appear
        nested inside this document's own fields' prefill options but
        aren't themselves fields already present in the document -
        i.e. fields that only come into existence via a pass-1
        substitution and need their own value collected for pass 2.
        Order follows first appearance; a field already in
        self.field_names is never included here even if it's also
        referenced inside a prefill option, since it already has its
        own entry."""
        known = set(self.field_names)
        internal = []
        seen = set()
        for name in self.field_names:
            for option in prefill_store.options_for(name):
                for m in FIELD_RE.finditer(option):
                    iname = m.group(1).strip()
                    if iname and iname not in known and iname not in seen:
                        seen.add(iname)
                        internal.append(iname)
        return internal

    def save(self, out_path, values):
        """Write a new .docx at out_path with {{field}} placeholders
        replaced, in the body as well as headers/footers. Substitution
        runs in exactly two passes:

        Pass 1 replaces every {{field}} that already exists in the
        document with values.get(field, ''). A value may itself be a
        prefill phrase that contains further {{other field}}
        references (see prefill_store.py) - those are inserted as
        literal text in this pass, not yet resolved.

        Pass 2 re-scans only the paragraphs touched in pass 1 for any
        {{field}} that pass 1's substitution introduced, and resolves
        those against the same `values` mapping. This lets a prefill
        option's nested reference either reuse the value already
        collected for a same-named document field, or a value the
        caller collected separately for a field that only exists
        inside prefill text (an "internal" field never present in the
        original document).

        There is no third pass: whatever {{...}} remains after pass 2
        (e.g. because a value itself literally contained a {{...}}
        with no corresponding entry in `values`) is left as-is in the
        output.

        Paragraphs with no fields at all are left exactly as they
        were, including all their original runs/formatting."""
        touched = []
        for info in self.paragraphs + self.header_footer_paragraphs:
            if not info.has_field:
                continue

            new_text = self.substitute(info.full_text, values)
            runs = info.para.runs
            if not runs:
                continue

            # Keep the first run (and its formatting) as the carrier for
            # the substituted text; drop every other run in the paragraph.
            first_run = runs[0]
            first_run.text = new_text
            for extra_run in runs[1:]:
                element = extra_run._element
                element.getparent().remove(element)

            touched.append(first_run)

        # Pass 2: resolve any {{field}} that pass 1's substitution just
        # introduced. Only paragraphs pass 1 actually touched can
        # possibly contain new placeholders, so this only re-scans those.
        for run in touched:
            if FIELD_RE.search(run.text or ""):
                run.text = self.substitute(run.text, values)

        self.document.save(out_path)
