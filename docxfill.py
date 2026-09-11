"""
docxfill.py - Tkinter GUI for filling out {{field}} docx templates.

Run with:  python3 docxfill.py [optional/path/to/template.docx]
"""

import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

from docx_engine import DocxTemplate, FIELD_RE
from prefill_store import PrefillStore


class ScrollableFrame(ttk.Frame):
    """A vertically-scrollable frame (Tkinter has no built-in one)."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self._window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self._window_id, width=e.width),
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Mouse wheel scrolling (Linux uses Button-4/5, Windows/mac use MouseWheel).
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class DocxFillApp:
    PLACEHOLDER_FG = "#c64600"
    PLACEHOLDER_BG = "#fff3e6"
    FILLED_FG = "#1a5fb4"

    def __init__(self, root, initial_path=None):
        self.root = root
        self.root.title("DocxFill")
        self.root.geometry("980x640")

        self.template = None
        self.values = {}          # field name -> tk.StringVar
        self.loaded_path = None

        self.prefill = PrefillStore()
        self.prefill_path = None

        self._build_ui()
        self._configure_preview_tags()

        if initial_path:
            self.load_template(initial_path)

    # ---- UI construction -------------------------------------------------

    def _build_ui(self):
        toolbar = ttk.Frame(self.root, padding=(8, 6))
        toolbar.pack(side="top", fill="x")

        ttk.Button(toolbar, text="Open Template…", command=self.on_open).pack(side="left")
        ttk.Button(toolbar, text="Load Prefill…", command=self.on_load_prefill).pack(side="left", padx=(6, 0))

        self.save_button = ttk.Button(toolbar, text="Save As…", command=self.on_save, state="disabled")
        self.save_button.pack(side="right")

        self.subtitle_var = tk.StringVar(value="No template loaded")
        ttk.Label(toolbar, textvariable=self.subtitle_var, foreground="#666666").pack(side="left", padx=12)

        self.prefill_status_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self.prefill_status_var, foreground="#666666").pack(side="left")

        paned = ttk.Panedwindow(self.root, orient="horizontal")
        paned.pack(side="top", fill="both", expand=True)

        # --- left: preview ---
        preview_frame = ttk.Frame(paned)
        self.preview_text = tk.Text(
            preview_frame, wrap="word", state="disabled", padx=12, pady=12,
            borderwidth=1, relief="solid",
        )
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scroll.set)
        self.preview_text.pack(side="left", fill="both", expand=True)
        preview_scroll.pack(side="right", fill="y")

        # --- right: fields panel ---
        fields_outer = ttk.Frame(paned)
        ttk.Label(fields_outer, text="Fields", font=("", 11, "bold")).pack(
            anchor="w", padx=12, pady=(12, 4)
        )
        self.fields_scroll = ScrollableFrame(fields_outer)
        self.fields_scroll.pack(fill="both", expand=True)
        self.fields_box = self.fields_scroll.inner

        self.no_doc_label = ttk.Label(
            self.fields_box,
            text="Open a .docx template with\n{{field}} placeholders to begin.",
            foreground="#888888",
            justify="center",
        )
        self.no_doc_label.pack(pady=24, padx=12)

        paned.add(preview_frame, weight=3)
        paned.add(fields_outer, weight=1)

    def _configure_preview_tags(self):
        base_font = tkfont.nametofont(self.preview_text.cget("font"))
        family, size = base_font.actual("family"), base_font.actual("size")

        self.preview_text.tag_configure("normal", font=(family, size))
        self.preview_text.tag_configure("bold", font=(family, size, "bold"))
        self.preview_text.tag_configure("italic", font=(family, size, "italic"))
        self.preview_text.tag_configure("bold_italic", font=(family, size, "bold italic"))
        self.preview_text.tag_configure("underline", underline=True)
        self.preview_text.tag_configure("filled", foreground=self.FILLED_FG)
        self.preview_text.tag_configure(
            "placeholder", foreground=self.PLACEHOLDER_FG, background=self.PLACEHOLDER_BG
        )

    # ---- fields panel -------------------------------------------------

    def _clear_fields_panel(self):
        for child in self.fields_box.winfo_children():
            child.destroy()

    def _build_fields_panel(self):
        previous_values = self._current_values()

        self._clear_fields_panel()
        self.values = {}

        doc_fields = list(self.template.field_names)
        internal_fields = self.template.internal_fields_for(self.prefill)

        if not doc_fields and not internal_fields:
            ttk.Label(
                self.fields_box,
                text="This template has no {{fields}} to fill in.",
                foreground="#888888",
                wraplength=220,
                justify="left",
            ).pack(anchor="w", padx=12, pady=12)
            return

        for name in doc_fields:
            self._add_field_row(name, self.prefill.options_for(name), previous_values.get(name, ""))

        if internal_fields:
            ttk.Separator(self.fields_box, orient="horizontal").pack(
                fill="x", padx=12, pady=(14, 6)
            )
            ttk.Label(
                self.fields_box,
                text="Additional fields (used inside prefill text)",
                font=("", 8, "italic"),
                foreground="#888888",
                wraplength=220,
                justify="left",
            ).pack(anchor="w", padx=12)
            for name in internal_fields:
                self._add_field_row(name, self.prefill.options_for(name), previous_values.get(name, ""))

    def _add_field_row(self, name, options, initial_value=""):
        row = ttk.Frame(self.fields_box)
        row.pack(fill="x", padx=12, pady=(6, 0))

        ttk.Label(row, text=name, font=("", 9, "bold")).pack(anchor="w")

        var = tk.StringVar(value=initial_value)
        if options:
            widget = ttk.Combobox(row, textvariable=var, values=list(options))
        else:
            widget = ttk.Entry(row, textvariable=var)
        widget.pack(fill="x", pady=(2, 0))

        var.trace_add("write", lambda *_, n=name: self._on_field_changed(n))
        self.values[name] = var

    def _on_field_changed(self, field_name):
        self.refresh_preview()

    def _current_values(self):
        return {name: var.get() for name, var in self.values.items()}

    # ---- preview rendering -------------------------------------------------

    @staticmethod
    def _run_font_tag(run):
        bold = bool(run.bold)
        italic = bool(run.italic)
        if bold and italic:
            return "bold_italic"
        if bold:
            return "bold"
        if italic:
            return "italic"
        return "normal"

    @staticmethod
    def _run_is_underline(run):
        return bool(run.underline)

    def refresh_preview(self):
        text = self.preview_text
        text.configure(state="normal")
        text.delete("1.0", "end")

        if self.template is None:
            text.insert("end", "Open a .docx template to see it here.")
            text.configure(state="disabled")
            return

        values = self._current_values()

        for info in self.template.paragraphs:
            self._render_paragraph(text, info, values)
            text.insert("end", "\n")

        text.configure(state="disabled")

    def _render_paragraph(self, text, info, values):
        """Insert one paragraph's text into the preview, resolving
        {{field}} matches against the full paragraph text (so a
        placeholder split across several runs by Word is still found)
        while still tagging each character with its *own* run's
        formatting (bold/italic/underline)."""
        full = info.full_text
        runs = info.para.runs

        # Map each run to the [start, end) character range it occupies
        # within `full`.
        run_spans = []
        offset = 0
        for run in runs:
            rtext = run.text or ""
            run_spans.append((offset, offset + len(rtext), run))
            offset += len(rtext)

        def run_at(char_pos):
            for start, end, run in run_spans:
                if start <= char_pos < end:
                    return run
            return run_spans[-1][2] if run_spans else None

        def tags_for(run):
            if run is None:
                return ["normal"]
            tags = [self._run_font_tag(run)]
            if self._run_is_underline(run):
                tags.append("underline")
            return tags

        def insert_literal_range(start, end):
            """Insert full[start:end], splitting at run boundaries so each
            slice gets its own run's formatting tags."""
            pos = start
            while pos < end:
                run = run_at(pos)
                run_end = next((e for s, e, r in run_spans if r is run), end)
                chunk_end = min(run_end, end)
                text.insert("end", full[pos:chunk_end], tuple(tags_for(run)))
                pos = chunk_end

        pos = 0
        for m in FIELD_RE.finditer(full):
            if m.start() > pos:
                insert_literal_range(pos, m.start())

            field_run = run_at(m.start())
            base_tags = tags_for(field_run)
            name = m.group(1).strip()
            val = values.get(name, "")
            if val:
                self._insert_resolved_value(text, val, base_tags, values)
            else:
                text.insert("end", "{{%s}}" % name, tuple(base_tags + ["placeholder"]))
            pos = m.end()

        if pos < len(full):
            insert_literal_range(pos, len(full))

    def _insert_resolved_value(self, text, value, base_tags, values):
        """Insert a field's resolved value into the preview, expanding any
        {{nested field}} references it contains exactly one level deep -
        matching save()'s two-pass substitution - and no further (there
        is no third pass, so an unresolved or self-referential nested
        placeholder is inserted as literal text rather than expanded
        again)."""
        pos = 0
        for m in FIELD_RE.finditer(value):
            if m.start() > pos:
                text.insert("end", value[pos:m.start()], tuple(base_tags + ["filled"]))
            name = m.group(1).strip()
            inner_val = values.get(name, "")
            if inner_val:
                text.insert("end", inner_val, tuple(base_tags + ["filled"]))
            else:
                text.insert("end", "{{%s}}" % name, tuple(base_tags + ["placeholder"]))
            pos = m.end()
        if pos < len(value):
            text.insert("end", value[pos:], tuple(base_tags + ["filled"]))

    # ---- open / save -------------------------------------------------

    def load_template(self, path):
        try:
            template = DocxTemplate(path)
        except Exception as exc:  # noqa: BLE001 - surface any load failure to the user
            messagebox.showerror("Couldn't open this file", str(exc))
            return

        self.template = template
        self.loaded_path = path
        self._build_fields_panel()
        self.refresh_preview()

        self.subtitle_var.set(os.path.basename(path))
        self.save_button.configure(state="normal")

    def on_open(self):
        path = filedialog.askopenfilename(
            title="Open Word Template",
            filetypes=[("Word documents", "*.docx")],
        )
        if path:
            self.load_template(path)

    def on_load_prefill(self):
        path = filedialog.askopenfilename(
            title="Load Prefill File",
            filetypes=[
                ("Text/Markdown files", "*.txt *.md"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            self.prefill.load(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Couldn't load prefill file", str(exc))
            return

        self.prefill_path = path
        self.prefill_status_var.set(f"Prefill: {os.path.basename(path)}")

        if self.template is not None:
            self._build_fields_panel()
            self.refresh_preview()

    def on_save(self):
        if self.template is None:
            return

        base = os.path.basename(self.loaded_path) if self.loaded_path else "filled.docx"
        no_ext, _ = os.path.splitext(base)
        suggested = f"{no_ext}-filled.docx" if no_ext else "filled.docx"

        path = filedialog.asksaveasfilename(
            title="Save Filled Document",
            defaultextension=".docx",
            initialfile=suggested,
            filetypes=[("Word documents", "*.docx")],
        )
        if not path:
            return

        try:
            self.template.save(path, self._current_values())
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Couldn't save the document", str(exc))
            return

        messagebox.showinfo("Saved", f"Saved filled document to:\n{path}")


def main():
    initial_path = sys.argv[1] if len(sys.argv) > 1 else None

    root = tk.Tk()
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except tk.TclError:
        pass

    app = DocxFillApp(root, initial_path=initial_path)
    root.mainloop()


if __name__ == "__main__":
    main()
