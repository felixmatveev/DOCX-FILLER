"""
prefill_store.py - Loads a "prefill phrasebook" file: a plain text or
Markdown file that offers a pick-list of canned values for one or more
fields, so a user filling out a template can select a common phrase
instead of retyping it every time.

File format:

    {{field name}}
    [prefill option 1]
    [prefill option 2]
    [prefill option 3 with {{another field}} inside]

    {{another field}}
    [option A]
    [option B]

Rules:
- A line that is *exactly* `{{name}}` (after trimming whitespace)
  starts a new field's option block; every following `[...]` line is
  one selectable option for that field, until the next `{{...}}`
  header or end of file.
- A line that is exactly `[...]` is one option. The option text is
  everything between the first `[` and the last `]` on the line, so
  brackets may safely appear inside the option text itself.
- Blank lines are ignored. Lines that are neither a `{{field}}`
  header nor a `[option]` line (stray prose, Markdown headings used
  as organizational comments, etc.) are ignored too, so the file can
  be lightly annotated without breaking parsing.
- An option's text may itself contain `{{another field}}`
  placeholders. These are only ever resolved one level deep at
  save/preview time (see DocxTemplate.save's two-pass substitution) -
  there is no recursive expansion.
- Multi-line options aren't supported: each option is a single line.
"""

import re

FIELD_HEADER_RE = re.compile(r"^\{\{\s*(.+?)\s*\}\}$")
OPTION_RE = re.compile(r"^\[(.*)\]$")


class PrefillStore:
    def __init__(self):
        # field name -> list[str] of raw option strings, in file order.
        # May themselves contain {{other field}} placeholders.
        self.options = {}
        self.path = None

    def load(self, path):
        """(Re)load this store's options from `path`, replacing any
        previously loaded content."""
        options = {}
        current_field = None

        with open(path, encoding="utf-8") as f:
            for raw_line in f:
                line = raw_line.strip()
                if not line:
                    continue

                header_match = FIELD_HEADER_RE.match(line)
                if header_match:
                    current_field = header_match.group(1).strip()
                    options.setdefault(current_field, [])
                    continue

                option_match = OPTION_RE.match(line)
                if option_match and current_field is not None:
                    options[current_field].append(option_match.group(1))
                    continue

                # Anything else (stray prose, headings used as comments,
                # an [option]-shaped line before any {{field}} header) is
                # silently ignored so the file can be lightly annotated.

        self.options = options
        self.path = path
        return self

    def options_for(self, field_name):
        return self.options.get(field_name, [])

    def has_options(self, field_name):
        return bool(self.options.get(field_name))

    def __bool__(self):
        return bool(self.options)
