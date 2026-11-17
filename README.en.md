# CRF – Cross-Reference Formatter

<!-- README-I18N:START -->
[Русский](README.md) | **English**
<!-- README-I18N:END -->

<p align="center">
  <img src="https://github.com/user-attachments/assets/8da27901-cf41-43f7-b704-85e4b2fb9422" alt="Иконка CRF" width="128" />
</p>

The tool automates cross-referencing in Word documents. It locates
captions for figures, tables, and equations in a `.docx` file, as well as
in-text mentions of these objects. Based on what it finds, the tool
inserts `SEQ` fields for automatic numbering and `REF` fields for
cross-references. After processing, the file is ready to use: inserting a
new element in the middle of the document causes the numbering to update
automatically.

Both Russian and English documents are supported. The UI is bilingual
as well.

## What it does

- Inserts `SEQ` fields into captions of figures, tables, and equations.
  Numbering is derived from the element's position in the document, so an
  incorrect order in the source is corrected: if "Figure 5" appears before
  "Figure 3", the numbers are re-ordered correctly after processing.
- Creates bookmarks on every numbered element. Bookmark names are
  sequential (`Fig_001`, `Fig_002`, …; `Tab_001`, …; `Eq_001`, …;
  `Src_001`, …), so they appear in ascending order in Word's
  "References → Cross-reference → Bookmark" dialog.
- Finds mentions in the text ("Figure 3", "Table 2.1", "Equation 4",
  "[5]", "[1, 3-5, p. 20]") and replaces them with `REF` fields. For a
  reference to be created, the text must contain the element's number.
- Handles OMML equations, including the ones embedded in tables (the
  classic "equation on the left, number on the right in a cell" layout).
- Detects the bibliography section by its heading ("Список литературы",
  "Библиография", "References", etc.), numbers the sources with fields,
  and creates bookmarks on them.
- Idempotent: if the document already contains `SEQ` fields from a
  previous run, the tool picks up the existing bookmarks instead of
  duplicating them.

## About references in the source text

The key requirement for automatic cross-referencing is that the text
contains mentions with numbers: "Figure 3", "Table 2.1", "[5]", etc.
If such mentions are absent, the tool does not invent them – references
have to be inserted manually in Word.

If the number in the source text did not match the actual position of
the element but instead referred to a non-existent one
(e.g. "see Figure 5" while the figure is actually the third
one), that reference will remain plain text: the tool does not guess what
it should point to. You can replace it with the correct one in Word via
"References → Cross-reference → Bookmark" – bookmarks in that dialog are listed in
ascending order, and the right one takes a couple of seconds to find.

## How to use

1. Launch `CRF.exe` (or `python main.py` from sources).
2. Pick a `.docx`.
3. In the top-right corner, set the document language – `ru` or `en`.
   This determines which captions and headings the tool will look for.
4. Click "Process & Save". A copy with the `_autonumbering` suffix will
   appear next to the original file.
5. Open the copy in Word, press `Ctrl+A` then `F9` – Word recomputes all
   fields and fills in the current numbers.

To check a document without modifying it, click "Check Only (No Save)":
the tool runs through all stages and reports results in the log, but
leaves the original file untouched.

## Running from sources

- Python 3.10 or newer.
- `pip install -r requirements.txt` – installs `python-docx`.

Then run `python main.py`.

## Requirements

- Python 3.10+ to run from sources.
- Windows 7+ for the built `.exe`.

## Feedback

If you come across a document the tool cannot parse, open an Issue and
attach a snippet (at minimum: the section heading and an example
caption). Regular expressions are refined as new formats appear. Pull
requests are welcome too; for larger changes, please discuss them in an
Issue first.

## License

Apache License 2.0. Full text in the [LICENSE](LICENSE) file.
