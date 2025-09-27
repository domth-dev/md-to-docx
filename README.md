# md-to-docx

Convert **Markdown (.md)** to **DOCX** using an optional **Word template**.  
Supports headings, paragraphs, nested lists that reuse template numbering, GFM tables, task lists, local & URL images, and LaTeX math (OMML with PNG fallback).

---

## Features

- **Template-aware lists**: ordered/bulleted lists reuse your template’s numbering (`numId`) at all nesting levels.
- **Markdown coverage**: headings (H1–H6), paragraphs, **bold/italic/strike**, `inline code`, block quotes, code blocks, horizontal rules.
- **Tables**: GFM tables with cell alignment (left/center/right).
- **Task lists**: `[x]` and `[ ]` checkboxes.
- **Images**:
  - Local images embedded.
  - **URL images** downloaded, validated by content-type and size, then embedded.
- **Math**:
  - LaTeX → **OMML** for common constructs (`^`, `_`, `\frac{a}{b}`, `\sqrt{...}`, `\sum`, `\int`).
  - PNG fallback via Matplotlib when OMML parse isn’t supported.
- **Hyperlinks**: Markdown links become real clickable Word hyperlinks.

---

## Install

```bash
pip install md-to-docx
```

````

> Optional dependency for auto-linking bare URLs:
> `pip install linkify-it-py`
> (Or disable `linkify` in code; see _Configuration_.)

---

## Quick Start

### CLI

```bash
# with entry point
md2docx input.md -t template.docx -o output.docx

# or via module
python -m md_to_docx input.md -t template.docx -o output.docx
```

Options:

- `-t/--template` – Word template (.docx) whose styles/numbering are reused
- `-o/--output` – output file (.docx); default: input name with `.docx`
- `--apply-base-font` – enforce a base font on all styles
- `--font-name` / `--font-size` – base font settings (used only with `--apply-base-font`)

### Python API

```python
from md_to_docx import MarkdownToDocx

md = open("input.md", encoding="utf-8").read()
conv = MarkdownToDocx(template_path="template.docx")
conv.convert_text(md, "output.docx")
```

---

## Markdown Examples

### Math (inline & display)

```md
Inline: The quadratic formula is $x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}$.

Block:

$$
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}
$$
```

### Tables

```md
| Left | Center | Right |
| :--- | :----: | ----: |
| a    |   b    |   123 |
```

### Task List

```md
- [x] Setup
- [ ] Implement
- [ ] Test
```

### Images

```md
![Local](images/logo.png)

![URL](https://example.com/path/to/image.png)
```

---

## Configuration

- **Template lists**: The converter reads your template’s numbering definitions to keep corporate bullets/numbers consistent.
- **Base font override**:

  - `--apply-base-font --font-name "Arial" --font-size 11` (use with care for corporate templates).

- **Auto-linking bare URLs**:

  - Install: `pip install linkify-it-py`
  - Or disable in code by setting `linkify=False` in the Markdown parser (see `parser.py`).

---

## Exit Codes (CLI)

| Code | Meaning                  |
| ---: | ------------------------ |
|    0 | Success                  |
|    2 | Input file not found     |
|    3 | Template error           |
|    4 | Image download error     |
|    5 | Numbering/list error     |
|    6 | Math conversion error    |
|    7 | General md-to-docx error |
|   10 | Unexpected error         |

---

## Dependencies

Lower bounds (aligned with tested environment):

- `python-docx >= 1.1.2`
- `markdown-it-py >= 3.0.0`
- `mdit-py-plugins >= 0.5.0`
- `Pillow >= 10.4.0`
- `matplotlib >= 3.9.2`
- `lxml >= 5.3.0`
- `requests >= 2.32.5`
- _(optional)_ `linkify-it-py` for auto-linking bare URLs

Install all as pinned versions for reproducible dev:

```bash
pip install -r requirements.txt
```

---

## Development

```bash
python -m venv .venv
source .venv/Scripts/activate      # Windows (Git Bash)
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -e .
```

Run CLI from source:

```bash
python -m md_to_docx input.md -t template.docx -o output.docx
```

Build:

```bash
pip install build twine
python -m build
twine check dist/*
```

---

## Troubleshooting

- **“Linkify enabled but not installed.”**
  Install `linkify-it-py` or disable `linkify` in the Markdown parser.

- **Bullet/number styles not applied**
  Ensure your **template** contains proper numbering definitions (`abstractNum` for `bullet` and `decimal`). The converter references existing numbering; it doesn’t invent styles.

- **URL image not embedded**
  The downloader validates `Content-Type` begins with `image/` and enforces a max size. If validation fails, a hyperlink is inserted instead.

- **Math looks like an image**
  For unsupported LaTeX constructs, the PNG fallback is used intentionally to keep the output readable.

---

## License

MIT
````
