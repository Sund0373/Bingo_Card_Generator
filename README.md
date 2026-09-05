# Bingo Card Generator

Generates unique 5x5 bingo cards from your own term list and outputs a print-ready
PDF — two cards per page, landscape.

- Upload a term list as Excel, pick how many cards, generate, print.
- Terms are forced to **ALL CAPS** and must be unique.
- Optional free space in the centre square.
- Output as **PDF** (print-ready), **Word**, or **Excel**.

---

## Quick start

Needs **git**, **Python 3.10+**, and **Node 16+**.

```bash
git clone https://github.com/Sund0373/Bingo_Card_Generator.git
cd Bingo_Card_Generator
npm run dev
```

That is it. The first run creates a Python virtual environment, installs
dependencies, starts the server, and opens <http://127.0.0.1:5000>. Later runs skip
straight to starting. The UI handles everything from there.

There is nothing to `npm install` — the launcher uses Node built-ins only.

| Command | Does |
|---|---|
| `npm run dev` | Set up if needed, then start the UI |
| `npm start` | Start without re-checking dependencies |
| `npm run setup` | Install dependencies only, do not start |

Extra arguments pass through to the app:

```bash
npm run dev -- --port 8000     # different port
npm run dev -- --no-browser    # do not open a browser
```

<details>
<summary>No Node on the machine? (Python only)</summary>

Node is only a convenience launcher — the project itself is pure Python.

```bat
python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
python app.py
```

On macOS / Linux use `python3 -m venv .venv && source .venv/bin/activate && ...`.
Every run after that: `.venv\Scripts\activate && python app.py`.

On Windows you can also just double-click **`run.bat`**, which does the same thing
with no terminal at all.
</details>

### If something goes wrong

| Symptom | Fix |
|---|---|
| `npm` not recognised | Install Node from [nodejs.org](https://nodejs.org/), or use the Python-only path above. |
| `Python was not found on PATH` | Install Python from [python.org](https://www.python.org/downloads/) and tick **Add python.exe to PATH**, then reopen the terminal. |
| `pip install` blocked by proxy | `pip install --proxy http://YOUR_PROXY:PORT -r requirements.txt` |
| Port 5000 already in use | `python app.py --port 8000` |
| VS Code says packages are missing | `Ctrl+Shift+P` → *Python: Select Interpreter* → pick the `.venv` entry. |
| Word output unavailable | Optional. PDF and Excel still work; `pip install python-docx` to enable it. |

No admin rights are needed — everything installs into the project's own `.venv`.

---

## Use

1. **Get a term template** — click *Download blank template* in the UI, or run
   `python make_template.py` to write `templates/terms_template.xlsx`.
   `make_template.py` refuses to overwrite an existing template, so it will not
   destroy terms you have already typed in. Use `--force` to overwrite deliberately.

2. **Fill in a sheet** (either one, or both — see below):

   | Sheet | Fill with | Draw behaviour |
   |---|---|---|
   | **By Column** | 15 terms under each of B/I/N/G/O | Each card draws 5 from that column's own 15. A term only ever appears in its assigned column. |
   | **Flat List** | 75 terms, any order | Each card draws 25 from all 75 at once. Terms can land anywhere, redrawn per card. |

   With a free space, the centre column draws 4 instead of 5 (column mode), or the
   card draws 24 instead of 25 (random mode).

   **You always pick the draw style, and it decides which sheet is read** — *Column*
   reads **By Column**, *Random* reads **Flat List**. Nothing is inferred from which
   sheet you filled in, so keeping both sheets populated is perfectly fine.

3. **Generate** — set the card count, header letters, and free space, then click
   *Generate cards*.
4. **Print** — print straight from the preview pane, or download the PDF.

### Term rules

- Converted to ALL CAPS automatically — type them in any case.
- Must be **unique** once upper-cased (`PTO` and `pto` are the same term).
  Duplicates are rejected with an error naming them, because a repeated term could
  otherwise land twice on the same card.
- Short terms work best. Cells are 0.96 inch and terms render at 18pt, so acronyms of
  up to about 6 characters fit on one line. Longer terms wrap, and shrink only if they
  still do not fit.

---

## Command line

The UI is optional — the generator runs standalone:

```bash
python generate_bingo.py --terms templates/terms_template.xlsx --mode column --num-cards 30 --free-space
```

| Flag | Default | Meaning |
|---|---|---|
| `--terms` | *required* | Path to your filled-in `.xlsx` |
| `--num-cards` | *required* | How many cards (1 or more) |
| `--free-space` | off | Free space in the centre square |
| `--mode` | *required* | `column` (reads By Column) or `random` (reads Flat List) |
| `--header` | template's | Exactly 5 characters, e.g. `CHASE` |
| `--format` | `pdf` | Any of `pdf`, `docx`, `xlsx` |
| `--output` | `output/bingo_cards.<ext>` | Output path (single format only) |
| `--seed` | random | Reproduces an identical batch |

```bash
# Word and Excel as well, with a fixed seed
python generate_bingo.py --terms my_terms.xlsx --mode random --num-cards 40 --free-space \
    --header CHASE --format pdf docx xlsx --seed 7
```

---

## Layout and cutting

Each landscape page holds two cards, sized for **two straight cuts**:

1. **Down the middle** (at 5.5 in) — separates the two cards.
2. **Across the bottom** — trims the waste strip.

Dashed guides mark both lines. They sit exactly on the cuts, so they disappear once
the page is trimmed. Finished card: **5.5 x 6.76 in**.

| | Value |
|---|---|
| Cell | 0.96 in |
| Card grid | 4.80 x 5.76 in |
| Page margin (left/right/top) | 0.40 / 0.40 / 0.50 in |
| Card edge to centre cut | 0.30 in each side |
| Card bottom to trim cut | 0.50 in |

All of it derives from the page centre line in `pdf_output.py`, so the two cards stay
symmetric about the vertical cut if you retune `CELL`, `CUT_CLEARANCE`, or
`BOTTOM_CLEARANCE`. Pass `cut_guides=False` to `write_pdf_cards()` to omit the guides.

> **Word output uses different geometry.** `docx_output.py` clones
> `templates/Bingo Layout.docx` as-is (1 in margins, 0.8 in cells, two text columns),
> so `.docx` cards are the original smaller size. The PDF is the print-ready format.

## Project files

| File | Role |
|---|---|
| `package.json` / `scripts/dev.js` | `npm run dev` launcher — venv, deps, start |
| `app.py` | Web UI |
| `generate_bingo.py` | Term loading, validation, card drawing, CLI |
| `pdf_output.py` | PDF renderer |
| `docx_output.py` | Word renderer (clones the template table) |
| `make_template.py` | Writes the blank terms workbook |
| `templates/` | Source assets — committed |
| `output/` | Generated files — git-ignored |

## Notes

- Runs on Flask's development server bound to `127.0.0.1`. That is fine for local
  use; putting it on a network would need a production WSGI server and auth.
- Works on Windows, macOS and Linux. The PDF uses Arial where it is installed and
  falls back to Helvetica, which is metrically identical, so output looks the same
  either way. See `FONT_CANDIDATES` in `pdf_output.py` to add a font path.
- Word output needs `python-docx`. It is optional — without it, PDF and Excel still
  work and the UI hides the Word option.
- Generated files under `output/` are ignored by git. Web runs land in
  `output/_web/` and are swept after 6 hours.

## License

[MIT](LICENSE) — free to use, modify, and redistribute, with attribution and no
warranty.

Dependency licenses are all permissive and compatible: Flask (BSD-3-Clause),
openpyxl (MIT), reportlab (BSD-3-Clause), python-docx (MIT).
