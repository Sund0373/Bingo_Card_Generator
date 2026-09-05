"""
Renders bingo cards straight to a print-ready PDF.

Geometry is taken from the user's 'Bingo Layout.docx':
    landscape 11 x 8.5in, 1in margins, two text columns with a 0.5in gutter,
    5 x 6 grid of 0.8in cells, header shaded 002060 in 48pt Arial Bold,
    FREE SPACE shaded C00000 in 13pt Arial Bold, terms in 14pt Arial black,
    0.5pt black gridlines.

Two cards per page: one at the top of each column.
"""

from pathlib import Path

from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

PAGE_W, PAGE_H = 11 * inch, 8.5 * inch
COLS, ROWS = 5, 5  # grid of terms; the header row is drawn on top of these

# Cards are sized around two straight cuts: one down the middle of the page and one
# across the bottom. Everything is derived from the centre line so the two cards stay
# symmetric about the vertical cut no matter how the constants are tuned.
CELL = 0.96 * inch
MARGIN_TOP = 0.5 * inch
CUT_CLEARANCE = 0.3 * inch      # card edge -> centre cut line
BOTTOM_CLEARANCE = 0.5 * inch   # card bottom -> bottom trim line

CARD_W = COLS * CELL
CARD_H = (ROWS + 1) * CELL      # +1 for the header row
CENTRE_X = PAGE_W / 2
LEFT_X = CENTRE_X - CUT_CLEARANCE - CARD_W
RIGHT_X = CENTRE_X + CUT_CLEARANCE
CARD_TOP = PAGE_H - MARGIN_TOP
CUT_Y = CARD_TOP - CARD_H - BOTTOM_CLEARANCE  # horizontal trim line

HEADER_FILL = HexColor("#002060")
FREE_FILL = HexColor("#C00000")
GUIDE_COLOR = HexColor("#B0B0B0")
GRID_WIDTH = 0.5  # points, matches Word's sz=4 single border

FONT, FONT_BOLD = "Arial", "Arial-Bold"
HEADER_PT = 58
TERM_PT = 18
FREE_PT = 16
MIN_PT = 8  # floor when shrinking a long term to fit its cell
CELL_PAD = 4  # points of horizontal breathing room inside a cell

# Arial lives in a different place on every OS, and may not be installed at all.
# Each entry is a (regular, bold) pair; the first pair where both files exist wins.
FONT_CANDIDATES = [
    # Windows
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    # macOS
    (
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    # Linux - Liberation Sans and DejaVu are metric-compatible stand-ins
    (
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ),
    (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ),
]


def _register_fonts() -> None:
    """Registers Arial if it can be found, otherwise falls back to reportlab's
    built-in Helvetica, which is metrically compatible and always available."""
    global FONT, FONT_BOLD

    if FONT in pdfmetrics.getRegisteredFontNames():
        return

    for regular, bold in FONT_CANDIDATES:
        if Path(regular).is_file() and Path(bold).is_file():
            try:
                pdfmetrics.registerFont(TTFont(FONT, regular))
                pdfmetrics.registerFont(TTFont(FONT_BOLD, bold))
                return
            except Exception:
                continue  # unreadable or unsupported file, try the next candidate

    FONT, FONT_BOLD = "Helvetica", "Helvetica-Bold"


def _wrap(text: str, font: str, size: float, max_width: float) -> list[str]:
    """Greedy word wrap; a single over-long word is left on its own line."""
    words, lines, current = text.split(), [], ""
    for word in words:
        trial = f"{current} {word}".strip()
        if pdfmetrics.stringWidth(trial, font, size) <= max_width or not current:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _fit(text: str, font: str, start_pt: float, max_width: float, max_height: float):
    """Shrinks the font until the wrapped text fits the cell. Returns (lines, size)."""
    size = start_pt
    while size >= MIN_PT:
        lines = _wrap(text, font, size, max_width)
        widest = max((pdfmetrics.stringWidth(l, font, size) for l in lines), default=0)
        if widest <= max_width and len(lines) * size * 1.15 <= max_height:
            return lines, size
        size -= 0.5
    return _wrap(text, font, MIN_PT, max_width), MIN_PT


def _draw_centred(c: canvas.Canvas, lines, font, size, cx, cy) -> None:
    """Draws wrapped lines centred on (cx, cy)."""
    leading = size * 1.15
    total = len(lines) * leading
    y = cy + total / 2 - leading + size * 0.32
    for line in lines:
        c.setFont(font, size)
        c.drawCentredString(cx, y, line)
        y -= leading


def _draw_card(c: canvas.Canvas, card, x, y_top, header, free_space, free_label) -> None:
    """Draws one card with its top-left corner at (x, y_top)."""
    table_w = COLS * CELL

    # --- header row ---
    hdr_y = y_top - CELL
    c.setFillColor(HEADER_FILL)
    c.rect(x, hdr_y, table_w, CELL, stroke=0, fill=1)
    c.setFillColor(white)
    for i, letter in enumerate(header):
        c.setFont(FONT_BOLD, HEADER_PT)
        c.drawCentredString(x + i * CELL + CELL / 2, hdr_y + CELL / 2 - HEADER_PT * 0.34, letter)

    # --- term cells ---
    for r in range(ROWS):
        cell_y = hdr_y - (r + 1) * CELL
        for col in range(COLS):
            cell_x = x + col * CELL
            cx, cy = cell_x + CELL / 2, cell_y + CELL / 2
            value = card[r][col]
            is_free = free_space and value == free_label

            if is_free:
                c.setFillColor(FREE_FILL)
                c.rect(cell_x, cell_y, CELL, CELL, stroke=0, fill=1)
                c.setFillColor(white)
                _draw_centred(c, ["FREE", "SPACE"], FONT_BOLD, FREE_PT, cx, cy)
            else:
                lines, size = _fit(value, FONT, TERM_PT, CELL - 2 * CELL_PAD, CELL - 2 * CELL_PAD)
                c.setFillColor(black)
                _draw_centred(c, lines, FONT, size, cx, cy)

    # --- gridlines over the full 6-row table ---
    c.setStrokeColor(black)
    c.setLineWidth(GRID_WIDTH)
    table_h = (ROWS + 1) * CELL
    for col in range(COLS + 1):
        c.line(x + col * CELL, y_top, x + col * CELL, y_top - table_h)
    for row in range(ROWS + 2):
        c.line(x, y_top - row * CELL, x + table_w, y_top - row * CELL)


def _draw_cut_guides(c: canvas.Canvas) -> None:
    """Dashed guides for the two cuts: down the middle, and across the bottom.
    Both sit exactly on the cut lines, so they disappear once the page is trimmed."""
    c.saveState()
    c.setStrokeColor(GUIDE_COLOR)
    c.setLineWidth(0.5)
    c.setDash(4, 4)
    c.line(CENTRE_X, PAGE_H, CENTRE_X, 0)  # vertical: separates the two cards
    c.line(0, CUT_Y, PAGE_W, CUT_Y)        # horizontal: trims the bottom waste
    c.restoreState()


def write_pdf_cards(
    cards: list[list[list[str]]],
    output_path: Path,
    free_space: bool,
    free_space_label: str,
    header: str = "BINGO",
    cut_guides: bool = True,
) -> None:
    _register_fonts()
    c = canvas.Canvas(str(output_path), pagesize=(PAGE_W, PAGE_H))
    c.setTitle(output_path.stem)

    slots = [LEFT_X, RIGHT_X]
    for idx, card in enumerate(cards):
        if idx % 2 == 0 and cut_guides:
            _draw_cut_guides(c)  # once per page, before the cards
        _draw_card(c, card, slots[idx % 2], CARD_TOP, header, free_space, free_space_label)
        if idx % 2 == 1 and idx < len(cards) - 1:
            c.showPage()
    c.save()
