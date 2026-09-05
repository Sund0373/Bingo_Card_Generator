"""
Generates random, unique 5x5 bingo cards from a list of custom terms.

Usage:
    python generate_bingo.py --terms templates/terms_template.xlsx --num-cards 30 --free-space
    python generate_bingo.py --terms my_terms.xlsx --num-cards 10 --output output/my_cards.xlsx

Input terms file must be an .xlsx with either:
    - a "By Column" sheet: headers B/I/N/G/O in row 1, 15 terms under each, OR
    - a "Flat List" sheet: header "Term" in A1, 75 terms below (any order)
See make_template.py to generate a blank starting template.
"""

import argparse
import random
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from pdf_output import write_pdf_cards

COLUMNS = ["B", "I", "N", "G", "O"]
TERMS_PER_COLUMN = 15
FREE_SPACE_LABEL = "FREE SPACE"

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=16)
FREE_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
CELL_FONT = Font(size=11)
THIN_BORDER = Border(*(Side(style="thin", color="999999"),) * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


FREE_ROW, FREE_COL = 2, 2  # centre square of the 5x5 grid


def _assert_unique(terms: list[str], context: str) -> None:
    """Terms must be globally unique so a term can never land twice on one card.
    Terms are already upper-cased on read, so 'PTO' and 'pto' collide here by design."""
    seen: dict[str, str] = {}
    dupes: list[str] = []
    for term in terms:
        key = term.casefold()
        if key in seen:
            dupes.append(term)
        else:
            seen[key] = term
    if dupes:
        raise ValueError(
            f"{context} contains duplicate terms: {sorted(set(dupes))}. "
            "Every term must be unique - remove or rename the repeats."
        )


def _read_column_sheet(ws) -> list[list[str]] | None:
    """Reads the 5x15 layout. Returns None if the sheet is empty."""
    by_col: list[list[str]] = []
    for col_idx in range(1, len(COLUMNS) + 1):
        column: list[str] = []
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col_idx).value
            if val is not None and str(val).strip():
                column.append(str(val).strip().upper())
        by_col.append(column)

    if not any(by_col):
        return None
    if not all(len(c) == TERMS_PER_COLUMN for c in by_col):
        counts = dict(zip(COLUMNS, (len(c) for c in by_col)))
        raise ValueError(
            f"'By Column' sheet must have exactly {TERMS_PER_COLUMN} terms per column. "
            f"Found: {counts}"
        )
    _assert_unique([t for col in by_col for t in col], "'By Column' sheet")
    return by_col


def _read_flat_sheet(ws) -> list[str] | None:
    """Reads the single-column layout. Returns None if the sheet is empty."""
    flat = [
        str(ws.cell(row=r, column=1).value).strip().upper()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=1).value is not None
        and str(ws.cell(row=r, column=1).value).strip()
    ]
    if not flat:
        return None
    expected = len(COLUMNS) * TERMS_PER_COLUMN
    if len(flat) != expected:
        raise ValueError(
            f"'Flat List' sheet must contain exactly {expected} terms. Found: {len(flat)}"
        )
    _assert_unique(flat, "'Flat List' sheet")
    return flat


def read_terms(path: Path, mode: str) -> tuple[str, object]:
    """Loads and validates the term list for an explicitly chosen draw style.

    Returns ('column', [[15 terms] x 5]) for per-column draws, or
    ('random', [75 terms]) for whole-pool draws. The draw style is never inferred,
    so filling in both sheets is fine - the style picks which one is read.
    """
    if mode not in ("column", "random"):
        raise ValueError(f"Draw style must be 'column' or 'random' (got {mode!r}).")
    wb = load_workbook(path, data_only=True)
    by_col = _read_column_sheet(wb["By Column"]) if "By Column" in wb.sheetnames else None
    flat = _read_flat_sheet(wb["Flat List"]) if "Flat List" in wb.sheetnames else None

    if by_col is None and flat is None:
        raise ValueError(
            "No terms found. Fill in either the 'By Column' sheet (15 per column) "
            "or the 'Flat List' sheet (75 terms)."
        )

    # The chosen style picks its own source sheet, so both sheets may be filled.
    if mode == "column":
        if by_col is None:
            raise ValueError(
                "Column draw needs the 'By Column' sheet (15 terms under each of B/I/N/G/O)."
            )
        return "column", by_col

    # random: the Flat List is the natural pool; fall back to flattening the columns
    return "random", flat if flat is not None else [t for col in by_col for t in col]


def generate_card(mode: str, data, free_space: bool) -> list[list[str]]:
    """Returns a 5x5 grid (5 rows of 5 values), drawn without replacement."""
    grid: list[list[str]] = [[""] * 5 for _ in range(5)]

    if mode == "random":
        # one draw of 24 (or 25) from the whole pool, not tied to columns
        needed = 24 if free_space else 25
        picks = iter(random.sample(data, needed))
        for r in range(5):
            for c in range(5):
                if free_space and (r, c) == (FREE_ROW, FREE_COL):
                    grid[r][c] = FREE_SPACE_LABEL
                else:
                    grid[r][c] = next(picks)
    else:
        # five independent draws of 5 (4 for the centre column) from that column's 15
        for c in range(5):
            needed = 4 if (free_space and c == FREE_COL) else 5
            picks = iter(random.sample(data[c], needed))
            for r in range(5):
                if free_space and (r, c) == (FREE_ROW, FREE_COL):
                    grid[r][c] = FREE_SPACE_LABEL
                else:
                    grid[r][c] = next(picks)

    return grid


def _max_unique_cards(mode: str, data, free_space: bool) -> int:
    """Number of distinct grids the draw can produce (permutations, since position matters)."""
    def perms(n: int, k: int) -> int:
        total = 1
        for offset in range(k):
            total *= n - offset
        return total

    if mode == "random":
        return perms(len(data), 24 if free_space else 25)
    total = 1
    for c in range(5):
        total *= perms(len(data[c]), 4 if (free_space and c == FREE_COL) else 5)
    return total


def generate_unique_cards(
    mode: str, data, num_cards: int, free_space: bool
) -> list[list[list[str]]]:
    max_possible = _max_unique_cards(mode, data, free_space)
    if num_cards > max_possible:
        raise ValueError(
            f"Requested {num_cards} cards, but only {max_possible} unique combinations "
            f"are possible from this term list."
        )

    cards: list[list[list[str]]] = []
    seen: set[tuple] = set()
    attempts, max_attempts = 0, num_cards * 200 + 1000
    while len(cards) < num_cards:
        attempts += 1
        if attempts > max_attempts:
            raise RuntimeError("Could not generate enough unique cards; try adding more terms.")
        card = generate_card(mode, data, free_space)
        key = tuple(tuple(row) for row in card)
        if key in seen:
            continue
        seen.add(key)
        cards.append(card)
    return cards


def write_cards(cards: list[list[list[str]]], output_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    for idx, card in enumerate(cards, start=1):
        ws = wb.create_sheet(f"Card {idx}")
        for col_idx, letter in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=letter)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(col_idx)].width = 22
        ws.row_dimensions[1].height = 28

        for r, row_vals in enumerate(card, start=2):
            ws.row_dimensions[r].height = 60
            for c, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.font = CELL_FONT
                cell.alignment = CENTER
                cell.border = THIN_BORDER
                if val == FREE_SPACE_LABEL:
                    cell.fill = FREE_FILL
                    cell.font = Font(size=11, bold=True)

    wb.save(output_path)


def positive_int(value: str) -> int:
    ivalue = int(value)
    if ivalue < 1:
        raise argparse.ArgumentTypeError(f"must be 1 or greater (got {ivalue})")
    return ivalue


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate random unique 5x5 bingo cards.")
    parser.add_argument("--terms", required=True, type=Path, help="Path to filled-in terms .xlsx")
    parser.add_argument(
        "--num-cards", required=True, type=positive_int, help="Number of cards to generate"
    )
    parser.add_argument("--free-space", action="store_true", help="Include a free center space")
    parser.add_argument(
        "--mode",
        choices=["column", "random"],
        required=True,
        help="Draw style: 'column' = 5 independent draws from each column's 15 "
        "(reads the 'By Column' sheet); 'random' = one draw from all 75 "
        "(reads the 'Flat List' sheet).",
    )
    parser.add_argument(
        "--format",
        choices=["pdf", "xlsx"],
        default="pdf",
        nargs="+",
        help="Output format(s) (default: pdf, print-ready two cards per page)",
    )
    parser.add_argument(
        "--header",
        default="BINGO",
        help="The 5 header letters, e.g. CHASE (default: BINGO)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output file path (default: output/bingo_cards.pdf or .xlsx)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Optional random seed for reproducibility")

    args = parser.parse_args()
    if args.header is not None and len(args.header) != len(COLUMNS):
        parser.error(f"--header must be exactly {len(COLUMNS)} characters (got '{args.header}')")
    return args


def main() -> None:
    args = parse_args()
    if args.seed is not None:
        random.seed(args.seed)

    mode, data = read_terms(args.terms, args.mode)
    cards = generate_unique_cards(mode, data, args.num_cards, args.free_space)
    print(f"Draw style: {mode}")

    out_dir = Path(__file__).parent / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    formats = args.format if isinstance(args.format, list) else [args.format]
    single = len(formats) == 1

    for fmt in formats:
        path = args.output if (args.output and single) else out_dir / f"bingo_cards.{fmt}"
        path.parent.mkdir(parents=True, exist_ok=True)

        if fmt == "pdf":
            write_pdf_cards(cards, path, args.free_space, FREE_SPACE_LABEL, args.header)
        else:
            write_cards(cards, path)

        print(f"Generated {len(cards)} unique card(s) -> {path}")


if __name__ == "__main__":
    main()
