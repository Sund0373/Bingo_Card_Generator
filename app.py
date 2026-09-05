"""
Local web UI for the bingo card generator.

    python app.py                 ->  http://127.0.0.1:5000
    python app.py --port 8000     ->  a different port
    python app.py --no-browser    ->  do not open a browser window

Flask's template folder is pointed at web/ because the project's own templates/
folder holds the Excel and Word source assets.
"""

import argparse
import math
import shutil
import threading
import time
import traceback
import uuid
import webbrowser
from pathlib import Path

from flask import (
    Flask,
    abort,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from werkzeug.utils import secure_filename

from generate_bingo import (
    COLUMNS,
    FREE_SPACE_LABEL,
    generate_unique_cards,
    read_terms,
    write_cards as write_xlsx_cards,
)
from pdf_output import write_pdf_cards

BASE = Path(__file__).parent
RUNS = BASE / "output" / "_web"
DOCX_TEMPLATE = BASE / "templates" / "Bingo Layout.docx"
TERMS_TEMPLATE = BASE / "templates" / "terms_template.xlsx"

MAX_CARDS = 500
CARDS_PER_PAGE = 2
RUN_TTL_SECONDS = 6 * 60 * 60

app = Flask(__name__, template_folder="web/templates", static_folder="web/static")
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB upload ceiling


def _docx_available() -> bool:
    """python-docx is optional; Word output is hidden when it is missing."""
    try:
        import docx  # noqa: F401
    except ImportError:
        return False
    return True


def _sweep_old_runs() -> None:
    """Deletes run folders older than the TTL so output/_web does not grow forever."""
    if not RUNS.exists():
        return
    cutoff = time.time() - RUN_TTL_SECONDS
    for folder in RUNS.iterdir():
        if folder.is_dir() and folder.stat().st_mtime < cutoff:
            shutil.rmtree(folder, ignore_errors=True)


def _run_dir(token: str) -> Path:
    """Resolves a run folder, rejecting anything that escapes RUNS."""
    folder = (RUNS / token).resolve()
    if folder.parent != RUNS.resolve() or not folder.is_dir():
        abort(404)
    return folder


def _form_defaults() -> dict:
    return {
        "num_cards": "30",
        "free_space": True,
        "header": "CHASE",
        "mode": "column",
        "formats": ["pdf"],
        "seed": "",
    }


def _page(**kwargs):
    kwargs.setdefault("form", _form_defaults())
    return render_template(
        "index.html", max_cards=MAX_CARDS, docx_ok=_docx_available(), **kwargs
    )


@app.route("/", methods=["GET"])
def index():
    return _page()


@app.route("/terms-template")
def terms_template():
    """Hands the user a blank terms workbook to fill in."""
    if not TERMS_TEMPLATE.exists():
        abort(404, "Run make_template.py first to create the terms template.")
    return send_file(TERMS_TEMPLATE, as_attachment=True, download_name="terms_template.xlsx")


@app.route("/generate", methods=["POST"])
def generate():
    _sweep_old_runs()

    form = {
        "num_cards": request.form.get("num_cards", "").strip(),
        "free_space": request.form.get("free_space") == "on",
        "header": request.form.get("header", "").strip().upper(),
        "mode": request.form.get("mode", "column"),
        "formats": request.form.getlist("formats") or ["pdf"],
        "seed": request.form.get("seed", "").strip(),
    }

    def fail(message: str):
        return _page(form=form, error=message), 400

    # --- validate inputs before touching the generator ---
    upload = request.files.get("terms")
    if upload is None or not upload.filename:
        return fail("Choose a terms .xlsx file to upload.")
    if not upload.filename.lower().endswith(".xlsx"):
        return fail("Terms file must be an .xlsx workbook.")

    try:
        num_cards = int(form["num_cards"])
    except ValueError:
        return fail("Number of cards must be a whole number.")
    if not 1 <= num_cards <= MAX_CARDS:
        return fail(f"Number of cards must be between 1 and {MAX_CARDS}.")

    if form["mode"] not in ("column", "random"):
        return fail("Draw style must be Column or Random.")

    if form["header"] and len(form["header"]) != len(COLUMNS):
        return fail(f"Header must be exactly {len(COLUMNS)} characters (e.g. CHASE).")

    seed = None
    if form["seed"]:
        try:
            seed = int(form["seed"])
        except ValueError:
            return fail("Seed must be a whole number, or left blank.")

    if "docx" in form["formats"]:
        if not _docx_available():
            return fail("Word output needs python-docx. Install it with: pip install python-docx")
        if not DOCX_TEMPLATE.exists():
            return fail(f"Word template missing: {DOCX_TEMPLATE.name}")

    token = uuid.uuid4().hex
    folder = RUNS / token
    folder.mkdir(parents=True, exist_ok=True)

    terms_path = folder / secure_filename(upload.filename)
    upload.save(terms_path)

    # --- generate ---
    try:
        import random

        if seed is not None:
            random.seed(seed)

        mode, data = read_terms(terms_path, form["mode"])
        cards = generate_unique_cards(mode, data, num_cards, form["free_space"])

        header = form["header"]
        files = []
        for fmt in ("pdf", "docx", "xlsx"):
            if fmt not in form["formats"]:
                continue
            out = folder / f"bingo_cards.{fmt}"
            if fmt == "pdf":
                write_pdf_cards(
                    cards, out, form["free_space"], FREE_SPACE_LABEL, header or "BINGO"
                )
            elif fmt == "docx":
                from docx_output import write_docx_cards

                write_docx_cards(
                    cards,
                    DOCX_TEMPLATE,
                    out,
                    form["free_space"],
                    FREE_SPACE_LABEL,
                    header or None,
                )
            else:
                write_xlsx_cards(cards, out)
            files.append(
                {
                    "fmt": fmt,
                    "name": out.name,
                    "size_kb": round(out.stat().st_size / 1024, 1),
                    "url": url_for("download", token=token, filename=out.name),
                }
            )
    except ValueError as exc:
        shutil.rmtree(folder, ignore_errors=True)
        return fail(str(exc))
    except Exception:
        traceback.print_exc()
        shutil.rmtree(folder, ignore_errors=True)
        return fail("Unexpected error while generating. See the server console for details.")

    pdf_url = (
        url_for("view", token=token, filename="bingo_cards.pdf")
        if (folder / "bingo_cards.pdf").exists()
        else None
    )

    return _page(
        form=form,
        result={
            "cards": len(cards),
            "mode": mode,
            "pages": math.ceil(len(cards) / CARDS_PER_PAGE),
            "files": files,
            "pdf_url": pdf_url,
        },
    )


@app.route("/download/<token>/<path:filename>")
def download(token, filename):
    return send_from_directory(
        _run_dir(token), secure_filename(filename), as_attachment=True
    )


@app.route("/view/<token>/<path:filename>")
def view(token, filename):
    """Serves the PDF inline so the browser's own viewer can display and print it."""
    return send_from_directory(_run_dir(token), secure_filename(filename))


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the bingo card generator UI.")
    parser.add_argument("--port", type=int, default=5000, help="Port (default 5000)")
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser")
    args = parser.parse_args()

    RUNS.mkdir(parents=True, exist_ok=True)
    url = f"http://127.0.0.1:{args.port}"
    print(f"Bingo Card Generator UI -> {url}   (Ctrl+C to stop)")
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False)


if __name__ == "__main__":
    main()
