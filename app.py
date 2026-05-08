from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

COMMON_FORMATS = [
    "pdf",
    "docx",
    "odt",
    "rtf",
    "txt",
    "html",
    "xlsx",
    "ods",
    "csv",
    "pptx",
    "odp",
    "png",
    "jpg",
]

FORMAT_PATTERN = re.compile(r"^[a-z0-9._:+-]+$")


def normalize_target_format(raw_value: str) -> str:
    normalized = raw_value.strip().lower()
    if normalized.startswith("."):
        normalized = normalized[1:]
    if not normalized:
        raise ValueError("Please provide an output format.")
    if not FORMAT_PATTERN.match(normalized):
        raise ValueError(
            "Invalid output format. Use values like pdf, docx, xlsx, txt, html, or pdf:writer_pdf_Export."
        )
    return normalized


def convert_with_soffice(input_file: Path, output_dir: Path, target_format: str) -> Path:
    cmd = [
        "soffice",
        "--headless",
        "--convert-to",
        target_format,
        "--outdir",
        str(output_dir),
        str(input_file),
    ]
    completed = subprocess.run(cmd, capture_output=True, text=True, check=False)

    if completed.returncode != 0:
        details = completed.stderr.strip() or completed.stdout.strip() or "Unknown soffice error."
        raise RuntimeError(f"Conversion failed: {details}")

    converted_files = sorted(
        [p for p in output_dir.iterdir() if p.is_file()],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if not converted_files:
        raise RuntimeError("Conversion did not produce an output file.")
    return converted_files[0]


@app.get("/")
def index():
    return render_template("index.html", common_formats=COMMON_FORMATS, error=None)


@app.post("/convert")
def convert():
    upload = request.files.get("file")
    if upload is None or upload.filename.strip() == "":
        return render_template(
            "index.html",
            common_formats=COMMON_FORMATS,
            error="Please choose a file to convert.",
        ), 400

    try:
        target_format = normalize_target_format(request.form.get("target_format", ""))
    except ValueError as error:
        return render_template(
            "index.html",
            common_formats=COMMON_FORMATS,
            error=str(error),
        ), 400

    filename = secure_filename(upload.filename)
    if not filename:
        filename = "upload.bin"

    working_dir = Path(tempfile.mkdtemp(prefix="soffice-convert-"))
    input_path = working_dir / filename
    output_dir = working_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    upload.save(input_path)

    try:
        converted_path = convert_with_soffice(input_path, output_dir, target_format)
    except RuntimeError as error:
        shutil.rmtree(working_dir, ignore_errors=True)
        return render_template(
            "index.html",
            common_formats=COMMON_FORMATS,
            error=str(error),
        ), 500

    download_name = f"{Path(filename).stem}{converted_path.suffix or ''}"
    response = send_file(
        converted_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/octet-stream",
    )
    response.call_on_close(lambda: shutil.rmtree(working_dir, ignore_errors=True))
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
