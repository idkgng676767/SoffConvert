from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
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

ALLOWED_FORMATS = set(COMMON_FORMATS)
UPLOAD_STEM = "upload"


def normalize_target_format(raw_value: str) -> str:
    normalized = raw_value.strip().lower()
    if normalized.startswith("."):
        normalized = normalized[1:]
    if not normalized:
        raise ValueError("Please choose an output format.")
    if normalized not in ALLOWED_FORMATS:
        raise ValueError("Please choose one of the formats from the dropdown.")
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


def unique_filename(filename: str, used_names: set[str]) -> str:
    if filename not in used_names:
        used_names.add(filename)
        return filename

    path = Path(filename)
    stem = path.stem or UPLOAD_STEM
    suffix = path.suffix
    counter = 2
    while True:
        candidate = f"{stem}-{counter}{suffix}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def render_index(error: str | None = None, selected_format: str = ""):
    normalized_selected = selected_format.strip().lower().lstrip(".")
    if normalized_selected not in ALLOWED_FORMATS:
        normalized_selected = ""
    return render_template(
        "index.html",
        common_formats=COMMON_FORMATS,
        error=error,
        selected_format=normalized_selected,
    )


@app.get("/")
def index():
    return render_index()


@app.post("/convert")
def convert():
    uploads = []
    for upload in request.files.getlist("file"):
        if upload is not None and upload.filename is not None and upload.filename.strip():
            uploads.append(upload)
    selected_format = request.form.get("target_format", "")
    if not uploads:
        return render_index(error="Please choose at least one file to convert.", selected_format=selected_format), 400

    try:
        target_format = normalize_target_format(selected_format)
    except ValueError as error:
        return render_index(error=str(error), selected_format=selected_format), 400

    working_dir = Path(tempfile.mkdtemp(prefix="soffice-convert-"))
    used_input_names: set[str] = set()
    converted_outputs: list[tuple[Path, str]] = []

    for index, upload in enumerate(uploads, start=1):
        raw_filename = upload.filename or ""
        filename = secure_filename(raw_filename)
        if not filename:
            filename = f"{UPLOAD_STEM}-{index}.bin"
        filename = unique_filename(filename, used_input_names)

        input_path = working_dir / filename
        output_dir = working_dir / f"output-{index}"
        output_dir.mkdir(parents=True, exist_ok=True)
        upload.save(input_path)

        try:
            converted_path = convert_with_soffice(input_path, output_dir, target_format)
        except RuntimeError as error:
            shutil.rmtree(working_dir, ignore_errors=True)
            return render_index(error=str(error), selected_format=target_format), 500

        download_name = f"{Path(filename).stem}{converted_path.suffix or ''}"
        converted_outputs.append((converted_path, download_name))

    if len(converted_outputs) == 1:
        converted_path, download_name = converted_outputs[0]
        response = send_file(
            converted_path,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/octet-stream",
        )
        response.call_on_close(lambda: shutil.rmtree(working_dir, ignore_errors=True))
        return response

    archive_path = working_dir / "converted-files.zip"
    used_archive_names: set[str] = set()
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for converted_path, download_name in converted_outputs:
            archive.write(converted_path, arcname=unique_filename(download_name, used_archive_names))

    response = send_file(
        archive_path,
        as_attachment=True,
        download_name="converted-files.zip",
        mimetype="application/zip",
    )
    response.call_on_close(lambda: shutil.rmtree(working_dir, ignore_errors=True))
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB limit
