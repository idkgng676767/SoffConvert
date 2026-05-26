from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import zipfile
from uuid import uuid4
from collections import deque
from pathlib import Path

from flask import Flask, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

app = Flask(__name__)
DEFAULT_MAX_TOTAL_UPLOAD_BYTES = 4 * 1024 * 1024 * 1024
SIZE_UNITS = {
    "B": 1,
    "KB": 1024,
    "MB": 1024**2,
    "GB": 1024**3,
    "TB": 1024**4,
}


def parse_upload_limit(raw_value: str | None, default_value: int) -> int:
    if not raw_value:
        return default_value
    normalized = raw_value.strip().replace("_", "")
    if not normalized:
        return default_value
    match = re.fullmatch(r"(?i)(\d+(?:\.\d+)?)\s*([kmgt]?b)?", normalized)
    if not match:
        return default_value
    amount = float(match.group(1))
    unit = (match.group(2) or "B").upper()
    multiplier = SIZE_UNITS.get(unit)
    if multiplier is None:
        return default_value
    total_bytes = int(amount * multiplier)
    if total_bytes <= 0:
        return default_value
    return total_bytes


def parse_nonnegative_int(raw_value: str | None, default_value: int) -> int:
    if raw_value is None:
        return default_value
    normalized = str(raw_value).strip()
    if not normalized:
        return default_value
    try:
        value = int(normalized)
    except ValueError:
        return default_value
    if value < 0:
        return default_value
    return value


def parse_bool(raw_value: str | None) -> bool:
    if not raw_value:
        return False
    normalized = raw_value.strip().lower()
    if not normalized:
        return False
    return normalized in {"1", "true", "yes", "on"}


def format_bytes_label(total_bytes: int) -> str:
    if total_bytes <= 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(total_bytes)
    index = 0
    while value >= 1024 and index < len(units) - 1:
        value /= 1024
        index += 1
    decimal_places = 0 if value >= 10 or index == 0 else 1
    return f"{value:.{decimal_places}f} {units[index]}"


MAX_TOTAL_UPLOAD_BYTES = parse_upload_limit(
    os.getenv("MAX_FILE_UPLOAD_SIZE"),
    DEFAULT_MAX_TOTAL_UPLOAD_BYTES,
)
MAX_TOTAL_UPLOAD_LABEL = format_bytes_label(MAX_TOTAL_UPLOAD_BYTES)
app.config["MAX_CONTENT_LENGTH"] = MAX_TOTAL_UPLOAD_BYTES

MAX_ZIP_UNCOMPRESSED_BYTES = parse_upload_limit(
    os.getenv("MAX_ZIP_UNCOMPRESSED_SIZE"),
    MAX_TOTAL_UPLOAD_BYTES,
)
MAX_ZIP_UNCOMPRESSED_LABEL = format_bytes_label(MAX_ZIP_UNCOMPRESSED_BYTES)

DEFAULT_MAX_CONCURRENT_CONVERSIONS = 2
DEFAULT_CONVERSION_WAIT_SECONDS = 30
DEFAULT_SOFFICE_TIMEOUT_SECONDS = 300
DEFAULT_RATE_LIMIT_REQUESTS = 30
DEFAULT_RATE_LIMIT_WINDOW_SECONDS = 60

MAX_CONCURRENT_CONVERSIONS = parse_nonnegative_int(
    os.getenv("MAX_CONCURRENT_CONVERSIONS"),
    DEFAULT_MAX_CONCURRENT_CONVERSIONS,
)
MAX_CONVERSION_WAIT_SECONDS = parse_nonnegative_int(
    os.getenv("MAX_CONVERSION_WAIT_SECONDS"),
    DEFAULT_CONVERSION_WAIT_SECONDS,
)
SOFFICE_TIMEOUT_SECONDS = parse_nonnegative_int(
    os.getenv("SOFFICE_TIMEOUT_SECONDS"),
    DEFAULT_SOFFICE_TIMEOUT_SECONDS,
)
RATE_LIMIT_REQUESTS = parse_nonnegative_int(
    os.getenv("RATE_LIMIT_REQUESTS"),
    DEFAULT_RATE_LIMIT_REQUESTS,
)
RATE_LIMIT_WINDOW_SECONDS = parse_nonnegative_int(
    os.getenv("RATE_LIMIT_WINDOW_SECONDS"),
    DEFAULT_RATE_LIMIT_WINDOW_SECONDS,
)
TRUST_PROXY_HEADERS = parse_bool(os.getenv("TRUST_PROXY_HEADERS"))

CONVERSION_SEMAPHORE = (
    threading.BoundedSemaphore(MAX_CONCURRENT_CONVERSIONS)
    if MAX_CONCURRENT_CONVERSIONS > 0
    else None
)
RATE_LIMIT_BUCKETS: dict[str, deque[float]] = {}
RATE_LIMIT_LOCK = threading.Lock()
RATE_LIMIT_CLEANUP_INTERVAL_SECONDS = max(RATE_LIMIT_WINDOW_SECONDS, 1)
RATE_LIMIT_LAST_CLEANUP = 0.0
SOFFICE_TIMEOUT = SOFFICE_TIMEOUT_SECONDS if SOFFICE_TIMEOUT_SECONDS > 0 else None

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

FORMAT_LOOKUP = {fmt: fmt for fmt in COMMON_FORMATS}
ALLOWED_INPUT_SUFFIXES = {f".{fmt}" for fmt in COMMON_FORMATS}
UPLOAD_STEM = "upload"


def normalize_target_format(raw_value: str) -> str:
    normalized = raw_value.strip().lower()
    if normalized.startswith("."):
        normalized = normalized[1:]
    if not normalized:
        raise ValueError("Please choose an output format.")
    safe_format = FORMAT_LOOKUP.get(normalized)
    if not safe_format:
        raise ValueError("Please choose one of the formats from the dropdown.")
    return safe_format


def get_client_ip() -> str:
    if TRUST_PROXY_HEADERS:
        forwarded = request.headers.get("X-Forwarded-For", "")
        if forwarded:
            candidate = forwarded.split(",")[0].strip()
            if candidate:
                return candidate
    return request.remote_addr or "unknown"


def is_rate_limited(client_id: str) -> bool:
    global RATE_LIMIT_LAST_CLEANUP
    if RATE_LIMIT_REQUESTS <= 0 or RATE_LIMIT_WINDOW_SECONDS <= 0:
        return False
    now = time.monotonic()
    window_start = now - RATE_LIMIT_WINDOW_SECONDS
    limited = False
    with RATE_LIMIT_LOCK:
        bucket = RATE_LIMIT_BUCKETS.get(client_id)
        if bucket is None:
            bucket = deque()
            RATE_LIMIT_BUCKETS[client_id] = bucket
        while bucket and bucket[0] < window_start:
            bucket.popleft()
        if len(bucket) >= RATE_LIMIT_REQUESTS:
            limited = True
        else:
            bucket.append(now)
        if not bucket:
            RATE_LIMIT_BUCKETS.pop(client_id, None)
        if now - RATE_LIMIT_LAST_CLEANUP >= RATE_LIMIT_CLEANUP_INTERVAL_SECONDS:
            for bucket_key, entries in list(RATE_LIMIT_BUCKETS.items()):
                while entries and entries[0] < window_start:
                    entries.popleft()
                if not entries:
                    RATE_LIMIT_BUCKETS.pop(bucket_key, None)
            RATE_LIMIT_LAST_CLEANUP = now
    return limited


def acquire_conversion_slot() -> bool:
    if CONVERSION_SEMAPHORE is None:
        return True
    if MAX_CONVERSION_WAIT_SECONDS <= 0:
        return CONVERSION_SEMAPHORE.acquire(blocking=False)
    return CONVERSION_SEMAPHORE.acquire(timeout=MAX_CONVERSION_WAIT_SECONDS)


def validate_zip_payload(path: Path, require_zip: bool = False) -> None:
    is_zip = zipfile.is_zipfile(path)
    if not is_zip:
        if require_zip:
            raise ValueError("Zip file is invalid or corrupted.")
        return
    try:
        with zipfile.ZipFile(path) as archive:
            total_uncompressed = 0
            for info in archive.infolist():
                if info.file_size > MAX_ZIP_UNCOMPRESSED_BYTES:
                    raise ValueError(
                        "Zip file expands beyond the allowed limit "
                        f"({MAX_ZIP_UNCOMPRESSED_LABEL})."
                    )
                if total_uncompressed > MAX_ZIP_UNCOMPRESSED_BYTES - info.file_size:
                    raise ValueError(
                        "Zip file expands beyond the allowed limit "
                        f"({MAX_ZIP_UNCOMPRESSED_LABEL})."
                    )
                total_uncompressed += info.file_size
                if total_uncompressed > MAX_ZIP_UNCOMPRESSED_BYTES:
                    raise ValueError(
                        "Zip file expands beyond the allowed limit "
                        f"({MAX_ZIP_UNCOMPRESSED_LABEL})."
                    )
    except zipfile.BadZipFile as error:
        raise ValueError("Zip file is invalid or corrupted.") from error


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
    try:
        completed = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            shell=False,
            timeout=SOFFICE_TIMEOUT,
        )
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(
            f"Conversion timed out after {SOFFICE_TIMEOUT_SECONDS} seconds."
        ) from error

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
    if normalized_selected not in FORMAT_LOOKUP:
        normalized_selected = ""
    return render_template(
        "index.html",
        common_formats=COMMON_FORMATS,
        error=error,
        selected_format=normalized_selected,
        upload_limit_bytes=MAX_TOTAL_UPLOAD_BYTES,
        upload_limit_label=MAX_TOTAL_UPLOAD_LABEL,
    )


@app.get("/")
def index():
    return render_index()


@app.post("/convert")
def convert():
    uploads = []
    for upload in request.files.getlist("file"):
        if upload and upload.filename and upload.filename.strip():
            uploads.append(upload)
    selected_format = request.form.get("target_format", "")
    if not uploads:
        return render_index(error="Please choose at least one file to convert.", selected_format=selected_format), 400

    try:
        target_format = normalize_target_format(selected_format)
    except ValueError as error:
        return render_index(error=str(error), selected_format=selected_format), 400

    slot_acquired = False
    client_id = get_client_ip()
    if is_rate_limited(client_id):
        return render_index(
            error="Too many requests. Please wait a moment and try again.",
            selected_format=target_format,
        ), 429

    slot_acquired = acquire_conversion_slot()
    if not slot_acquired:
        return render_index(
            error="Too many conversions in progress. Please try again shortly.",
            selected_format=target_format,
        ), 429

    working_dir = Path(tempfile.mkdtemp(prefix="soffice-convert-"))
    used_download_names: set[str] = set()
    converted_outputs: list[tuple[Path, str]] = []

    try:
        for index, upload in enumerate(uploads, start=1):
            raw_filename = upload.filename or ""
            safe_upload_name = secure_filename(raw_filename)
            if not safe_upload_name:
                safe_upload_name = f"{UPLOAD_STEM}-{index}.bin"
            safe_upload_name = unique_filename(safe_upload_name, used_download_names)
            original_suffix = Path(safe_upload_name).suffix.lower()
            safe_suffix = original_suffix
            if safe_suffix not in ALLOWED_INPUT_SUFFIXES:
                safe_suffix = ".bin"
            input_filename = f"{UPLOAD_STEM}-{uuid4().hex}{safe_suffix}"

            input_path = working_dir / input_filename
            output_dir = working_dir / f"output-{index}"
            output_dir.mkdir(parents=True, exist_ok=True)
            upload.save(input_path)

            try:
                validate_zip_payload(input_path, require_zip=original_suffix == ".zip")
            except ValueError as error:
                shutil.rmtree(working_dir, ignore_errors=True)
                return render_index(error=str(error), selected_format=target_format), 400

            try:
                converted_path = convert_with_soffice(input_path, output_dir, target_format)
            except RuntimeError as error:
                shutil.rmtree(working_dir, ignore_errors=True)
                return render_index(error=str(error), selected_format=target_format), 500

            download_name = f"{Path(safe_upload_name).stem}{converted_path.suffix or ''}"
            converted_outputs.append((converted_path, download_name))
    finally:
        if slot_acquired and CONVERSION_SEMAPHORE is not None:
            CONVERSION_SEMAPHORE.release()

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
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_STORED) as archive:
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


@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(_: RequestEntityTooLarge):
    return (
        render_index(error=f"Upload too large. Max total upload size is {MAX_TOTAL_UPLOAD_LABEL}."),
        413,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
