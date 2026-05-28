from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
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
DEFAULT_ONLYOFFICE_TIMEOUT_SECONDS = 300
DEFAULT_ONLYOFFICE_POLL_INTERVAL_SECONDS = 1
DEFAULT_ONLYOFFICE_TOKEN_TTL_SECONDS = 600
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
ONLYOFFICE_URL = os.getenv("ONLYOFFICE_URL", "").strip()
ONLYOFFICE_PUBLIC_URL = os.getenv("ONLYOFFICE_PUBLIC_URL", "").strip()
ONLYOFFICE_TIMEOUT_SECONDS = parse_nonnegative_int(
    os.getenv("ONLYOFFICE_TIMEOUT_SECONDS"),
    DEFAULT_ONLYOFFICE_TIMEOUT_SECONDS,
)
ONLYOFFICE_POLL_INTERVAL_SECONDS = parse_nonnegative_int(
    os.getenv("ONLYOFFICE_POLL_INTERVAL_SECONDS"),
    DEFAULT_ONLYOFFICE_POLL_INTERVAL_SECONDS,
)
ONLYOFFICE_TOKEN_TTL_SECONDS = parse_nonnegative_int(
    os.getenv("ONLYOFFICE_TOKEN_TTL_SECONDS"),
    max(DEFAULT_ONLYOFFICE_TOKEN_TTL_SECONDS, ONLYOFFICE_TIMEOUT_SECONDS),
)
ONLYOFFICE_JWT_SECRET = os.getenv("ONLYOFFICE_JWT_SECRET", "").strip()
ONLYOFFICE_JWT_HEADER = os.getenv("ONLYOFFICE_JWT_HEADER", "").strip() or "Authorization"
CONVERTER_BACKEND_SETTING = os.getenv("CONVERTER_BACKEND", "").strip().lower()
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
ONLYOFFICE_TIMEOUT = ONLYOFFICE_TIMEOUT_SECONDS if ONLYOFFICE_TIMEOUT_SECONDS > 0 else None
if ONLYOFFICE_POLL_INTERVAL_SECONDS <= 0:
    ONLYOFFICE_POLL_INTERVAL_SECONDS = DEFAULT_ONLYOFFICE_POLL_INTERVAL_SECONDS
if ONLYOFFICE_TOKEN_TTL_SECONDS <= 0:
    ONLYOFFICE_TOKEN_TTL_SECONDS = DEFAULT_ONLYOFFICE_TOKEN_TTL_SECONDS

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
ONLYOFFICE_FILE_TOKENS: dict[str, tuple[Path, float]] = {}
ONLYOFFICE_FILE_LOCK = threading.Lock()


def resolve_converter_backend() -> str:
    if CONVERTER_BACKEND_SETTING in {"soffice", "onlyoffice"}:
        return CONVERTER_BACKEND_SETTING
    if ONLYOFFICE_URL:
        return "onlyoffice"
    return "soffice"


ACTIVE_CONVERTER_BACKEND = resolve_converter_backend()


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
    except zipfile.BadZipFile as error:
        raise ValueError("Zip file is invalid or corrupted.") from error


def cleanup_onlyoffice_tokens(now: float | None = None) -> None:
    if not ONLYOFFICE_FILE_TOKENS:
        return
    if now is None:
        now = time.monotonic()
    expired_tokens = [
        token
        for token, (_, expires_at) in ONLYOFFICE_FILE_TOKENS.items()
        if expires_at <= now
    ]
    for token in expired_tokens:
        ONLYOFFICE_FILE_TOKENS.pop(token, None)


def register_onlyoffice_file(path: Path) -> str:
    token = uuid4().hex
    expires_at = time.monotonic() + ONLYOFFICE_TOKEN_TTL_SECONDS
    with ONLYOFFICE_FILE_LOCK:
        cleanup_onlyoffice_tokens()
        ONLYOFFICE_FILE_TOKENS[token] = (path, expires_at)
    return token


def remove_onlyoffice_file(token: str) -> None:
    with ONLYOFFICE_FILE_LOCK:
        ONLYOFFICE_FILE_TOKENS.pop(token, None)


def get_onlyoffice_file(token: str) -> Path | None:
    now = time.monotonic()
    with ONLYOFFICE_FILE_LOCK:
        cleanup_onlyoffice_tokens(now)
        entry = ONLYOFFICE_FILE_TOKENS.get(token)
        if not entry:
            return None
        path, expires_at = entry
        if expires_at <= now:
            ONLYOFFICE_FILE_TOKENS.pop(token, None)
            return None
    return path


def get_public_base_url() -> str:
    if ONLYOFFICE_PUBLIC_URL:
        return ONLYOFFICE_PUBLIC_URL.rstrip("/")
    return request.url_root.rstrip("/")


def b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")


def create_onlyoffice_jwt(payload: dict) -> str:
    header = {"alg": "HS256", "typ": "JWT"}
    header_segment = b64url_encode(
        json.dumps(header, separators=(",", ":"), sort_keys=True).encode("utf-8")
    )
    payload_segment = b64url_encode(
        json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    )
    signing_input = f"{header_segment}.{payload_segment}".encode("utf-8")
    signature = hmac.new(ONLYOFFICE_JWT_SECRET.encode("utf-8"), signing_input, hashlib.sha256).digest()
    return f"{header_segment}.{payload_segment}.{b64url_encode(signature)}"


def build_onlyoffice_headers(token: str | None = None) -> dict[str, str]:
    headers = {"Content-Type": "application/json"}
    if token:
        header_value = token
        if ONLYOFFICE_JWT_HEADER.lower() == "authorization":
            header_value = f"{'Bearer'} {token}"
        headers[ONLYOFFICE_JWT_HEADER] = header_value
    return headers


def request_onlyoffice_conversion(payload: dict) -> dict:
    endpoint = f"{ONLYOFFICE_URL.rstrip('/')}/ConvertService.ashx"
    request_payload = payload
    token = None
    if ONLYOFFICE_JWT_SECRET:
        request_payload = dict(payload)
        token = create_onlyoffice_jwt(request_payload)
        request_payload["token"] = token
    data = json.dumps(request_payload).encode("utf-8")
    request_obj = urllib.request.Request(
        endpoint,
        data=data,
        headers=build_onlyoffice_headers(token),
        method="POST",
    )
    try:
        with urllib.request.urlopen(request_obj, timeout=ONLYOFFICE_TIMEOUT) as response:
            response_data = response.read()
    except urllib.error.HTTPError as error:
        details = error.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"OnlyOffice conversion request failed ({error.code}): {details or error.reason}"
        ) from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"OnlyOffice conversion request failed: {error.reason}") from error
    try:
        return json.loads(response_data.decode("utf-8"))
    except json.JSONDecodeError as error:
        raise RuntimeError("OnlyOffice returned invalid JSON.") from error


def onlyoffice_conversion_complete(response: dict) -> bool:
    end_convert = response.get("endConvert")
    if isinstance(end_convert, bool):
        return end_convert
    if isinstance(end_convert, str):
        return end_convert.lower() == "true"
    status = response.get("status")
    try:
        return int(status) == 2
    except (TypeError, ValueError):
        return False


def download_onlyoffice_file(file_url: str, output_path: Path) -> None:
    headers: dict[str, str] = {}
    if ONLYOFFICE_JWT_SECRET:
        token = create_onlyoffice_jwt({"url": file_url})
        if ONLYOFFICE_JWT_HEADER.lower() == "authorization":
            headers[ONLYOFFICE_JWT_HEADER] = f"{'Bearer'} {token}"
        else:
            headers[ONLYOFFICE_JWT_HEADER] = token
    request_obj = urllib.request.Request(file_url, headers=headers)
    try:
        with urllib.request.urlopen(request_obj, timeout=ONLYOFFICE_TIMEOUT) as response:
            with output_path.open("wb") as output_file:
                shutil.copyfileobj(response, output_file)
    except urllib.error.HTTPError as error:
        details = error.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"OnlyOffice download failed ({error.code}): {details or error.reason}"
        ) from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"OnlyOffice download failed: {error.reason}") from error


def convert_with_onlyoffice(
    input_file: Path,
    output_dir: Path,
    target_format: str,
    public_base_url: str,
    original_suffix: str,
) -> Path:
    if not ONLYOFFICE_URL:
        raise RuntimeError("OnlyOffice Document Server URL is not configured.")
    if not public_base_url:
        raise RuntimeError("Public base URL is required for OnlyOffice conversion.")
    filetype = (original_suffix or input_file.suffix).lstrip(".").lower()
    if not filetype:
        raise RuntimeError("Unable to determine input file type for OnlyOffice conversion.")
    token = register_onlyoffice_file(input_file)
    try:
        quoted_name = urllib.parse.quote(input_file.name)
        file_url = f"{public_base_url}/internal/onlyoffice/{token}/{quoted_name}"
        payload = {
            "async": False,
            "filetype": filetype,
            "outputtype": target_format,
            "title": input_file.name,
            "url": file_url,
        }
        deadline = None
        if ONLYOFFICE_TIMEOUT_SECONDS > 0:
            deadline = time.monotonic() + ONLYOFFICE_TIMEOUT_SECONDS
        while True:
            response = request_onlyoffice_conversion(payload)
            error_code = response.get("error")
            if error_code not in (None, 0, "0", False):
                raise RuntimeError(f"OnlyOffice conversion failed (error {error_code}).")
            if onlyoffice_conversion_complete(response):
                converted_url = (
                    response.get("fileUrl")
                    or response.get("fileURL")
                    or response.get("file_url")
                    or response.get("url")
                )
                if not converted_url:
                    raise RuntimeError("OnlyOffice conversion did not return a file URL.")
                parsed_path = urllib.parse.urlparse(converted_url).path
                suffix = Path(parsed_path).suffix or f".{target_format}"
                output_path = output_dir / f"{input_file.stem}{suffix}"
                download_onlyoffice_file(converted_url, output_path)
                return output_path
            if deadline and time.monotonic() >= deadline:
                raise RuntimeError(
                    f"OnlyOffice conversion timed out after {ONLYOFFICE_TIMEOUT_SECONDS} seconds."
                )
            time.sleep(ONLYOFFICE_POLL_INTERVAL_SECONDS)
    finally:
        remove_onlyoffice_file(token)


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


@app.get("/internal/onlyoffice/<token>")
@app.get("/internal/onlyoffice/<token>/<filename>")
def onlyoffice_file(token: str, filename: str | None = None):
    path = get_onlyoffice_file(token)
    if not path or not path.exists():
        return "Not found", 404
    return send_file(
        path,
        as_attachment=False,
        download_name=filename or path.name,
        mimetype="application/octet-stream",
    )


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

    backend = ACTIVE_CONVERTER_BACKEND
    public_base_url = ""
    if backend == "onlyoffice":
        public_base_url = get_public_base_url()

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
            storage_filename = f"{UPLOAD_STEM}-{uuid4().hex}{safe_suffix}"

            input_path = working_dir / storage_filename
            output_dir = working_dir / f"output-{index}"
            output_dir.mkdir(parents=True, exist_ok=True)
            upload.save(input_path)

            try:
                validate_zip_payload(input_path, require_zip=original_suffix == ".zip")
            except ValueError as error:
                shutil.rmtree(working_dir, ignore_errors=True)
                return render_index(error=str(error), selected_format=target_format), 400

            try:
                if backend == "onlyoffice":
                    converted_path = convert_with_onlyoffice(
                        input_path,
                        output_dir,
                        target_format,
                        public_base_url,
                        original_suffix,
                    )
                else:
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
