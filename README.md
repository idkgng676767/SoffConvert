# SoffConvert

SoffConvert is a lightweight Flask web app for local file conversion powered by LibreOffice (`soffice`).
Upload a document, choose an output format, and download the converted result from a simple browser UI.

## Key Features

- Local conversion using your own LibreOffice installation
- Web interface for quick single or batch conversion
- Supports common output formats (PDF, DOCX, ODT, RTF, TXT, HTML, XLSX, ODS, CSV, PPTX, ODP, PNG, JPG)
- Uses a dropdown menu for output format selection
- Validates output format against supported dropdown values
- Upload size limit defaults to 4 GB total (configurable via `MAX_FILE_UPLOAD_SIZE`)

## Project Structure

- `app.py` — Flask app and conversion logic
- `templates/index.html` — UI template
- `static/styles.css` — Styling
- `requirements.txt` — Python dependencies

## Requirements

- Python 3.9+
- LibreOffice installed and `soffice` available on your `PATH`

Check LibreOffice availability:

```bash
soffice --version
```

## Setup

1. Clone the repository.
2. Create and activate a virtual environment.
3. Install Python dependencies.

```bash
cd SoffConvert
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
cd SoffConvert
python app.py
```

Then open `http://127.0.0.1:5000` (or `http://localhost:5000`) in your browser.

## Configuration

- `MAX_FILE_UPLOAD_SIZE`: total upload limit for all uploaded files. Accepts bytes or a unit suffix
  (`KB`, `MB`, `GB`, `TB`). Example: `MAX_FILE_UPLOAD_SIZE=512MB`.
- `MAX_ZIP_UNCOMPRESSED_SIZE`: total uncompressed byte limit for uploaded zip files. Defaults to
  `MAX_FILE_UPLOAD_SIZE`.
- `MAX_CONCURRENT_CONVERSIONS`: maximum conversions processed at once per app process (default `2`).
  Set to `0` to disable the in-process limit.
- `MAX_CONVERSION_WAIT_SECONDS`: how long to wait for a conversion slot before failing (default `30`).
  Set to `0` to fail fast.
- `SOFFICE_TIMEOUT_SECONDS`: timeout for LibreOffice conversions (default `300`). Set to `0` to disable.
- `RATE_LIMIT_REQUESTS`: requests allowed per IP within the rate-limit window (default `30`).
- `RATE_LIMIT_WINDOW_SECONDS`: rate-limit window duration in seconds (default `60`). Set either value
  to `0` to disable rate limiting.

## Usage

1. Select one or more files to upload (the upload meter shows total usage).
2. Choose a target output format from the dropdown menu.
3. Click **Convert files**.
4. Download the converted file, or a zip when multiple files are uploaded.

### Output format notes

- Supported format values are the options shown in the dropdown.
- Format values are normalized to lowercase, and a leading `.` is removed.
- Invalid or unsupported format values are rejected.

## Error Handling

The app returns clear form errors for:

- Missing file upload
- Missing or invalid target format
- LibreOffice conversion failures
- Missing conversion output file

## Security and Operational Notes

- Uploaded files are written to a temporary directory and cleaned up after response completion.
- Filenames are sanitized before saving.
- Zip uploads are inspected for total uncompressed size before processing to prevent decompression bombs.
- Rate limiting and conversion concurrency limits are enforced in-process; use a reverse proxy for
  shared limits across multiple workers.
- This app is intended for local/trusted use; add authentication and stricter controls before internet exposure.

## License

This project is licensed under the MIT License. See `LICENSE`.
