# SoffConvert

SoffConvert is a lightweight Flask web app for local file conversion powered by LibreOffice (`soffice`).
Upload a document, choose an output format, and download the converted result from a simple browser UI.

## Key Features

- Local conversion using your own LibreOffice installation
- Web interface for quick one-file-at-a-time conversion
- Supports common output formats (PDF, DOCX, ODT, RTF, TXT, HTML, XLSX, ODS, CSV, PPTX, ODP, PNG, JPG)
- Uses a dropdown menu for output format selection
- Validates output format against supported dropdown values
- Upload size limit set to 200 MB

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

## Usage

1. Select a file to upload.
2. Choose a target output format from the dropdown menu.
3. Click **Convert file**.
4. Download the converted file returned by the app.

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
- This app is intended for local/trusted use; add authentication and stricter controls before internet exposure.

## License

This project is licensed under the MIT License. See `LICENSE`.
