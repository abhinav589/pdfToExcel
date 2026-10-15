# Voter PDF to Excel web app

A Streamlit web interface around the `279_pdf_to_excel.py` OCR pipeline. The OCR extraction steps are preserved; the script accepts a temporary work-directory setting so hosted runs do not write generated page images into the deployed source tree. This first version supports the same three-column, ten-card-row layout as the supplied voter-list PDF, with configurable page range and an optional expected-record count for a single upload.

## Features

- Upload one PDF or a batch of up to five PDFs (50 MB each; 100 MB total).
- Set the first and last voter-list pages. Last page `0` means the final PDF page.
- Optionally provide the expected record count for a single PDF. A mismatch is reported as a warning; a workbook is still created.
- Preview extracted rows and use the visible download buttons for individual Excel workbooks, or a ZIP for a batch.
- Temporary uploaded PDFs, rendered pages, OCR crops, and generated work files are kept in a per-job temporary directory outside the deployed source tree and removed after each run. Downloadable workbook bytes remain in the active Streamlit session so the user can download them; use “Clear results from this session” when finished.

## Run locally on macOS

1. Unzip the package and open the `voter_ocr_webapp` folder.
2. Double-click `setup_mac.command`. It installs the required Homebrew system packages and Python packages in a local `.venv`. Homebrew is required. The Hindi Tesseract language package is provided by Homebrew's `tesseract-lang` formula.
3. Double-click `run_local.command`. The app opens at `http://127.0.0.1:8501`.

You can also run it from Terminal after setup with `.venv/bin/streamlit run app.py --server.address=127.0.0.1 --server.port=8501`. This binds the app to localhost, so it is not exposed to your LAN or the public internet by default.

## Run in Docker

The included `Dockerfile` installs Poppler and Tesseract's Hindi/English data and starts Streamlit on `$PORT` (default 8501). Build and run locally:

```bash
docker build -t voter-pdf-excel .
docker run --rm -p 8501:8501 voter-pdf-excel
```

Open `http://127.0.0.1:8501`. The same Dockerfile can be used by cloud providers that support Docker web services. Streamlit Community Cloud instead uses `requirements.txt` and `packages.txt` directly.

## Deploy free on Streamlit Community Cloud

1. Create a GitHub repository containing these files. Do not add voter PDFs, generated workbooks, or OCR output.
2. In Streamlit Community Cloud, choose **Create app**, select the repository and branch, and set the entrypoint to `app.py`.
3. Deploy. Community Cloud reads `requirements.txt` for Python packages and `packages.txt` for Linux packages such as Poppler and Tesseract.
4. In **App settings → Sharing**, set the app to private if it should only be available to authorized viewers. Then share the resulting `streamlit.app` URL with those viewers.

Free hosted resources are limited and shared. OCR can take several minutes for large PDFs, and the service may be throttled. This app has no sign-in or access control: anyone who can reach its URL can use it. Uploaded files are processed by the hosting service; only deploy it for documents you are authorized to upload. Temporary-file deletion is implemented, but this is not a guarantee against infrastructure-level logging or memory retention.

## Layout limitation

The bundled extractor uses fixed card crop geometry for the current voter-list format. Page numbers and record counts can vary, but materially different page/card layouts need a separate layout configuration or extractor.
