# Voter PDF to Excel web app

A Streamlit web interface around the `279_pdf_to_excel.py` OCR pipeline. The OCR extraction steps are preserved; the script accepts a temporary work-directory setting so hosted runs do not write generated page images into the deployed source tree. This first version supports the same three-column, ten-card-row layout as the supplied voter-list PDF, with configurable page range and an optional expected-record count for a single upload.

## Features

- Upload one PDF or a batch of up to five PDFs (50 MB each; 100 MB total).
- Set the first and last voter-list pages. Last page `0` means the final PDF page.
- Optionally provide the expected record count for a single PDF. A mismatch is reported as a warning; a workbook is still created.
- Preview extracted rows and download individual Excel workbooks, or a ZIP for a batch.
- Temporary uploaded PDFs, rendered pages, OCR crops, and generated work files are kept in a per-job temporary directory outside the deployed source tree and removed after each run. Downloadable workbook bytes remain in the active Streamlit session so the user can download them; use “Clear results from this session” when finished.

## Run locally

Install Poppler and Tesseract with Hindi and English language data for your operating system. Then install Python dependencies and start the app:

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

On Debian/Ubuntu, the system packages listed in `packages.txt` are required. On macOS, install Poppler and Tesseract (including Hindi trained data) with your preferred package manager.

## Deploy free on Streamlit Community Cloud

1. Create a GitHub repository containing these files. Do not add voter PDFs, generated workbooks, or OCR output.
2. In Streamlit Community Cloud, choose **Create app**, select the repository and branch, and set the entrypoint to `app.py`.
3. Deploy. Community Cloud reads `requirements.txt` for Python packages and `packages.txt` for Linux packages such as Poppler and Tesseract.
4. Share the resulting `streamlit.app` URL with intended users.

Free hosted resources are limited and shared. OCR can take several minutes for large PDFs, and the service may be throttled. This app has no sign-in or access control: anyone who can reach its URL can use it. Uploaded files are processed by the hosting service; only deploy it for documents you are authorized to upload. Temporary-file deletion is implemented, but this is not a guarantee against infrastructure-level logging or memory retention.

## Layout limitation

The bundled extractor uses fixed card crop geometry for the current voter-list format. Page numbers and record counts can vary, but materially different page/card layouts need a separate layout configuration or extractor.
