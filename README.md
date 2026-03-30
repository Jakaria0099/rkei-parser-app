# RKEI Form Processor

A simple Streamlit web app that accepts RKEI `.docx` forms and produces an Excel workbook with the extracted data.

## Sheets in the output

| Sheet | Description |
|---|---|
| `master_rows` | All parsed row-level data |
| `file_level_extraction` | One row per uploaded file |
| `pivot_pathway_uoa` | Pivot summary by pathway / UoA |

## Usage

1. **Open the app** — visit the Streamlit Cloud link (or run locally with `streamlit run app.py`).
2. **Upload** one or more `.docx` RKEI forms using the file uploader.
3. **Click** the **Process** button.
4. **Download** the resulting `final_output.xlsx` file.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

> **Note:** You must also place your `rkei_parser.py` (containing `process_files`) in the repository root.

## Deployment on Streamlit Community Cloud

1. Push this repo (public) to GitHub.
2. Place your `rkei_parser.py` file in the repo root.
3. Go to [share.streamlit.io](https://share.streamlit.io), connect the repo, and deploy.
4. Share the generated URL with your team.

## License

MIT