# RFP Compliance Analyzer — Streamlit Demo

Clause-level C / NC / PC compliance analysis for legal and commercial teams. Upload an RFP and a playbook, get an editable compliance matrix, export to Word or Excel.

## Stack

- **Streamlit** — frontend and hosting
- **Anthropic Claude (Sonnet 4.5)** — clause extraction and classification
- **python-docx / openpyxl** — Word and Excel export

## Quick deploy (15 minutes)

See `DEPLOYMENT_GUIDE.md` for the full step-by-step walkthrough.

1. Get an Anthropic API key → [console.anthropic.com](https://console.anthropic.com)
2. Push the code to a GitHub repo
3. Connect the repo to [streamlit.io/cloud](https://streamlit.io/cloud)
4. Add your API key as a secret
5. Deploy. You get a public URL like `elt-rfp-analyzer.streamlit.app`

## Files in this repo

| File | Purpose |
|---|---|
| `streamlit_app.py` | Main application (single file) |
| `requirements.txt` | Python dependencies |
| `secrets.toml.template` | Template for local API key (copy to `.streamlit/secrets.toml`) |
| `.gitignore` | Prevents secrets and junk from being committed |
| `DEPLOYMENT_GUIDE.md` | Full step-by-step setup and deployment instructions |

## Local development

```bash
pip install -r requirements.txt
mkdir -p .streamlit
cp secrets.toml.template .streamlit/secrets.toml
# Edit .streamlit/secrets.toml and paste your real API key
streamlit run streamlit_app.py
```

## Demo limitations (by design)

- **Session-based storage** — closing the browser tab clears results. Sufficient for demoing the workflow; no database needed.
- **No authentication** — anyone with the URL can use the tool. Fine for demos with peers; add auth before any real pilot.
- **API cost per user** — every run costs approx. $0.10–$0.30. Your API key pays for all users. Set billing limits on your Anthropic console.
- **PDFs without text layer won't work** — run scanned PDFs through OCR first (Adobe Acrobat Pro, or free tools like Tesseract).
