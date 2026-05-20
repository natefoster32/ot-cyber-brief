# OT Cybersecurity Weekly Brief

A web app that pulls OT cybersecurity news from the past 7 days across themed queries (EU regulation, NERC-CIP / TSA / CISA, ICS security, breaches, competitive landscape) and renders a clean, Banneker-branded weekly brief.

Built for Industrial Defender market intel — designed so a non-technical user can open the URL, click a button, and read the brief.

## How it works

- Hits Google News RSS across ~50 themed queries
- Filters to last 7 days, dedupes by URL and title
- Renders themed sections with hyperlinked stories
- Optional Word doc download

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`).

## Deploy

See `DEPLOY.md` for step-by-step Streamlit Community Cloud deployment.
