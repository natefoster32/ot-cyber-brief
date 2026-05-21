# Banneker News Tracker

Multi-tenant Streamlit app where any Banneker coworker can spin up a personalized news brief for their portco, deal, or vertical in ~4 minutes. Optional weekly/daily email delivery.

Built on Google News RSS + Banneker visual style. All free to operate.

## Architecture

- **app.py** — Streamlit UI (home / create / view routing via query params)
- **core.py** — scraping, news filter, ranking, DOCX export, email HTML builder
- **storage.py** — GitHub Gist as the config store (`trackers.json` file)
- **templates.py** — starter templates per portco type
- **email_sender.py** — Resend API wrapper
- **scripts/send_emails.py** — daily cron worker (run by GitHub Actions)
- **.github/workflows/send_emails.yml** — schedule

## Run locally

```bash
pip install -r requirements.txt
# Create .streamlit/secrets.toml with GITHUB_GIST_ID, GITHUB_PAT,
# RESEND_API_KEY, RESEND_FROM
streamlit run app.py
```

## Deploy

See [DEPLOY.md](DEPLOY.md). Setup takes ~25 minutes one-time.
