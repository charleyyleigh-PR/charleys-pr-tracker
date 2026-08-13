# Charley’s PR Tracker 🎀

A Streamlit app for influencers to discover, score, save and track PR, gifting, ambassador, affiliate, UGC and creator opportunities.

## Features

- Pretty pink creator dashboard
- Creator profile
- Google Custom Search discovery
- Brand/program page scraping
- Email extraction
- Rule-based eligibility scoring
- Favorites and status tracking
- Application tracker
- Manual opportunity entry
- AI application answer generator
- AI PR email generator
- CSV export
- SQLite persistence
- No automatic application submission

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Google search setup

The discovery feature uses the Google Custom Search JSON API.

Set:

```bash
GOOGLE_API_KEY=...
GOOGLE_CSE_ID=...
```

Or enter both values in the app sidebar.

## Optional OpenAI

For AI-generated application answers and outreach emails:

```bash
OPENAI_API_KEY=...
```

The rest of the app works without OpenAI.

## Database

The app automatically creates:

```text
pr_pink.db
```

in the same directory.


## Persistent Google Sheets storage

This version stores the creator profile and opportunity database in a Google Sheet.
Create a Google Sheet, share it with your Google Cloud service-account email as **Editor**, and add these values to Streamlit Community Cloud **Secrets**:

```toml
google_sheet_id = "YOUR_SPREADSHEET_ID"

[gcp_service_account]
type = "service_account"
project_id = "YOUR_PROJECT_ID"
private_key_id = "YOUR_PRIVATE_KEY_ID"
private_key = """-----BEGIN PRIVATE KEY-----
YOUR_PRIVATE_KEY
-----END PRIVATE KEY-----
"""
client_email = "YOUR_SERVICE_ACCOUNT@YOUR_PROJECT.iam.gserviceaccount.com"
client_id = "YOUR_CLIENT_ID"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "YOUR_CLIENT_CERT_URL"
universe_domain = "googleapis.com"
```

The app automatically creates two tabs in that spreadsheet: `Profile` and `Opportunities`.
