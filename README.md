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
