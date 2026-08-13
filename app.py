
import os
import re
import json
from datetime import datetime, date
from urllib.parse import urlparse, urljoin

import pandas as pd
import gspread
import requests
import streamlit as st
from bs4 import BeautifulSoup

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


APP_TITLE = "Charley’s PR Tracker ✨"

PINK = "#ff5fa2"
LIGHT_PINK = "#fff1f7"
SOFT_PINK = "#ffd6e7"
HOT_PINK = "#ff3f92"
PURPLE = "#8b5cf6"
TEXT = "#3d2b36"
MUTED = "#7c6672"

OPPORTUNITY_TYPES = [
    "PR / Gifting",
    "Ambassador",
    "Creator Program",
    "Affiliate",
    "UGC",
    "Paid Collaboration",
    "Influencer Program",
    "Partnership",
    "Unknown",
]

STATUSES = [
    "Not Applied",
    "Want to Apply",
    "Applied",
    "Follow-up Due",
    "Responded",
    "Accepted",
    "Rejected",
    "PR Received",
    "Collaboration",
    "Not Interested",
]

CATEGORIES = [
    "Beauty",
    "Makeup",
    "Skincare",
    "Hair",
    "Fashion",
    "Wellness",
    "Jewellery",
    "Lifestyle",
    "Other",
]

REGIONS = ["UK", "US", "Europe", "Worldwide", "Unknown"]


# -----------------------------
# Google Sheets storage
# -----------------------------
PROFILE_HEADERS = [
    "name", "email", "country", "niche", "instagram_url",
    "instagram_followers", "tiktok_url", "tiktok_followers",
    "youtube_url", "youtube_followers", "average_views",
    "engagement_rate", "audience", "creator_bio", "media_kit_url"
]

OPPORTUNITY_HEADERS = [
    "id", "brand", "category", "opportunity_type", "program_name",
    "application_url", "contact_email", "brand_website", "instagram",
    "region", "min_followers", "requirements", "compensation",
    "eligibility", "match_score", "score_reason", "status", "favorite",
    "date_found", "date_applied", "follow_up_date", "contact_person",
    "response", "products_received", "deliverables", "content_due_date",
    "payment", "notes", "source_query", "unique_key"
]


def storage_configured():
    try:
        return bool(st.secrets.get("google_sheet_id")) and "gcp_service_account" in st.secrets
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def sheets_client():
    if not storage_configured():
        return None
    creds = dict(st.secrets["gcp_service_account"])
    return gspread.service_account_from_dict(creds)


@st.cache_resource(show_spinner=False)
def spreadsheet():
    client = sheets_client()
    if client is None:
        return None
    return client.open_by_key(st.secrets["google_sheet_id"])


def get_or_create_worksheet(title, headers):
    ss = spreadsheet()
    if ss is None:
        return None
    try:
        ws = ss.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows=1000, cols=max(20, len(headers)))
    existing = ws.row_values(1)
    if not existing:
        ws.update(range_name="A1", values=[headers])
    elif existing != headers:
        # Preserve existing data while ensuring any newly added columns exist.
        merged = existing + [h for h in headers if h not in existing]
        ws.update(range_name="A1", values=[merged])
    return ws


def init_db():
    if not storage_configured():
        return
    get_or_create_worksheet("Profile", PROFILE_HEADERS)
    get_or_create_worksheet("Opportunities", OPPORTUNITY_HEADERS)


def load_profile():
    if not storage_configured():
        return {}
    try:
        ws = get_or_create_worksheet("Profile", PROFILE_HEADERS)
        rows = ws.get_all_records()
        if not rows:
            return {}
        p = rows[0]
        for key in ["instagram_followers", "tiktok_followers", "youtube_followers", "average_views"]:
            try:
                p[key] = int(float(p.get(key) or 0))
            except Exception:
                p[key] = 0
        try:
            p["engagement_rate"] = float(p.get("engagement_rate") or 0)
        except Exception:
            p["engagement_rate"] = 0.0
        return p
    except Exception as e:
        st.error(f"Google Sheets profile read failed: {e}")
        return {}


def save_profile(data):
    if not storage_configured():
        raise RuntimeError("Google Sheets is not connected yet.")
    ws = get_or_create_worksheet("Profile", PROFILE_HEADERS)
    values = [data.get(h, "") for h in PROFILE_HEADERS]
    if ws.row_count < 2:
        ws.add_rows(1)
    ws.update(range_name=f"A2:{gspread.utils.rowcol_to_a1(2, len(PROFILE_HEADERS)).split('2')[0]}2", values=[values])


def make_unique_key(brand, application_url, contact_email, brand_website):
    base = "|".join([
        (brand or "").strip().lower(),
        (application_url or "").strip().lower(),
        (contact_email or "").strip().lower(),
        (brand_website or "").strip().lower(),
    ])
    return re.sub(r"\s+", " ", base)


def _coerce_opportunity_df(df):
    if df.empty:
        return pd.DataFrame(columns=OPPORTUNITY_HEADERS)
    for c in OPPORTUNITY_HEADERS:
        if c not in df.columns:
            df[c] = ""
    for c in ["id", "min_followers", "match_score", "favorite"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df[OPPORTUNITY_HEADERS]


def insert_opportunity(data):
    if not storage_configured():
        raise RuntimeError("Google Sheets is not connected yet.")
    data = dict(data)
    if not data.get("date_found"):
        data["date_found"] = datetime.now().strftime("%Y-%m-%d")

    data["unique_key"] = make_unique_key(
        data.get("brand"), data.get("application_url"),
        data.get("contact_email"), data.get("brand_website")
    )

    ws = get_or_create_worksheet("Opportunities", OPPORTUNITY_HEADERS)
    rows = ws.get_all_records()
    existing_keys = {str(r.get("unique_key", "")) for r in rows}
    if data["unique_key"] in existing_keys:
        return False

    ids = []
    for r in rows:
        try:
            ids.append(int(float(r.get("id") or 0)))
        except Exception:
            pass
    data["id"] = max(ids, default=0) + 1

    row = [data.get(h, "") for h in OPPORTUNITY_HEADERS]
    ws.append_row(row, value_input_option="USER_ENTERED")
    return True


def get_opportunities():
    if not storage_configured():
        return pd.DataFrame(columns=OPPORTUNITY_HEADERS)
    try:
        ws = get_or_create_worksheet("Opportunities", OPPORTUNITY_HEADERS)
        rows = ws.get_all_records()
        df = pd.DataFrame(rows)
        return _coerce_opportunity_df(df).sort_values("id", ascending=False) if rows else pd.DataFrame(columns=OPPORTUNITY_HEADERS)
    except Exception as e:
        st.error(f"Google Sheets opportunity read failed: {e}")
        return pd.DataFrame(columns=OPPORTUNITY_HEADERS)


def update_opportunity(opp_id, fields):
    if not fields:
        return
    if not storage_configured():
        raise RuntimeError("Google Sheets is not connected yet.")
    ws = get_or_create_worksheet("Opportunities", OPPORTUNITY_HEADERS)
    values = ws.get_all_values()
    if not values:
        return
    headers = values[0]
    try:
        id_col = headers.index("id")
    except ValueError:
        return
    target_row = None
    for idx, row in enumerate(values[1:], start=2):
        try:
            if int(float(row[id_col])) == int(opp_id):
                target_row = idx
                break
        except Exception:
            continue
    if target_row is None:
        return
    updates = []
    for key, value in fields.items():
        if key not in headers:
            continue
        col = headers.index(key) + 1
        updates.append({"range": gspread.utils.rowcol_to_a1(target_row, col), "values": [[value]]})
    if updates:
        ws.batch_update(updates)


def delete_opportunity(opp_id):
    if not storage_configured():
        raise RuntimeError("Google Sheets is not connected yet.")
    ws = get_or_create_worksheet("Opportunities", OPPORTUNITY_HEADERS)
    values = ws.get_all_values()
    if not values:
        return
    headers = values[0]
    try:
        id_col = headers.index("id")
    except ValueError:
        return
    for idx, row in enumerate(values[1:], start=2):
        try:
            if int(float(row[id_col])) == int(opp_id):
                ws.delete_rows(idx)
                return
        except Exception:
            continue

# -----------------------------
# UI
# -----------------------------
def inject_css():
    st.markdown(f"""
    <style>
        .stApp {{
            background:
                radial-gradient(circle at 8% 8%, rgba(255,214,231,.75), transparent 22%),
                radial-gradient(circle at 92% 0%, rgba(237,221,255,.60), transparent 22%),
                linear-gradient(180deg, #fff9fc 0%, #fff5fa 100%);
            color: {TEXT};
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #ffd9e9 0%, #fff3f8 55%, #f8efff 100%);
            border-right: 1px solid rgba(255,95,162,.18);
        }}

        h1, h2, h3 {{
            color: #602b47 !important;
            letter-spacing: -0.02em;
        }}

        .hero {{
            background: linear-gradient(135deg, rgba(255,255,255,.96), rgba(255,234,243,.96));
            border: 1px solid rgba(255,95,162,.22);
            box-shadow: 0 16px 45px rgba(155, 72, 112, .10);
            padding: 28px 30px;
            border-radius: 28px;
            margin-bottom: 20px;
        }}

        .hero-title {{
            font-size: 42px;
            font-weight: 850;
            color: #5d2944;
            margin-bottom: 4px;
        }}

        .hero-sub {{
            color: {MUTED};
            font-size: 16px;
        }}

        .metric-card {{
            background: rgba(255,255,255,.92);
            border: 1px solid rgba(255,95,162,.16);
            border-radius: 22px;
            padding: 18px 20px;
            box-shadow: 0 10px 28px rgba(126, 73, 102, .08);
            min-height: 118px;
        }}

        .metric-label {{
            color: {MUTED};
            font-size: 13px;
            font-weight: 700;
        }}

        .metric-value {{
            color: #5d2944;
            font-size: 31px;
            font-weight: 850;
            margin-top: 7px;
        }}

        .pink-chip {{
            display: inline-block;
            padding: 5px 11px;
            margin: 2px 3px 2px 0;
            border-radius: 999px;
            background: #ffe1ed;
            color: #9f3467;
            font-size: 12px;
            font-weight: 700;
        }}

        .opportunity-card {{
            background: rgba(255,255,255,.95);
            border: 1px solid rgba(255,95,162,.17);
            border-radius: 22px;
            padding: 18px 20px;
            margin-bottom: 12px;
            box-shadow: 0 8px 24px rgba(130, 75, 102, .07);
        }}

        .score-high {{ color: #169b62; font-weight: 800; }}
        .score-mid {{ color: #c88713; font-weight: 800; }}
        .score-low {{ color: #c24963; font-weight: 800; }}

        div.stButton > button {{
            border-radius: 14px;
            border: 1px solid rgba(255,95,162,.22);
            font-weight: 750;
        }}

        div.stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, {PINK}, {HOT_PINK});
            color: white;
            border: none;
            box-shadow: 0 7px 18px rgba(255,63,146,.22);
        }}

        [data-testid="stMetric"] {{
            background: rgba(255,255,255,.84);
            border: 1px solid rgba(255,95,162,.14);
            padding: 14px;
            border-radius: 18px;
        }}

        .stTextInput input, .stTextArea textarea, .stNumberInput input {{
            border-radius: 13px !important;
        }}

        .section-note {{
            color: {MUTED};
            font-size: 13px;
            margin-top: -8px;
            margin-bottom: 12px;
        }}
    </style>
    """, unsafe_allow_html=True)


def hero(title, subtitle, emoji="✨"):
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-title">{emoji} {title}</div>
            <div class="hero-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label, value, note=""):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div style="font-size:12px;color:{MUTED};margin-top:4px;">{note}</div>
    </div>
    """


# -----------------------------
# Search + extraction
# -----------------------------
def google_search(query, api_key, cse_id, num=10):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {"key": api_key, "cx": cse_id, "q": query, "num": min(num, 10)}
    r = requests.get(url, params=params, timeout=20)
    r.raise_for_status()
    payload = r.json()
    return payload.get("items", [])


def fetch_page(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; CharleysPRTracker/1.0; +https://example.com)"
    }
    r = requests.get(url, headers=headers, timeout=15, allow_redirects=True)
    r.raise_for_status()
    if "text/html" not in r.headers.get("content-type", ""):
        return "", r.url
    return r.text[:1_500_000], r.url


def extract_emails(text):
    return sorted(set(re.findall(
        r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
        text or "",
    )))


def best_contact_email(emails):
    if not emails:
        return ""
    priority = ["pr", "press", "creator", "influencer", "partnership", "collab", "marketing", "affiliate"]
    for keyword in priority:
        for e in emails:
            if keyword in e.lower():
                return e
    return emails[0]


def classify_opportunity(text):
    t = (text or "").lower()
    if "ambassador" in t:
        return "Ambassador"
    if "affiliate" in t:
        return "Affiliate"
    if "ugc" in t:
        return "UGC"
    if "gifting" in t or "gifted" in t or "pr list" in t:
        return "PR / Gifting"
    if "creator program" in t or "creator programme" in t:
        return "Creator Program"
    if "influencer" in t:
        return "Influencer Program"
    if "partnership" in t or "collaborat" in t:
        return "Partnership"
    return "Unknown"


def infer_category(text):
    t = (text or "").lower()
    checks = [
        ("Skincare", ["skincare", "skin care", "serum", "moisturizer", "moisturiser"]),
        ("Makeup", ["makeup", "cosmetics", "lipstick", "foundation", "mascara"]),
        ("Hair", ["haircare", "hair care", "shampoo", "conditioner"]),
        ("Fashion", ["fashion", "clothing", "apparel"]),
        ("Jewellery", ["jewelry", "jewellery"]),
        ("Wellness", ["wellness", "self care", "supplement"]),
        ("Beauty", ["beauty", "cosmetic"]),
        ("Lifestyle", ["lifestyle"]),
    ]
    for cat, words in checks:
        if any(w in t for w in words):
            return cat
    return "Other"


def infer_region(text):
    t = (text or "").lower()
    if "worldwide" in t or "global" in t or "international" in t:
        return "Worldwide"
    if any(x in t for x in ["united kingdom", " uk ", "british", "england", "scotland", "wales"]):
        return "UK"
    if any(x in t for x in ["united states", " usa ", " u.s.", "america"]):
        return "US"
    if "europe" in t or "european" in t or " eu " in t:
        return "Europe"
    return "Unknown"


def extract_min_followers(text):
    t = (text or "").lower().replace(",", "")
    patterns = [
        r"(\d+)\s*k\+?\s*(?:followers|follower)",
        r"(?:minimum|min\.?|at least)\s*(\d+)\s*k\s*(?:followers|follower)?",
        r"(?:minimum|min\.?|at least)\s*(\d{3,})\s*(?:followers|follower)",
    ]
    for pat in patterns:
        m = re.search(pat, t)
        if m:
            n = int(m.group(1))
            if "k" in m.group(0):
                n *= 1000
            return n
    return 0


def clean_brand_from_title(title, url):
    title = re.sub(r"\s+[-|–—]\s+.*$", "", title or "").strip()
    generic = {
        "ambassador program", "affiliate program", "creator program",
        "influencer program", "pr list", "partnerships", "collaborate with us"
    }
    if title and title.lower() not in generic and len(title) <= 70:
        return title
    host = urlparse(url).netloc.lower().replace("www.", "")
    return host.split(".")[0].replace("-", " ").title()


def rule_match_score(profile, region, min_followers, category, opportunity_type, text):
    score = 55
    reasons = []
    country = (profile.get("country") or "").lower()
    niche = (profile.get("niche") or "").lower()

    followers = max(
        int(profile.get("instagram_followers") or 0),
        int(profile.get("tiktok_followers") or 0),
        int(profile.get("youtube_followers") or 0),
    )

    if region == "Worldwide":
        score += 12
        reasons.append("Worldwide program")
    elif region == "UK" and ("uk" in country or "united kingdom" in country or "england" in country):
        score += 15
        reasons.append("UK eligibility matches profile")
    elif region in ("US", "Europe") and region.lower() not in country:
        score -= 10
        reasons.append(f"{region} region may not match")

    if min_followers == 0:
        score += 8
        reasons.append("No public follower minimum found")
    elif followers >= min_followers:
        score += 18
        reasons.append("Follower count meets stated minimum")
    else:
        score -= 28
        reasons.append("Follower count appears below stated minimum")

    if niche and category.lower() in niche:
        score += 12
        reasons.append("Category matches creator niche")
    elif any(word in niche for word in ["beauty", "lifestyle", "fashion", "makeup", "skin", "hair"]):
        if category in ["Beauty", "Makeup", "Skincare", "Hair", "Fashion", "Lifestyle", "Jewellery", "Wellness"]:
            score += 7
            reasons.append("Broad niche fit")

    if opportunity_type != "Unknown":
        score += 5
        reasons.append(f"Clear {opportunity_type} opportunity")

    text_l = (text or "").lower()
    if "closed" in text_l or "applications are closed" in text_l:
        score -= 35
        reasons.append("Page may indicate applications are closed")

    score = max(0, min(100, score))
    eligibility = "Excellent match" if score >= 80 else "Possible match" if score >= 60 else "Low match"
    return score, eligibility, " • ".join(reasons[:5])


def result_to_opportunity(item, profile, source_query):
    title = item.get("title", "")
    link = item.get("link", "")
    snippet = item.get("snippet", "")
    combined = f"{title}\n{snippet}"

    html = ""
    final_url = link
    page_text = combined
    email = ""

    try:
        html, final_url = fetch_page(link)
        soup = BeautifulSoup(html, "html.parser")
        page_text = " ".join(soup.stripped_strings)[:100_000]
        email = best_contact_email(extract_emails(page_text))
    except Exception:
        pass

    full_text = f"{combined}\n{page_text}"
    opp_type = classify_opportunity(full_text)
    category = infer_category(full_text)
    region = infer_region(full_text)
    min_followers = extract_min_followers(full_text)
    brand = clean_brand_from_title(title, final_url)
    score, eligibility, score_reason = rule_match_score(
        profile, region, min_followers, category, opp_type, full_text
    )

    return {
        "brand": brand,
        "category": category,
        "opportunity_type": opp_type,
        "program_name": title[:180],
        "application_url": final_url,
        "contact_email": email,
        "brand_website": f"{urlparse(final_url).scheme}://{urlparse(final_url).netloc}" if final_url else "",
        "region": region,
        "min_followers": min_followers,
        "requirements": snippet[:700],
        "compensation": "Unknown",
        "eligibility": eligibility,
        "match_score": score,
        "score_reason": score_reason,
        "status": "Not Applied",
        "source_query": source_query,
    }


# -----------------------------
# AI helpers
# -----------------------------
def get_openai_client(api_key):
    if not api_key or OpenAI is None:
        return None
    return OpenAI(api_key=api_key)


def ai_text(client, system, user):
    if not client:
        return None
    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return response.output_text
    except Exception as e:
        return f"AI generation failed: {e}"


def creator_context(profile):
    return json.dumps(profile, indent=2, ensure_ascii=False)


# -----------------------------
# Pages
# -----------------------------
def dashboard_page():
    hero(
        "Charley’s PR Tracker",
        "Your pretty little command centre for PR lists, ambassador programs, gifting and creator opportunities.",
        "🎀"
    )

    df = get_opportunities()
    total = len(df)
    strong = int((df["match_score"] >= 80).sum()) if total else 0
    applied = int(df["status"].isin(["Applied", "Follow-up Due", "Responded", "Accepted", "PR Received", "Collaboration"]).sum()) if total else 0
    accepted = int(df["status"].isin(["Accepted", "PR Received", "Collaboration"]).sum()) if total else 0

    cols = st.columns(4)
    cards = [
        ("Opportunities", total, "all saved"),
        ("Strong matches", strong, "80+ match score"),
        ("Applied", applied, "applications started"),
        ("Wins", accepted, "accepted / PR / collabs"),
    ]
    for col, card in zip(cols, cards):
        with col:
            st.markdown(metric_card(*card), unsafe_allow_html=True)

    st.write("")
    st.subheader("🔥 Best opportunities")

    if df.empty:
        st.info("No opportunities yet. Open **Find Opportunities** and start discovering brands.")
        return

    best = df[df["status"] != "Not Interested"].sort_values(
        ["favorite", "match_score"], ascending=[False, False]
    ).head(8)

    for _, row in best.iterrows():
        score_class = "score-high" if row["match_score"] >= 80 else "score-mid" if row["match_score"] >= 60 else "score-low"
        fav = "💗" if row["favorite"] else "♡"
        st.markdown(
            f"""
            <div class="opportunity-card">
              <div style="display:flex;justify-content:space-between;gap:15px;">
                <div>
                  <div style="font-size:20px;font-weight:850;color:#5d2944;">{fav} {row['brand']}</div>
                  <div style="margin-top:5px;">
                    <span class="pink-chip">{row['category']}</span>
                    <span class="pink-chip">{row['opportunity_type']}</span>
                    <span class="pink-chip">{row['region']}</span>
                    <span class="pink-chip">{row['status']}</span>
                  </div>
                </div>
                <div class="{score_class}" style="font-size:24px;">{int(row['match_score'])}/100</div>
              </div>
              <div style="margin-top:10px;color:{MUTED};font-size:13px;">{row['score_reason'] or ''}</div>
            </div>
            """,
            unsafe_allow_html=True
        )


def profile_page():
    hero("Creator Profile", "Set this once so the app can judge which opportunities are actually worth applying for.", "💖")
    p = load_profile()

    with st.form("profile_form"):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Creator name", value=p.get("name", ""))
            email = st.text_input("Email", value=p.get("email", ""))
            country = st.text_input("Country", value=p.get("country", "United Kingdom"))
            niche = st.text_input("Niche", value=p.get("niche", "Beauty / Lifestyle"))
            instagram_url = st.text_input("Instagram URL", value=p.get("instagram_url", ""))
            instagram_followers = st.number_input("Instagram followers", min_value=0, value=int(p.get("instagram_followers") or 0), step=100)
            tiktok_url = st.text_input("TikTok URL", value=p.get("tiktok_url", ""))
            tiktok_followers = st.number_input("TikTok followers", min_value=0, value=int(p.get("tiktok_followers") or 0), step=100)

        with c2:
            youtube_url = st.text_input("YouTube URL", value=p.get("youtube_url", ""))
            youtube_followers = st.number_input("YouTube followers", min_value=0, value=int(p.get("youtube_followers") or 0), step=100)
            average_views = st.number_input("Average views", min_value=0, value=int(p.get("average_views") or 0), step=100)
            engagement_rate = st.number_input("Engagement rate (%)", min_value=0.0, value=float(p.get("engagement_rate") or 0), step=0.1)
            audience = st.text_area("Audience demographics", value=p.get("audience", ""), height=95)
            creator_bio = st.text_area("Creator bio", value=p.get("creator_bio", ""), height=125)
            media_kit_url = st.text_input("Media kit URL", value=p.get("media_kit_url", ""))

        submitted = st.form_submit_button("Save creator profile 💾", type="primary", use_container_width=True)

    if submitted:
        save_profile({
            "name": name,
            "email": email,
            "country": country,
            "niche": niche,
            "instagram_url": instagram_url,
            "instagram_followers": instagram_followers,
            "tiktok_url": tiktok_url,
            "tiktok_followers": tiktok_followers,
            "youtube_url": youtube_url,
            "youtube_followers": youtube_followers,
            "average_views": average_views,
            "engagement_rate": engagement_rate,
            "audience": audience,
            "creator_bio": creator_bio,
            "media_kit_url": media_kit_url,
        })
        st.success("Profile saved ✨")


def find_page(google_key, google_cse):
    hero("Find Opportunities", "Search the web for PR, gifting, ambassador, affiliate, UGC and creator programs.", "🔎")
    profile = load_profile()

    if not profile:
        st.warning("Create the creator profile first so match scores are useful.")

    mode = st.radio(
        "Discovery mode",
        ["Program Search", "Brand Discovery"],
        horizontal=True,
        help="Program Search looks for public applications. Brand Discovery is broader and finds brands plus partnership pages."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        category = st.selectbox("Category", CATEGORIES, index=0)
    with c2:
        region = st.selectbox("Target region", ["Worldwide", "UK", "US", "Europe"])
    with c3:
        results_per_query = st.slider("Results per search", 3, 10, 8)

    custom = st.text_input(
        "Optional keyword",
        placeholder="e.g. small brands, skincare, TikTok, under 10k followers"
    )

    if mode == "Program Search":
        base_queries = [
            f'{category} "ambassador program" {region}',
            f'{category} "PR list" apply {region}',
            f'{category} "creator program" {region}',
            f'{category} influencer gifting application {region}',
            f'{category} affiliate creator program {region}',
            f'{category} UGC creator application {region}',
        ]
    else:
        base_queries = [
            f'{category} brands {region} creator partnerships',
            f'{category} brands {region} influencers collaborate',
            f'{category} brands {region} ambassador',
            f'{category} brands {region} affiliate',
            f'{category} brands {region} gifting creators',
        ]

    query = st.selectbox("Search query", base_queries)
    if custom:
        query = f"{query} {custom}"

    st.caption(f"Search preview: **{query}**")

    if st.button("Find opportunities ✨", type="primary", use_container_width=True):
        if not google_key or not google_cse:
            st.error("Add your Google Custom Search API key and Search Engine ID in the sidebar first.")
            return

        progress = st.progress(0)
        status = st.empty()
        try:
            items = google_search(query, google_key, google_cse, results_per_query)
        except Exception as e:
            st.error(f"Search failed: {e}")
            return

        found = 0
        for i, item in enumerate(items):
            status.write(f"Checking {i+1}/{len(items)} — {item.get('title','')[:70]}")
            try:
                opp = result_to_opportunity(item, profile, query)
                insert_opportunity(opp)
                found += 1
            except Exception:
                pass
            progress.progress((i + 1) / max(1, len(items)))

        status.empty()
        st.success(f"Finished 💗 Checked {len(items)} results and added/updated the database without duplicating exact matches.")
        st.rerun()


def database_page():
    hero("Opportunity Database", "Filter, shortlist and manage every brand opportunity in one place.", "🗂️")
    df = get_opportunities()

    if df.empty:
        st.info("Your database is empty.")
        return

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        category_filter = st.multiselect("Category", sorted(df["category"].dropna().unique().tolist()))
    with f2:
        type_filter = st.multiselect("Opportunity", sorted(df["opportunity_type"].dropna().unique().tolist()))
    with f3:
        status_filter = st.multiselect("Status", STATUSES)
    with f4:
        min_score = st.slider("Minimum score", 0, 100, 0)

    filtered = df.copy()
    if category_filter:
        filtered = filtered[filtered["category"].isin(category_filter)]
    if type_filter:
        filtered = filtered[filtered["opportunity_type"].isin(type_filter)]
    if status_filter:
        filtered = filtered[filtered["status"].isin(status_filter)]
    filtered = filtered[filtered["match_score"] >= min_score]

    sort = st.selectbox("Sort", ["Match score", "Newest", "Brand A-Z", "Favorites first"])
    if sort == "Match score":
        filtered = filtered.sort_values("match_score", ascending=False)
    elif sort == "Newest":
        filtered = filtered.sort_values("id", ascending=False)
    elif sort == "Brand A-Z":
        filtered = filtered.sort_values("brand")
    else:
        filtered = filtered.sort_values(["favorite", "match_score"], ascending=[False, False])

    st.caption(f"{len(filtered)} opportunities")

    for _, row in filtered.iterrows():
        with st.expander(
            f"{'💗' if row['favorite'] else '♡'}  {row['brand']}  —  {row['opportunity_type']}  —  {int(row['match_score'])}/100",
            expanded=False
        ):
            left, right = st.columns([2, 1])

            with left:
                st.write(f"**Program:** {row['program_name'] or '—'}")
                st.write(f"**Category:** {row['category']}  •  **Region:** {row['region']}")
                st.write(f"**Eligibility:** {row['eligibility'] or '—'}")
                st.write(f"**Why:** {row['score_reason'] or '—'}")
                st.write(f"**Requirements:** {row['requirements'] or '—'}")
                st.write(f"**Minimum followers found:** {int(row['min_followers'] or 0):,}" if row['min_followers'] else "**Minimum followers found:** No public minimum")
                if row["contact_email"]:
                    st.write(f"**PR/contact email:** `{row['contact_email']}`")
                if row["application_url"]:
                    st.link_button("Open application / source ↗", row["application_url"])
                if row["brand_website"]:
                    st.link_button("Open brand website ↗", row["brand_website"])

            with right:
                new_status = st.selectbox(
                    "Status",
                    STATUSES,
                    index=STATUSES.index(row["status"]) if row["status"] in STATUSES else 0,
                    key=f"status_{row['id']}"
                )
                fav = st.checkbox("Favorite 💗", value=bool(row["favorite"]), key=f"fav_{row['id']}")
                notes = st.text_area("Notes", value=row["notes"] or "", key=f"notes_{row['id']}")

                if st.button("Save changes", key=f"save_{row['id']}", use_container_width=True):
                    fields = {
                        "status": new_status,
                        "favorite": 1 if fav else 0,
                        "notes": notes,
                    }
                    if new_status == "Applied" and not row["date_applied"]:
                        fields["date_applied"] = datetime.now().strftime("%Y-%m-%d")
                    update_opportunity(int(row["id"]), fields)
                    st.success("Saved")
                    st.rerun()

                if st.button("Delete", key=f"delete_{row['id']}", use_container_width=True):
                    delete_opportunity(int(row["id"]))
                    st.rerun()

    st.divider()
    export_cols = [
        "brand", "category", "opportunity_type", "program_name", "application_url",
        "contact_email", "region", "min_followers", "eligibility", "match_score",
        "status", "favorite", "date_found", "date_applied", "follow_up_date", "notes"
    ]
    csv = filtered[export_cols].to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered CSV", csv, "charleys_pr_tracker_opportunities.csv", "text/csv", use_container_width=True)


def manual_add_page():
    hero("Add Opportunity", "Found something on TikTok, Instagram or by hand? Save it here.", "➕")
    profile = load_profile()

    with st.form("add_form"):
        c1, c2 = st.columns(2)
        with c1:
            brand = st.text_input("Brand *")
            category = st.selectbox("Category", CATEGORIES)
            opp_type = st.selectbox("Opportunity type", OPPORTUNITY_TYPES)
            program_name = st.text_input("Program name")
            application_url = st.text_input("Application URL")
            contact_email = st.text_input("PR / influencer email")
            brand_website = st.text_input("Brand website")

        with c2:
            region = st.selectbox("Region", REGIONS)
            min_followers = st.number_input("Minimum followers (0 = unknown)", min_value=0, step=100)
            requirements = st.text_area("Requirements")
            compensation = st.text_input("Compensation", placeholder="Gifted / commission / paid / unknown")
            notes = st.text_area("Notes")

        submitted = st.form_submit_button("Save opportunity 💗", type="primary", use_container_width=True)

    if submitted:
        if not brand.strip():
            st.error("Brand name is required.")
            return
        score, eligibility, reason = rule_match_score(
            profile, region, min_followers, category, opp_type, requirements
        )
        insert_opportunity({
            "brand": brand.strip(),
            "category": category,
            "opportunity_type": opp_type,
            "program_name": program_name,
            "application_url": application_url,
            "contact_email": contact_email,
            "brand_website": brand_website,
            "region": region,
            "min_followers": min_followers,
            "requirements": requirements,
            "compensation": compensation,
            "eligibility": eligibility,
            "match_score": score,
            "score_reason": reason,
            "status": "Not Applied",
            "notes": notes,
            "source_query": "Manual",
        })
        st.success("Saved ✨")


def application_assistant_page(openai_key):
    hero("Application Assistant", "Generate tailored answers or PR emails. Nothing is sent or submitted automatically.", "💌")
    df = get_opportunities()
    profile = load_profile()

    if df.empty:
        st.info("Add some opportunities first.")
        return

    options = {
        f"{row['brand']} — {row['opportunity_type']} — #{row['id']}": int(row["id"])
        for _, row in df.iterrows()
    }
    selected_label = st.selectbox("Choose opportunity", list(options.keys()))
    opp_id = options[selected_label]
    row = df[df["id"] == opp_id].iloc[0].to_dict()

    client = get_openai_client(openai_key)

    tab1, tab2 = st.tabs(["Application answers", "PR email"])

    with tab1:
        questions = st.text_area(
            "Paste the application question(s)",
            height=180,
            placeholder="Why do you want to work with our brand?\nTell us about your audience..."
        )
        if st.button("Generate answers ✨", key="gen_answers", type="primary", use_container_width=True):
            if not questions.strip():
                st.warning("Paste at least one application question.")
            elif not client:
                st.error("Add an OpenAI API key in the sidebar to use AI writing.")
            else:
                prompt = f"""
CREATOR PROFILE:
{creator_context(profile)}

BRAND OPPORTUNITY:
{json.dumps(row, indent=2, ensure_ascii=False)}

APPLICATION QUESTIONS:
{questions}

Write polished, natural answers for a creator applying to this specific opportunity.
Do not invent achievements, audience demographics, usage of the brand, or prior partnerships.
Keep the tone warm, confident, personable and not corporate.
Separate each answer clearly.
"""
                result = ai_text(
                    client,
                    "You are an expert creator partnership application writer.",
                    prompt
                )
                st.text_area("Generated application", value=result or "", height=350)

    with tab2:
        extra = st.text_area(
            "Optional angle / note",
            placeholder="e.g. she already uses their lip products / wants gifting / interested in UGC"
        )
        if st.button("Generate PR email 💕", key="gen_email", type="primary", use_container_width=True):
            if not client:
                st.error("Add an OpenAI API key in the sidebar to use AI writing.")
            else:
                prompt = f"""
CREATOR PROFILE:
{creator_context(profile)}

BRAND OPPORTUNITY:
{json.dumps(row, indent=2, ensure_ascii=False)}

OPTIONAL USER NOTE:
{extra}

Draft a short personalized PR/collaboration outreach email.
Do not claim the creator already uses the brand unless the note/profile says so.
Mention relevant creator strengths and make the ask clear.
Keep it concise enough to actually get read.
Return a subject line followed by the email.
"""
                result = ai_text(
                    client,
                    "You write concise, high-converting creator-to-brand outreach emails.",
                    prompt
                )
                st.text_area("Generated email", value=result or "", height=320)


def tracker_page():
    hero("Application Tracker", "Keep tabs on applications, follow-ups, replies, PR packages and collaborations.", "📋")
    df = get_opportunities()

    if df.empty:
        st.info("Nothing to track yet.")
        return

    track = df[df["status"] != "Not Applied"].copy()
    if track.empty:
        st.info("Mark an opportunity as **Want to Apply** or **Applied** and it will appear here.")
        return

    cols = [
        "brand", "opportunity_type", "status", "match_score",
        "date_applied", "follow_up_date", "response", "products_received",
        "deliverables", "content_due_date", "payment", "notes"
    ]

    edited = st.data_editor(
        track[["id"] + cols],
        hide_index=True,
        disabled=["id", "brand", "opportunity_type", "match_score"],
        column_config={
            "status": st.column_config.SelectboxColumn("Status", options=STATUSES),
            "match_score": st.column_config.ProgressColumn("Match", min_value=0, max_value=100),
            "date_applied": st.column_config.TextColumn("Applied"),
            "follow_up_date": st.column_config.TextColumn("Follow-up"),
        },
        use_container_width=True,
        num_rows="fixed",
    )

    if st.button("Save tracker changes 💾", type="primary"):
        original = track.set_index("id")
        for _, r in edited.iterrows():
            oid = int(r["id"])
            updates = {}
            for c in cols:
                new_val = r[c]
                old_val = original.loc[oid, c]
                if pd.isna(new_val):
                    new_val = ""
                if pd.isna(old_val):
                    old_val = ""
                if str(new_val) != str(old_val):
                    updates[c] = new_val
            if updates:
                update_opportunity(oid, updates)
        st.success("Tracker updated ✨")
        st.rerun()


def stats_page():
    hero("Stats", "See what kinds of opportunities are actually turning into wins.", "📊")
    df = get_opportunities()
    if df.empty:
        st.info("Stats will appear once opportunities are saved.")
        return

    st.subheader("Status")
    status_counts = df["status"].value_counts()
    st.bar_chart(status_counts)

    st.subheader("Average match score by opportunity")
    avg = df.groupby("opportunity_type")["match_score"].mean().sort_values(ascending=False)
    st.bar_chart(avg)

    applied = df[df["status"].isin(["Applied", "Follow-up Due", "Responded", "Accepted", "Rejected", "PR Received", "Collaboration"])]
    wins = df[df["status"].isin(["Accepted", "PR Received", "Collaboration"])]

    c1, c2, c3 = st.columns(3)
    c1.metric("Applications", len(applied))
    c2.metric("Wins", len(wins))
    rate = (len(wins) / len(applied) * 100) if len(applied) else 0
    c3.metric("Win rate", f"{rate:.1f}%")


# -----------------------------
# App
# -----------------------------
def main():
    st.set_page_config(
        page_title="Charley’s PR Tracker",
        page_icon="🎀",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_db()
    inject_css()

    with st.sidebar:
        st.markdown("## 🎀 Charley’s PR Tracker")
        st.caption("Creator PR & ambassador finder")

        page = st.radio(
            "Navigation",
            [
                "Dashboard",
                "Creator Profile",
                "Find Opportunities",
                "Opportunity Database",
                "Add Opportunity",
                "Application Assistant",
                "Application Tracker",
                "Stats",
            ],
            label_visibility="collapsed",
        )

        st.divider()
        st.markdown("### 🔑 Connections")

        if storage_configured():
            st.success("Google Sheets connected ✅")
        else:
            st.warning("Google Sheets not connected")

        google_key = st.text_input(
            "Google API key",
            value=st.secrets.get("GOOGLE_API_KEY", os.getenv("GOOGLE_API_KEY", "")),
            type="password",
            help="Used for Google Custom Search."
        )
        google_cse = st.text_input(
            "Google Search Engine ID",
            value=st.secrets.get("GOOGLE_CSE_ID", os.getenv("GOOGLE_CSE_ID", "")),
            type="password",
        )
        openai_key = st.text_input(
            "OpenAI API key",
            value=st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", "")),
            type="password",
            help="Optional. Used only for application/email generation."
        )

        st.caption("Keys entered here are used for the current app session. For deployment, store them as environment variables or Streamlit secrets.")

        st.divider()
        st.caption("No applications or emails are automatically submitted in this version. 💗")

    if page == "Dashboard":
        dashboard_page()
    elif page == "Creator Profile":
        profile_page()
    elif page == "Find Opportunities":
        find_page(google_key, google_cse)
    elif page == "Opportunity Database":
        database_page()
    elif page == "Add Opportunity":
        manual_add_page()
    elif page == "Application Assistant":
        application_assistant_page(openai_key)
    elif page == "Application Tracker":
        tracker_page()
    elif page == "Stats":
        stats_page()


if __name__ == "__main__":
    main()
