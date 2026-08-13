import os
import re
import json
from datetime import datetime
from urllib.parse import urlparse

import pandas as pd
import gspread
import requests
import streamlit as st
from bs4 import BeautifulSoup

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# =========================================================
# APP SETTINGS
# =========================================================

APP_TITLE = "Charley’s PR Tracker ✨"

PINK = "#ff5fa2"
HOT_PINK = "#ff3f92"
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

REGIONS = [
    "UK",
    "US",
    "Europe",
    "Worldwide",
    "Unknown",
]


# =========================================================
# GOOGLE SHEETS
# =========================================================

PROFILE_HEADERS = [
    "name",
    "email",
    "country",
    "niche",
    "instagram_url",
    "instagram_followers",
    "tiktok_url",
    "tiktok_followers",
    "youtube_url",
    "youtube_followers",
    "average_views",
    "engagement_rate",
    "audience",
    "creator_bio",
    "media_kit_url",
]


OPPORTUNITY_HEADERS = [
    "id",
    "brand",
    "category",
    "opportunity_type",
    "program_name",
    "application_url",
    "contact_email",
    "brand_website",
    "instagram",
    "region",
    "min_followers",
    "requirements",
    "compensation",
    "eligibility",
    "match_score",
    "score_reason",
    "status",
    "favorite",
    "date_found",
    "date_applied",
    "follow_up_date",
    "contact_person",
    "response",
    "products_received",
    "deliverables",
    "content_due_date",
    "payment",
    "notes",
    "source_query",
    "unique_key",
]


def storage_configured():
    try:
        return (
            bool(st.secrets.get("google_sheet_id"))
            and "gcp_service_account" in st.secrets
        )
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def sheets_client():

    if not storage_configured():
        return None

    credentials = dict(
        st.secrets["gcp_service_account"]
    )

    return gspread.service_account_from_dict(
        credentials
    )


@st.cache_resource(show_spinner=False)
def spreadsheet():

    client = sheets_client()

    if client is None:
        return None

    return client.open_by_key(
        st.secrets["google_sheet_id"]
    )


def get_or_create_worksheet(title, headers):

    ss = spreadsheet()

    if ss is None:
        return None

    try:
        ws = ss.worksheet(title)

    except gspread.WorksheetNotFound:

        ws = ss.add_worksheet(
            title=title,
            rows=1000,
            cols=max(20, len(headers)),
        )

    current_headers = ws.row_values(1)

    if not current_headers:

        ws.update(
            range_name="A1",
            values=[headers],
        )

    else:

        missing = [
            h
            for h in headers
            if h not in current_headers
        ]

        if missing:

            new_headers = (
                current_headers + missing
            )

            ws.update(
                range_name="A1",
                values=[new_headers],
            )

    return ws


def init_storage():

    if not storage_configured():
        return

    get_or_create_worksheet(
        "Profile",
        PROFILE_HEADERS,
    )

    get_or_create_worksheet(
        "Opportunities",
        OPPORTUNITY_HEADERS,
    )


def load_profile():

    if not storage_configured():
        return {}

    try:

        ws = get_or_create_worksheet(
            "Profile",
            PROFILE_HEADERS,
        )

        rows = ws.get_all_records()

        if not rows:
            return {}

        profile = rows[0]

        integer_fields = [
            "instagram_followers",
            "tiktok_followers",
            "youtube_followers",
            "average_views",
        ]

        for field in integer_fields:

            try:
                profile[field] = int(
                    float(
                        profile.get(field) or 0
                    )
                )

            except Exception:
                profile[field] = 0

        try:
            profile["engagement_rate"] = float(
                profile.get(
                    "engagement_rate"
                )
                or 0
            )

        except Exception:
            profile["engagement_rate"] = 0.0

        return profile

    except Exception as e:

        st.error(
            f"Google Sheets profile read failed: {e}"
        )

        return {}


def save_profile(data):

    if not storage_configured():
        raise RuntimeError(
            "Google Sheets is not connected."
        )

    ws = get_or_create_worksheet(
        "Profile",
        PROFILE_HEADERS,
    )

    values = [
        data.get(header, "")
        for header in PROFILE_HEADERS
    ]

    end_cell = gspread.utils.rowcol_to_a1(
        2,
        len(PROFILE_HEADERS),
    )

    ws.update(
        range_name=f"A2:{end_cell}",
        values=[values],
    )


def make_unique_key(
    brand,
    application_url,
    contact_email,
    brand_website,
):

    parts = [
        brand or "",
        application_url or "",
        contact_email or "",
        brand_website or "",
    ]

    return "|".join(
        p.strip().lower()
        for p in parts
    )


def get_opportunities():

    if not storage_configured():

        return pd.DataFrame(
            columns=OPPORTUNITY_HEADERS
        )

    try:

        ws = get_or_create_worksheet(
            "Opportunities",
            OPPORTUNITY_HEADERS,
        )

        rows = ws.get_all_records()

        if not rows:

            return pd.DataFrame(
                columns=OPPORTUNITY_HEADERS
            )

        df = pd.DataFrame(rows)

        for column in OPPORTUNITY_HEADERS:

            if column not in df.columns:
                df[column] = ""

        for column in [
            "id",
            "min_followers",
            "match_score",
            "favorite",
        ]:

            df[column] = pd.to_numeric(
                df[column],
                errors="coerce",
            ).fillna(0).astype(int)

        return df[
            OPPORTUNITY_HEADERS
        ].sort_values(
            "id",
            ascending=False,
        )

    except Exception as e:

        st.error(
            f"Google Sheets read failed: {e}"
        )

        return pd.DataFrame(
            columns=OPPORTUNITY_HEADERS
        )


def insert_opportunity(data):

    if not storage_configured():

        raise RuntimeError(
            "Google Sheets is not connected."
        )

    data = dict(data)

    if not data.get("date_found"):

        data["date_found"] = (
            datetime.now()
            .strftime("%Y-%m-%d")
        )

    data["unique_key"] = make_unique_key(
        data.get("brand"),
        data.get("application_url"),
        data.get("contact_email"),
        data.get("brand_website"),
    )

    ws = get_or_create_worksheet(
        "Opportunities",
        OPPORTUNITY_HEADERS,
    )

    rows = ws.get_all_records()

    existing_keys = {
        str(
            row.get(
                "unique_key",
                "",
            )
        )
        for row in rows
    }

    if data["unique_key"] in existing_keys:
        return False

    ids = []

    for row in rows:

        try:

            ids.append(
                int(
                    float(
                        row.get("id") or 0
                    )
                )
            )

        except Exception:
            pass

    data["id"] = max(
        ids,
        default=0,
    ) + 1

    values = [
        data.get(header, "")
        for header in OPPORTUNITY_HEADERS
    ]

    ws.append_row(
        values,
        value_input_option="USER_ENTERED",
    )

    return True


def update_opportunity(
    opportunity_id,
    fields,
):

    if not fields:
        return

    ws = get_or_create_worksheet(
        "Opportunities",
        OPPORTUNITY_HEADERS,
    )

    values = ws.get_all_values()

    if not values:
        return

    headers = values[0]

    try:
        id_column = headers.index("id")

    except ValueError:
        return

    target_row = None

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):

        try:

            if int(
                float(
                    row[id_column]
                )
            ) == int(
                opportunity_id
            ):

                target_row = row_number
                break

        except Exception:
            continue

    if target_row is None:
        return

    updates = []

    for key, value in fields.items():

        if key not in headers:
            continue

        column = (
            headers.index(key) + 1
        )

        updates.append(
            {
                "range": gspread.utils.rowcol_to_a1(
                    target_row,
                    column,
                ),
                "values": [[value]],
            }
        )

    if updates:
        ws.batch_update(updates)


def delete_opportunity(
    opportunity_id
):

    ws = get_or_create_worksheet(
        "Opportunities",
        OPPORTUNITY_HEADERS,
    )

    values = ws.get_all_values()

    if not values:
        return

    headers = values[0]

    try:
        id_column = headers.index("id")

    except ValueError:
        return

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):

        try:

            if int(
                float(
                    row[id_column]
                )
            ) == int(
                opportunity_id
            ):

                ws.delete_rows(
                    row_number
                )

                return

        except Exception:
            continue


# =========================================================
# SERPER SEARCH
# =========================================================

def serper_search(
    query,
    api_key,
    num=10,
):

    url = (
        "https://google.serper.dev/search"
    )

    headers = {
        "X-API-KEY": api_key,
        "Content-Type": "application/json",
    }

    payload = {
        "q": query,
        "num": min(
            max(
                int(num),
                1,
            ),
            20,
        ),
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=20,
    )

    response.raise_for_status()

    data = response.json()

    results = []

    for item in data.get(
        "organic",
        [],
    ):

        results.append(
            {
                "title": item.get(
                    "title",
                    "",
                ),
                "link": item.get(
                    "link",
                    "",
                ),
                "snippet": item.get(
                    "snippet",
                    "",
                ),
            }
        )

    return results


# =========================================================
# WEB PAGE ANALYSIS
# =========================================================

def fetch_page(url):

    headers = {
        "User-Agent":
        "Mozilla/5.0 CharleysPRTracker/1.0"
    }

    response = requests.get(
        url,
        headers=headers,
        timeout=15,
        allow_redirects=True,
    )

    response.raise_for_status()

    content_type = response.headers.get(
        "content-type",
        "",
    )

    if "text/html" not in content_type:
        return "", response.url

    return (
        response.text[:1_500_000],
        response.url,
    )


def extract_emails(text):

    pattern = (
        r"\b[A-Za-z0-9._%+-]+"
        r"@[A-Za-z0-9.-]+"
        r"\.[A-Za-z]{2,}\b"
    )

    return sorted(
        set(
            re.findall(
                pattern,
                text or "",
            )
        )
    )


def best_contact_email(emails):

    if not emails:
        return ""

    priority = [
        "pr",
        "press",
        "creator",
        "influencer",
        "partnership",
        "collab",
        "marketing",
        "affiliate",
    ]

    for word in priority:

        for email in emails:

            if word in email.lower():
                return email

    return emails[0]


def classify_opportunity(text):

    text = (
        text or ""
    ).lower()

    if "ambassador" in text:
        return "Ambassador"

    if "affiliate" in text:
        return "Affiliate"

    if "ugc" in text:
        return "UGC"

    if (
        "gifting" in text
        or "gifted" in text
        or "pr list" in text
    ):
        return "PR / Gifting"

    if (
        "creator program" in text
        or "creator programme" in text
    ):
        return "Creator Program"

    if "influencer" in text:
        return "Influencer Program"

    if (
        "partnership" in text
        or "collaborat" in text
    ):
        return "Partnership"

    return "Unknown"


def infer_category(text):

    text = (
        text or ""
    ).lower()

    categories = [

        (
            "Skincare",
            [
                "skincare",
                "skin care",
                "serum",
                "moisturizer",
                "moisturiser",
            ],
        ),

        (
            "Makeup",
            [
                "makeup",
                "cosmetics",
                "lipstick",
                "foundation",
                "mascara",
            ],
        ),

        (
            "Hair",
            [
                "haircare",
                "hair care",
                "shampoo",
                "conditioner",
            ],
        ),

        (
            "Fashion",
            [
                "fashion",
                "clothing",
                "apparel",
            ],
        ),

        (
            "Jewellery",
            [
                "jewelry",
                "jewellery",
            ],
        ),

        (
            "Wellness",
            [
                "wellness",
                "self care",
            ],
        ),

        (
            "Beauty",
            [
                "beauty",
                "cosmetic",
            ],
        ),

        (
            "Lifestyle",
            [
                "lifestyle",
            ],
        ),

    ]

    for category, words in categories:

        if any(
            word in text
            for word in words
        ):

            return category

    return "Other"


def infer_region(text):

    text = (
        text or ""
    ).lower()

    if (
        "worldwide" in text
        or "global" in text
        or "international" in text
    ):
        return "Worldwide"

    if any(
        word in text
        for word in [
            "united kingdom",
            " uk ",
            "british",
            "england",
            "scotland",
            "wales",
        ]
    ):
        return "UK"

    if any(
        word in text
        for word in [
            "united states",
            " usa ",
            " u.s.",
        ]
    ):
        return "US"

    if (
        "europe" in text
        or "european" in text
    ):
        return "Europe"

    return "Unknown"


def extract_min_followers(text):

    text = (
        text or ""
    ).lower().replace(
        ",",
        "",
    )

    patterns = [

        r"(\d+)\s*k\+?\s*followers",

        r"(?:minimum|at least)"
        r"\s*(\d+)\s*k"
        r"\s*followers?",

        r"(?:minimum|at least)"
        r"\s*(\d{3,})"
        r"\s*followers?",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if match:

            value = int(
                match.group(1)
            )

            if "k" in match.group(0):
                value *= 1000

            return value

    return 0


def brand_from_result(
    title,
    url,
):

    title = re.sub(
        r"\s+[-|–—]\s+.*$",
        "",
        title or "",
    ).strip()

    generic_titles = {
        "ambassador program",
        "affiliate program",
        "creator program",
        "influencer program",
        "pr list",
        "partnerships",
        "collaborate with us",
    }

    if (
        title
        and title.lower()
        not in generic_titles
        and len(title) <= 70
    ):

        return title

    domain = (
        urlparse(url)
        .netloc
        .lower()
        .replace(
            "www.",
            "",
        )
    )

    name = (
        domain
        .split(".")[0]
        .replace(
            "-",
            " ",
        )
        .title()
    )

    return name


# =========================================================
# MATCH SCORE
# =========================================================

def calculate_match_score(
    profile,
    region,
    min_followers,
    category,
    opportunity_type,
    text,
):

    score = 55
    reasons = []

    country = (
        profile.get(
            "country",
            "",
        )
        or ""
    ).lower()

    niche = (
        profile.get(
            "niche",
            "",
        )
        or ""
    ).lower()

    followers = max(

        int(
            profile.get(
                "instagram_followers",
                0,
            )
            or 0
        ),

        int(
            profile.get(
                "tiktok_followers",
                0,
            )
            or 0
        ),

        int(
            profile.get(
                "youtube_followers",
                0,
            )
            or 0
        ),
    )

    if region == "Worldwide":

        score += 12

        reasons.append(
            "Worldwide program"
        )

    elif (
        region == "UK"
        and (
            "uk" in country
            or "united kingdom"
            in country
            or "england" in country
        )
    ):

        score += 15

        reasons.append(
            "UK eligibility matches profile"
        )

    elif (
        region in [
            "US",
            "Europe",
        ]
        and region.lower()
        not in country
    ):

        score -= 10

        reasons.append(
            f"{region} may not match profile location"
        )

    if min_followers == 0:

        score += 8

        reasons.append(
            "No public follower minimum found"
        )

    elif followers >= min_followers:

        score += 18

        reasons.append(
            "Follower count meets requirement"
        )

    else:

        score -= 28

        reasons.append(
            "Follower count may be below requirement"
        )

    if (
        category.lower()
        in niche
    ):

        score += 12

        reasons.append(
            "Category matches creator niche"
        )

    elif any(
        word in niche
        for word in [
            "beauty",
            "lifestyle",
            "fashion",
            "makeup",
            "skin",
            "hair",
        ]
    ):

        if category in CATEGORIES:

            score += 7

            reasons.append(
                "Broad creator niche fit"
            )

    if opportunity_type != "Unknown":

        score += 5

        reasons.append(
            f"Clear {opportunity_type} opportunity"
        )

    if (
        "applications are closed"
        in text.lower()
        or "applications closed"
        in text.lower()
    ):

        score -= 35

        reasons.append(
            "Applications may be closed"
        )

    score = max(
        0,
        min(
            100,
            score,
        ),
    )

    if score >= 80:

        eligibility = (
            "Excellent match"
        )

    elif score >= 60:

        eligibility = (
            "Possible match"
        )

    else:

        eligibility = (
            "Low match"
        )

    return (
        score,
        eligibility,
        " • ".join(
            reasons[:5]
        ),
    )


def search_result_to_opportunity(
    result,
    profile,
    source_query,
):

    title = result.get(
        "title",
        "",
    )

    link = result.get(
        "link",
        "",
    )

    snippet = result.get(
        "snippet",
        "",
    )

    basic_text = (
        f"{title}\n{snippet}"
    )

    page_text = basic_text
    final_url = link
    contact_email = ""

    try:

        html, final_url = fetch_page(
            link
        )

        soup = BeautifulSoup(
            html,
            "html.parser",
        )

        page_text = " ".join(
            soup.stripped_strings
        )[:100000]

        emails = extract_emails(
            page_text
        )

        contact_email = (
            best_contact_email(
                emails
            )
        )

    except Exception:
        pass

    full_text = (
        f"{basic_text}\n{page_text}"
    )

    opportunity_type = (
        classify_opportunity(
            full_text
        )
    )

    category = infer_category(
        full_text
    )

    region = infer_region(
        full_text
    )

    min_followers = (
        extract_min_followers(
            full_text
        )
    )

    brand = brand_from_result(
        title,
        final_url,
    )

    (
        score,
        eligibility,
        score_reason,
    ) = calculate_match_score(

        profile,
        region,
        min_followers,
        category,
        opportunity_type,
        full_text,
    )

    parsed_url = urlparse(
        final_url
    )

    website = ""

    if parsed_url.netloc:

        website = (
            f"{parsed_url.scheme}"
            f"://"
            f"{parsed_url.netloc}"
        )

    return {

        "brand":
        brand,

        "category":
        category,

        "opportunity_type":
        opportunity_type,

        "program_name":
        title[:180],

        "application_url":
        final_url,

        "contact_email":
        contact_email,

        "brand_website":
        website,

        "region":
        region,

        "min_followers":
        min_followers,

        "requirements":
        snippet[:700],

        "compensation":
        "Unknown",

        "eligibility":
        eligibility,

        "match_score":
        score,

        "score_reason":
        score_reason,

        "status":
        "Not Applied",

        "source_query":
        source_query,
    }


# =========================================================
# OPENAI
# =========================================================

def get_openai_client(api_key):

    if (
        not api_key
        or OpenAI is None
    ):
        return None

    return OpenAI(
        api_key=api_key
    )


def generate_ai_text(
    client,
    system,
    user,
):

    if not client:
        return None

    try:

        response = (
            client.responses.create(
                model="gpt-4.1-mini",
                input=[
                    {
                        "role":
                        "system",
                        "content":
                        system,
                    },
                    {
                        "role":
                        "user",
                        "content":
                        user,
                    },
                ],
            )
        )

        return response.output_text

    except Exception as e:

        return (
            f"AI generation failed: {e}"
        )


# =========================================================
# UI
# =========================================================

def inject_css():

    st.markdown(
        f"""
        <style>

        .stApp {{
            background:
            radial-gradient(
                circle at 8% 8%,
                rgba(255,214,231,.75),
                transparent 22%
            ),
            radial-gradient(
                circle at 92% 0%,
                rgba(237,221,255,.60),
                transparent 22%
            ),
            linear-gradient(
                180deg,
                #fff9fc 0%,
                #fff5fa 100%
            );

            color: {TEXT};
        }}

        [data-testid="stSidebar"] {{

            background:
            linear-gradient(
                180deg,
                #ffd9e9 0%,
                #fff3f8 55%,
                #f8efff 100%
            );

            border-right:
            1px solid
            rgba(255,95,162,.18);
        }}

        h1, h2, h3 {{
            color: #602b47 !important;
        }}

        .hero {{

            background:
            linear-gradient(
                135deg,
                rgba(255,255,255,.96),
                rgba(255,234,243,.96)
            );

            border:
            1px solid
            rgba(255,95,162,.22);

            padding:
            28px 30px;

            border-radius:
            28px;

            margin-bottom:
            20px;

            box-shadow:
            0 16px 45px
            rgba(155,72,112,.10);
        }}

        .hero-title {{

            font-size:
            42px;

            font-weight:
            850;

            color:
            #5d2944;
        }}

        .hero-sub {{

            color:
            {MUTED};

            font-size:
            16px;
        }}

        .metric-card {{

            background:
            rgba(255,255,255,.92);

            border:
            1px solid
            rgba(255,95,162,.16);

            border-radius:
            22px;

            padding:
            18px 20px;

            min-height:
            110px;
        }}

        .metric-label {{

            color:
            {MUTED};

            font-size:
            13px;

            font-weight:
            700;
        }}

        .metric-value {{

            color:
            #5d2944;

            font-size:
            31px;

            font-weight:
            850;

            margin-top:
            7px;
        }}

        .pink-chip {{

            display:
            inline-block;

            padding:
            5px 11px;

            margin:
            2px;

            border-radius:
            999px;

            background:
            #ffe1ed;

            color:
            #9f3467;

            font-size:
            12px;

            font-weight:
            700;
        }}

        div.stButton > button {{

            border-radius:
            14px;

            font-weight:
            750;

            border:
            1px solid
            rgba(255,95,162,.22);
        }}

        div.stButton > button[kind="primary"] {{

            background:
            linear-gradient(
                135deg,
                {PINK},
                {HOT_PINK}
            );

            color:
            white;

            border:
            none;
        }}

        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(
    title,
    subtitle,
    emoji="✨",
):

    st.markdown(
        f"""
        <div class="hero">

            <div class="hero-title">
                {emoji} {title}
            </div>

            <div class="hero-sub">
                {subtitle}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(
    label,
    value,
    note="",
):

    return f"""
    <div class="metric-card">

        <div class="metric-label">
            {label}
        </div>

        <div class="metric-value">
            {value}
        </div>

        <div style="
        font-size:12px;
        color:{MUTED};
        margin-top:4px;
        ">
            {note}
        </div>

    </div>
    """


# =========================================================
# DASHBOARD
# =========================================================

def dashboard_page():

    hero(
        "Charley’s PR Tracker",
        "Find, save and track PR, gifting and ambassador opportunities.",
        "🎀",
    )

    df = get_opportunities()

    total = len(df)

    strong = (
        int(
            (
                df["match_score"]
                >= 80
            ).sum()
        )
        if total
        else 0
    )

    applied = (
        int(
            df["status"].isin(
                [
                    "Applied",
                    "Follow-up Due",
                    "Responded",
                    "Accepted",
                    "PR Received",
                    "Collaboration",
                ]
            ).sum()
        )
        if total
        else 0
    )

    wins = (
        int(
            df["status"].isin(
                [
                    "Accepted",
                    "PR Received",
                    "Collaboration",
                ]
            ).sum()
        )
        if total
        else 0
    )

    columns = st.columns(4)

    cards = [

        (
            "Opportunities",
            total,
            "saved",
        ),

        (
            "Strong Matches",
            strong,
            "80+ score",
        ),

        (
            "Applied",
            applied,
            "applications",
        ),

        (
            "Wins",
            wins,
            "accepted / PR",
        ),
    ]

    for column, card in zip(
        columns,
        cards,
    ):

        with column:

            st.markdown(
                metric_card(
                    *card
                ),
                unsafe_allow_html=True,
            )

    st.write("")

    st.subheader(
        "🔥 Best Opportunities"
    )

    if df.empty:

        st.info(
            "No opportunities yet. Use Find Opportunities."
        )

        return

    best = (
        df[
            df["status"]
            != "Not Interested"
        ]
        .sort_values(
            [
                "favorite",
                "match_score",
            ],
            ascending=[
                False,
                False,
            ],
        )
        .head(8)
    )

    for _, row in best.iterrows():

        favorite = (
            "💗"
            if row["favorite"]
            else "♡"
        )

        st.markdown(
            f"""
            ### {favorite} {row['brand']}

            **{row['opportunity_type']}**

            Match: **{int(row['match_score'])}/100**

            {row['score_reason'] or ''}
            """
        )

        st.divider()


# =========================================================
# PROFILE
# =========================================================

def profile_page():

    hero(
        "Creator Profile",
        "Charley’s information is used to score opportunities.",
        "💖",
    )

    profile = load_profile()

    with st.form(
        "profile_form"
    ):

        left, right = st.columns(2)

        with left:

            name = st.text_input(
                "Creator name",
                value=profile.get(
                    "name",
                    "",
                ),
            )

            email = st.text_input(
                "Email",
                value=profile.get(
                    "email",
                    "",
                ),
            )

            country = st.text_input(
                "Country",
                value=profile.get(
                    "country",
                    "United Kingdom",
                ),
            )

            niche = st.text_input(
                "Niche",
                value=profile.get(
                    "niche",
                    "Beauty / Lifestyle",
                ),
            )

            instagram_url = st.text_input(
                "Instagram URL",
                value=profile.get(
                    "instagram_url",
                    "",
                ),
            )

            instagram_followers = st.number_input(
                "Instagram followers",
                min_value=0,
                value=int(
                    profile.get(
                        "instagram_followers",
                        0,
                    )
                    or 0
                ),
                step=100,
            )

            tiktok_url = st.text_input(
                "TikTok URL",
                value=profile.get(
                    "tiktok_url",
                    "",
                ),
            )

            tiktok_followers = st.number_input(
                "TikTok followers",
                min_value=0,
                value=int(
                    profile.get(
                        "tiktok_followers",
                        0,
                    )
                    or 0
                ),
                step=100,
            )

        with right:

            youtube_url = st.text_input(
                "YouTube URL",
                value=profile.get(
                    "youtube_url",
                    "",
                ),
            )

            youtube_followers = st.number_input(
                "YouTube followers",
                min_value=0,
                value=int(
                    profile.get(
                        "youtube_followers",
                        0,
                    )
                    or 0
                ),
                step=100,
            )

            average_views = st.number_input(
                "Average views",
                min_value=0,
                value=int(
                    profile.get(
                        "average_views",
                        0,
                    )
                    or 0
                ),
                step=100,
            )

            engagement_rate = st.number_input(
                "Engagement rate (%)",
                min_value=0.0,
                value=float(
                    profile.get(
                        "engagement_rate",
                        0,
                    )
                    or 0
                ),
                step=0.1,
            )

            audience = st.text_area(
                "Audience demographics",
                value=profile.get(
                    "audience",
                    "",
                ),
            )

            creator_bio = st.text_area(
                "Creator bio",
                value=profile.get(
                    "creator_bio",
                    "",
                ),
            )

            media_kit_url = st.text_input(
                "Media kit URL",
                value=profile.get(
                    "media_kit_url",
                    "",
                ),
            )

        submit = st.form_submit_button(
            "Save creator profile 💾",
            type="primary",
            use_container_width=True,
        )

    if submit:

        save_profile(
            {

                "name":
                name,

                "email":
                email,

                "country":
                country,

                "niche":
                niche,

                "instagram_url":
                instagram_url,

                "instagram_followers":
                instagram_followers,

                "tiktok_url":
                tiktok_url,

                "tiktok_followers":
                tiktok_followers,

                "youtube_url":
                youtube_url,

                "youtube_followers":
                youtube_followers,

                "average_views":
                average_views,

                "engagement_rate":
                engagement_rate,

                "audience":
                audience,

                "creator_bio":
                creator_bio,

                "media_kit_url":
                media_kit_url,
            }
        )

        st.success(
            "Profile saved ✨"
        )


# =========================================================
# FIND OPPORTUNITIES
# =========================================================

def find_page(
    serper_key
):

    hero(
        "Find Opportunities",
        "Search for PR, gifting, ambassador, affiliate and creator programs.",
        "🔎",
    )

    profile = load_profile()

    if not profile:

        st.warning(
            "Create the Creator Profile first for better match scores."
        )

    mode = st.radio(

        "Discovery mode",

        [
            "Program Search",
            "Brand Discovery",
        ],

        horizontal=True,
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        category = st.selectbox(
            "Category",
            CATEGORIES,
        )

    with c2:

        region = st.selectbox(
            "Target region",
            [
                "Worldwide",
                "UK",
                "US",
                "Europe",
            ],
        )

    with c3:

        result_count = st.slider(
            "Results",
            3,
            20,
            10,
        )

    custom = st.text_input(
        "Optional keyword",
        placeholder="e.g. small brands, TikTok, micro influencer",
    )

    if mode == "Program Search":

        queries = [

            f'{category} "ambassador program" {region}',

            f'{category} "PR list" apply {region}',

            f'{category} "creator program" {region}',

            f'{category} influencer gifting {region}',

            f'{category} affiliate creator program {region}',

            f'{category} UGC creator application {region}',
        ]

    else:

        queries = [

            f'{category} brands {region} creator partnerships',

            f'{category} brands {region} influencer collaboration',

            f'{category} brands {region} ambassador',

            f'{category} brands {region} affiliate',

            f'{category} brands {region} gifting creators',
        ]

    query = st.selectbox(
        "Search query",
        queries,
    )

    if custom:
        query += f" {custom}"

    st.caption(
        f"Search: {query}"
    )

    if st.button(
        "Find opportunities ✨",
        type="primary",
        use_container_width=True,
    ):

        if not serper_key:

            st.error(
                "Opportunity search isn't connected. Check SERPER_API_KEY in Streamlit Secrets."
            )

            return

        try:

            results = serper_search(
                query,
                serper_key,
                result_count,
            )

        except Exception as e:

            st.error(
                f"Search failed: {e}"
            )

            return

        if not results:

            st.warning(
                "No search results found."
            )

            return

        progress = st.progress(0)

        status = st.empty()

        added = 0

        duplicate_count = 0

        for index, result in enumerate(
            results
        ):

            status.write(
                f"Checking {index + 1}/{len(results)} — {result.get('title', '')[:70]}"
            )

            try:

                opportunity = (
                    search_result_to_opportunity(
                        result,
                        profile,
                        query,
                    )
                )

                inserted = (
                    insert_opportunity(
                        opportunity
                    )
                )

                if inserted:
                    added += 1

                else:
                    duplicate_count += 1

            except Exception:
                pass

            progress.progress(
                (
                    index + 1
                )
                / len(results)
            )

        status.empty()

        st.success(
            f"Finished 💗 Added {added} new opportunities. "
            f"Skipped {duplicate_count} duplicates."
        )


# =========================================================
# DATABASE
# =========================================================

def database_page():

    hero(
        "Opportunity Database",
        "Review, favorite and track every opportunity.",
        "🗂️",
    )

    df = get_opportunities()

    if df.empty:

        st.info(
            "No opportunities yet."
        )

        return

    f1, f2, f3 = st.columns(3)

    with f1:

        category_filter = (
            st.multiselect(
                "Category",
                sorted(
                    df[
                        "category"
                    ]
                    .dropna()
                    .unique()
                    .tolist()
                ),
            )
        )

    with f2:

        type_filter = (
            st.multiselect(
                "Type",
                sorted(
                    df[
                        "opportunity_type"
                    ]
                    .dropna()
                    .unique()
                    .tolist()
                ),
            )
        )

    with f3:

        minimum_score = st.slider(
            "Minimum score",
            0,
            100,
            0,
        )

    filtered = df.copy()

    if category_filter:

        filtered = filtered[
            filtered[
                "category"
            ].isin(
                category_filter
            )
        ]

    if type_filter:

        filtered = filtered[
            filtered[
                "opportunity_type"
            ].isin(
                type_filter
            )
        ]

    filtered = filtered[
        filtered[
            "match_score"
        ]
        >= minimum_score
    ]

    filtered = filtered.sort_values(
        "match_score",
        ascending=False,
    )

    st.caption(
        f"{len(filtered)} opportunities"
    )

    for _, row in filtered.iterrows():

        favorite_icon = (
            "💗"
            if row["favorite"]
            else "♡"
        )

        title = (
            f"{favorite_icon} "
            f"{row['brand']} — "
            f"{row['opportunity_type']} — "
            f"{int(row['match_score'])}/100"
        )

        with st.expander(title):

            left, right = st.columns(
                [2, 1]
            )

            with left:

                st.write(
                    f"**Program:** {row['program_name'] or '—'}"
                )

                st.write(
                    f"**Category:** {row['category']}"
                )

                st.write(
                    f"**Region:** {row['region']}"
                )

                st.write(
                    f"**Eligibility:** {row['eligibility']}"
                )

                st.write(
                    f"**Why:** {row['score_reason'] or '—'}"
                )

                st.write(
                    f"**Requirements:** {row['requirements'] or '—'}"
                )

                if row[
                    "contact_email"
                ]:

                    st.write(
                        f"**Contact:** `{row['contact_email']}`"
                    )

                if row[
                    "application_url"
                ]:

                    st.link_button(
                        "Open application / source ↗",
                        row[
                            "application_url"
                        ],
                    )

                if row[
                    "brand_website"
                ]:

                    st.link_button(
                        "Brand website ↗",
                        row[
                            "brand_website"
                        ],
                    )

            with right:

                status = st.selectbox(

                    "Status",

                    STATUSES,

                    index=(
                        STATUSES.index(
                            row[
                                "status"
                            ]
                        )
                        if row[
                            "status"
                        ]
                        in STATUSES
                        else 0
                    ),

                    key=f"status_{row['id']}",
                )

                favorite = st.checkbox(

                    "Favorite 💗",

                    value=bool(
                        row[
                            "favorite"
                        ]
                    ),

                    key=f"fav_{row['id']}",
                )

                notes = st.text_area(

                    "Notes",

                    value=(
                        row[
                            "notes"
                        ]
                        or ""
                    ),

                    key=f"notes_{row['id']}",
                )

                if st.button(
                    "Save changes",
                    key=f"save_{row['id']}",
                    use_container_width=True,
                ):

                    update_fields = {

                        "status":
                        status,

                        "favorite":
                        1 if favorite else 0,

                        "notes":
                        notes,
                    }

                    if (
                        status
                        == "Applied"
                        and not row[
                            "date_applied"
                        ]
                    ):

                        update_fields[
                            "date_applied"
                        ] = (
                            datetime.now()
                            .strftime(
                                "%Y-%m-%d"
                            )
                        )

                    update_opportunity(
                        int(
                            row["id"]
                        ),
                        update_fields,
                    )

                    st.success(
                        "Saved ✨"
                    )

                    st.rerun()

                if st.button(

                    "Delete",

                    key=f"delete_{row['id']}",

                    use_container_width=True,
                ):

                    delete_opportunity(
                        int(
                            row["id"]
                        )
                    )

                    st.rerun()


# =========================================================
# MANUAL ADD
# =========================================================

def manual_add_page():

    hero(
        "Add Opportunity",
        "Save a PR opportunity you found yourself.",
        "➕",
    )

    profile = load_profile()

    with st.form(
        "manual_add"
    ):

        left, right = st.columns(2)

        with left:

            brand = st.text_input(
                "Brand *"
            )

            category = st.selectbox(
                "Category",
                CATEGORIES,
            )

            opportunity_type = (
                st.selectbox(
                    "Opportunity type",
                    OPPORTUNITY_TYPES,
                )
            )

            program_name = (
                st.text_input(
                    "Program name"
                )
            )

            application_url = (
                st.text_input(
                    "Application URL"
                )
            )

            contact_email = (
                st.text_input(
                    "PR / influencer email"
                )
            )

            brand_website = (
                st.text_input(
                    "Brand website"
                )
            )

        with right:

            region = st.selectbox(
                "Region",
                REGIONS,
            )

            min_followers = (
                st.number_input(
                    "Minimum followers",
                    min_value=0,
                    step=100,
                )
            )

            requirements = (
                st.text_area(
                    "Requirements"
                )
            )

            compensation = (
                st.text_input(
                    "Compensation",
                    placeholder="Gifted / commission / paid",
                )
            )

            notes = st.text_area(
                "Notes"
            )

        submit = (
            st.form_submit_button(
                "Save opportunity 💗",
                type="primary",
                use_container_width=True,
            )
        )

    if submit:

        if not brand.strip():

            st.error(
                "Brand name is required."
            )

            return

        (
            score,
            eligibility,
            reason,
        ) = calculate_match_score(

            profile,

            region,

            min_followers,

            category,

            opportunity_type,

            requirements,
        )

        insert_opportunity(
            {

                "brand":
                brand.strip(),

                "category":
                category,

                "opportunity_type":
                opportunity_type,

                "program_name":
                program_name,

                "application_url":
                application_url,

                "contact_email":
                contact_email,

                "brand_website":
                brand_website,

                "region":
                region,

                "min_followers":
                min_followers,

                "requirements":
                requirements,

                "compensation":
                compensation,

                "eligibility":
                eligibility,

                "match_score":
                score,

                "score_reason":
                reason,

                "status":
                "Not Applied",

                "notes":
                notes,

                "source_query":
                "Manual",
            }
        )

        st.success(
            "Opportunity saved ✨"
        )


# =========================================================
# APPLICATION ASSISTANT
# =========================================================

def application_assistant_page(
    openai_key
):

    hero(
        "Application Assistant",
        "Generate application answers and PR emails. Nothing is automatically submitted.",
        "💌",
    )

    df = get_opportunities()

    profile = load_profile()

    if df.empty:

        st.info(
            "Add opportunities first."
        )

        return

    options = {

        f"{row['brand']} — {row['opportunity_type']} — #{row['id']}":
        int(
            row["id"]
        )

        for _, row
        in df.iterrows()
    }

    selected = st.selectbox(
        "Choose opportunity",
        list(
            options.keys()
        ),
    )

    opportunity_id = options[
        selected
    ]

    row = (
        df[
            df["id"]
            == opportunity_id
        ]
        .iloc[0]
        .to_dict()
    )

    client = get_openai_client(
        openai_key
    )

    tab1, tab2 = st.tabs(
        [
            "Application Answers",
            "PR Email",
        ]
    )

    with tab1:

        questions = st.text_area(

            "Paste the application questions",

            height=180,
        )

        if st.button(
            "Generate answers ✨",
            type="primary",
        ):

            if not questions.strip():

                st.warning(
                    "Paste some questions first."
                )

            elif not client:

                st.error(
                    "Add OPENAI_API_KEY to Streamlit Secrets first."
                )

            else:

                prompt = f"""
CREATOR PROFILE:
{json.dumps(profile, indent=2)}

BRAND OPPORTUNITY:
{json.dumps(row, indent=2)}

APPLICATION QUESTIONS:
{questions}

Write natural, polished application answers.

Do not invent achievements, statistics, previous collaborations or claims that the creator uses the brand.

Keep the answers confident, warm and personal.
"""

                result = generate_ai_text(

                    client,

                    "You help social media creators apply to brand partnership programs.",

                    prompt,
                )

                st.text_area(

                    "Generated answers",

                    value=result or "",

                    height=350,
                )

    with tab2:

        extra = st.text_area(
            "Optional note",
            placeholder="Anything specific Charley wants mentioned",
        )

        if st.button(
            "Generate PR email 💕",
            type="primary",
        ):

            if not client:

                st.error(
                    "Add OPENAI_API_KEY to Streamlit Secrets first."
                )

            else:

                prompt = f"""
CREATOR PROFILE:
{json.dumps(profile, indent=2)}

BRAND:
{json.dumps(row, indent=2)}

EXTRA NOTE:
{extra}

Write a short personalised creator-to-brand PR or collaboration email.

Do not claim she already uses the brand unless the profile or note says that.

Include a subject line.

Keep it concise and natural.
"""

                result = generate_ai_text(

                    client,

                    "You write creator outreach emails to beauty and lifestyle brands.",

                    prompt,
                )

                st.text_area(

                    "Generated email",

                    value=result or "",

                    height=300,
                )


# =========================================================
# APPLICATION TRACKER
# =========================================================

def tracker_page():

    hero(
        "Application Tracker",
        "Track applications, responses and PR wins.",
        "📋",
    )

    df = get_opportunities()

    if df.empty:

        st.info(
            "Nothing to track yet."
        )

        return

    track = df[
        df["status"]
        != "Not Applied"
    ].copy()

    if track.empty:

        st.info(
            "Change an opportunity to Want to Apply or Applied first."
        )

        return

    columns = [

        "id",

        "brand",

        "opportunity_type",

        "status",

        "match_score",

        "date_applied",

        "follow_up_date",

        "response",

        "products_received",

        "deliverables",

        "payment",

        "notes",
    ]

    edited = st.data_editor(

        track[columns],

        hide_index=True,

        disabled=[
            "id",
            "brand",
            "opportunity_type",
            "match_score",
        ],

        column_config={

            "status":
            st.column_config.SelectboxColumn(
                "Status",
                options=STATUSES,
            ),

            "match_score":
            st.column_config.ProgressColumn(
                "Match",
                min_value=0,
                max_value=100,
            ),
        },

        use_container_width=True,
    )

    if st.button(
        "Save tracker changes 💾",
        type="primary",
    ):

        original = (
            track
            .set_index(
                "id"
            )
        )

        for _, row in edited.iterrows():

            opportunity_id = int(
                row["id"]
            )

            updates = {}

            for column in columns:

                if column in [
                    "id",
                    "brand",
                    "opportunity_type",
                    "match_score",
                ]:
                    continue

                new_value = row[
                    column
                ]

                old_value = original.loc[
                    opportunity_id,
                    column,
                ]

                if pd.isna(
                    new_value
                ):
                    new_value = ""

                if pd.isna(
                    old_value
                ):
                    old_value = ""

                if str(
                    new_value
                ) != str(
                    old_value
                ):

                    updates[
                        column
                    ] = new_value

            if updates:

                update_opportunity(
                    opportunity_id,
                    updates,
                )

        st.success(
            "Tracker updated ✨"
        )

        st.rerun()


# =========================================================
# STATS
# =========================================================

def stats_page():

    hero(
        "Stats",
        "See which opportunities are turning into wins.",
        "📊",
    )

    df = get_opportunities()

    if df.empty:

        st.info(
            "Stats will appear once opportunities are saved."
        )

        return

    st.subheader(
        "Application Status"
    )

    st.bar_chart(
        df[
            "status"
        ].value_counts()
    )

    st.subheader(
        "Average Match Score"
    )

    averages = (

        df.groupby(
            "opportunity_type"
        )[
            "match_score"
        ]

        .mean()

        .sort_values(
            ascending=False
        )
    )

    st.bar_chart(
        averages
    )

    applications = df[

        df[
            "status"
        ].isin(
            [
                "Applied",
                "Follow-up Due",
                "Responded",
                "Accepted",
                "Rejected",
                "PR Received",
                "Collaboration",
            ]
        )
    ]

    wins = df[

        df[
            "status"
        ].isin(
            [
                "Accepted",
                "PR Received",
                "Collaboration",
            ]
        )
    ]

    rate = (

        len(wins)
        / len(applications)
        * 100

        if len(
            applications
        )

        else 0
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Applications",
        len(
            applications
        ),
    )

    c2.metric(
        "Wins",
        len(
            wins
        ),
    )

    c3.metric(
        "Win Rate",
        f"{rate:.1f}%",
    )


# =========================================================
# MAIN APP
# =========================================================

def main():

    st.set_page_config(

        page_title=
        "Charley’s PR Tracker",

        page_icon=
        "🎀",

        layout=
        "wide",

        initial_sidebar_state=
        "expanded",
    )

    init_storage()

    inject_css()

    with st.sidebar:

        st.markdown(
            "## 🎀 Charley’s PR Tracker"
        )

        st.caption(
            "Creator PR & ambassador finder"
        )

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

            label_visibility=
            "collapsed",
        )

        st.divider()

        st.markdown(
            "### 🔑 Connections"
        )

        if storage_configured():

            st.success(
                "Google Sheets connected ✅"
            )

        else:

            st.warning(
                "Google Sheets not connected"
            )

        serper_key = st.secrets.get(

            "SERPER_API_KEY",

            os.getenv(
                "SERPER_API_KEY",
                "",
            ),
        )

        if serper_key:

            st.success(
                "Opportunity search connected ✅"
            )

        else:

            st.warning(
                "Opportunity search not connected"
            )

        openai_key = st.secrets.get(

            "OPENAI_API_KEY",

            os.getenv(
                "OPENAI_API_KEY",
                "",
            ),
        )

        if openai_key:

            st.success(
                "AI assistant connected ✅"
            )

        else:

            st.caption(
                "AI assistant not connected — optional"
            )

        st.divider()

        st.caption(
            "Nothing is automatically submitted or emailed. 💗"
        )

    if page == "Dashboard":

        dashboard_page()

    elif page == "Creator Profile":

        profile_page()

    elif page == "Find Opportunities":

        find_page(
            serper_key
        )

    elif page == "Opportunity Database":

        database_page()

    elif page == "Add Opportunity":

        manual_add_page()

    elif page == "Application Assistant":

        application_assistant_page(
            openai_key
        )

    elif page == "Application Tracker":

        tracker_page()

    elif page == "Stats":

        stats_page()


if __name__ == "__main__":
    main()
