import os, re, json
from datetime import datetime
from urllib.parse import urlparse

import pandas as pd
import requests
import streamlit as st
import gspread
from bs4 import BeautifulSoup

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

APP_NAME = "Charley’s PR Tracker"

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

OPP_TYPES = [
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

REGIONS = [
    "UK",
    "US",
    "Europe",
    "Worldwide",
    "Unknown",
]

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

OPP_HEADERS = [
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


# =========================
# UI
# =========================

def inject_css():
    st.markdown(
        """
        <style>

        .stApp{
            background:
            radial-gradient(
                circle at 10% 5%,
                rgba(255,210,230,.7),
                transparent 24%
            ),
            radial-gradient(
                circle at 90% 0%,
                rgba(235,220,255,.6),
                transparent 24%
            ),
            linear-gradient(
                180deg,
                #fffafd 0%,
                #fff3f8 100%
            );
        }

        [data-testid="stSidebar"]{
            background:
            linear-gradient(
                180deg,
                #ffd9e9 0%,
                #fff4f8 58%,
                #f7efff 100%
            );

            border-right:
            1px solid rgba(255,95,162,.18);
        }

        h1,h2,h3{
            color:#5f2b48!important;
        }

        .hero{
            background:
            linear-gradient(
                135deg,
                rgba(255,255,255,.98),
                rgba(255,235,244,.98)
            );

            border:
            1px solid rgba(255,95,162,.22);

            border-radius:26px;

            padding:24px 28px;

            margin-bottom:18px;

            box-shadow:
            0 14px 36px rgba(140,75,105,.09);
        }

        .hero-title{
            font-size:38px;
            font-weight:850;
            color:#5f2b48;
        }

        .hero-sub{
            font-size:15px;
            color:#7c6672;
            margin-top:5px;
        }

        .metric-card{
            background:rgba(255,255,255,.94);

            border:
            1px solid rgba(255,95,162,.16);

            border-radius:20px;

            padding:17px 18px;

            min-height:100px;

            box-shadow:
            0 8px 22px rgba(140,75,105,.06);
        }

        .metric-label{
            font-size:13px;
            color:#7c6672;
            font-weight:700;
        }

        .metric-value{
            font-size:29px;
            color:#5f2b48;
            font-weight:850;
            margin-top:6px;
        }

        div.stButton > button{
            border-radius:13px;
            font-weight:750;
        }

        div.stButton > button[kind="primary"]{
            background:
            linear-gradient(
                135deg,
                #ff5fa2,
                #ff3f92
            );

            color:white;
            border:none;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(
    title,
    subtitle,
    emoji="🎀",
):
    st.markdown(
        f"""<div class="hero"><div class="hero-title">{emoji} {title}</div><div class="hero-sub">{subtitle}</div></div>""",
        unsafe_allow_html=True,
    )


def metric_card(
    label,
    value,
    note="",
):
    return f"""<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div style="font-size:12px;color:#7c6672;margin-top:4px;">{note}</div></div>"""


# =========================
# GOOGLE SHEETS
# =========================

def storage_ready():
    try:
        return (
            bool(
                st.secrets.get(
                    "google_sheet_id"
                )
            )
            and
            "gcp_service_account"
            in st.secrets
        )
    except Exception:
        return False


@st.cache_resource(
    show_spinner=False
)
def sheet_client():

    if not storage_ready():
        return None

    return (
        gspread
        .service_account_from_dict(
            dict(
                st.secrets[
                    "gcp_service_account"
                ]
            )
        )
    )


@st.cache_resource(
    show_spinner=False
)
def workbook():

    client = sheet_client()

    if not client:
        return None

    return client.open_by_key(
        st.secrets[
            "google_sheet_id"
        ]
    )


def worksheet(
    title,
    headers,
):

    ss = workbook()

    if ss is None:
        return None

    try:

        ws = ss.worksheet(
            title
        )

    except gspread.WorksheetNotFound:

        ws = ss.add_worksheet(
            title=title,
            rows=1000,
            cols=max(
                20,
                len(headers),
            ),
        )

    current = ws.row_values(
        1
    )

    if not current:

        ws.update(
            range_name="A1",
            values=[
                headers
            ],
        )

    else:

        merged = (
            current
            +
            [
                h
                for h
                in headers
                if h not in current
            ]
        )

        if merged != current:

            ws.update(
                range_name="A1",
                values=[
                    merged
                ],
            )

    return ws


def init_storage():

    if storage_ready():

        worksheet(
            "Profile",
            PROFILE_HEADERS,
        )

        worksheet(
            "Opportunities",
            OPP_HEADERS,
        )


def load_profile():

    if not storage_ready():
        return {}

    try:

        rows = worksheet(
            "Profile",
            PROFILE_HEADERS,
        ).get_all_records()

        if not rows:
            return {}

        profile = rows[0]

        number_fields = [
            "instagram_followers",
            "tiktok_followers",
            "youtube_followers",
            "average_views",
        ]

        for key in number_fields:

            try:
                profile[key] = int(
                    float(
                        profile.get(
                            key
                        )
                        or 0
                    )
                )

            except Exception:
                profile[key] = 0

        try:
            profile[
                "engagement_rate"
            ] = float(
                profile.get(
                    "engagement_rate"
                )
                or 0
            )

        except Exception:
            profile[
                "engagement_rate"
            ] = 0.0

        return profile

    except Exception as e:

        st.error(
            f"Google Sheets profile read failed: {e}"
        )

        return {}


def save_profile(
    data
):

    ws = worksheet(
        "Profile",
        PROFILE_HEADERS,
    )

    values = [
        data.get(
            header,
            "",
        )
        for header
        in PROFILE_HEADERS
    ]

    end = (
        gspread.utils
        .rowcol_to_a1(
            2,
            len(
                PROFILE_HEADERS
            ),
        )
    )

    ws.update(
        range_name=f"A2:{end}",
        values=[
            values
        ],
    )


def unique_key(
    data
):

    fields = [
        "brand",
        "application_url",
        "contact_email",
        "brand_website",
    ]

    return "|".join(
        str(
            data.get(
                field,
                "",
            )
            or ""
        )
        .strip()
        .lower()

        for field
        in fields
    )


def get_opportunities():

    if not storage_ready():

        return pd.DataFrame(
            columns=OPP_HEADERS
        )

    try:

        rows = worksheet(
            "Opportunities",
            OPP_HEADERS,
        ).get_all_records()

        if not rows:

            return pd.DataFrame(
                columns=OPP_HEADERS
            )

        df = pd.DataFrame(
            rows
        )

        for column in OPP_HEADERS:

            if column not in df.columns:
                df[column] = ""

        for column in [
            "id",
            "min_followers",
            "match_score",
            "favorite",
        ]:

            df[column] = (
                pd.to_numeric(
                    df[column],
                    errors="coerce",
                )
                .fillna(0)
                .astype(int)
            )

        return (
            df[
                OPP_HEADERS
            ]
            .sort_values(
                "id",
                ascending=False,
            )
        )

    except Exception as e:

        st.error(
            f"Google Sheets opportunity read failed: {e}"
        )

        return pd.DataFrame(
            columns=OPP_HEADERS
        )


def insert_opportunity(
    data
):

    ws = worksheet(
        "Opportunities",
        OPP_HEADERS,
    )

    rows = (
        ws.get_all_records()
    )

    data = dict(
        data
    )

    if not data.get(
        "date_found"
    ):
        data[
            "date_found"
        ] = (
            datetime.now()
            .strftime(
                "%Y-%m-%d"
            )
        )

    data[
        "unique_key"
    ] = unique_key(
        data
    )

    existing = {
        str(
            row.get(
                "unique_key",
                "",
            )
        )
        for row
        in rows
    }

    if (
        data[
            "unique_key"
        ]
        in existing
    ):
        return False

    ids = []

    for row in rows:

        try:

            ids.append(
                int(
                    float(
                        row.get(
                            "id"
                        )
                        or 0
                    )
                )
            )

        except Exception:
            pass

    data[
        "id"
    ] = max(
        ids,
        default=0,
    ) + 1

    values = [
        data.get(
            header,
            "",
        )
        for header
        in OPP_HEADERS
    ]

    ws.append_row(
        values,
        value_input_option=
        "USER_ENTERED",
    )

    return True


def update_opportunity(
    opp_id,
    fields,
):

    ws = worksheet(
        "Opportunities",
        OPP_HEADERS,
    )

    values = (
        ws.get_all_values()
    )

    if not values:
        return

    headers = values[0]

    if "id" not in headers:
        return

    id_col = headers.index(
        "id"
    )

    target = None

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):

        try:

            if int(
                float(
                    row[
                        id_col
                    ]
                )
            ) == int(
                opp_id
            ):

                target = (
                    row_number
                )

                break

        except Exception:
            pass

    if not target:
        return

    updates = []

    for key, value in fields.items():

        if key not in headers:
            continue

        column = (
            headers.index(
                key
            )
            + 1
        )

        updates.append(
            {
                "range":
                gspread.utils
                .rowcol_to_a1(
                    target,
                    column,
                ),

                "values":
                [
                    [value]
                ],
            }
        )

    if updates:
        ws.batch_update(
            updates
        )


def delete_opportunity(
    opp_id
):

    ws = worksheet(
        "Opportunities",
        OPP_HEADERS,
    )

    values = (
        ws.get_all_values()
    )

    if not values:
        return

    headers = values[0]

    if "id" not in headers:
        return

    id_col = headers.index(
        "id"
    )

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):

        try:

            if int(
                float(
                    row[
                        id_col
                    ]
                )
            ) == int(
                opp_id
            ):

                ws.delete_rows(
                    row_number
                )

                return

        except Exception:
            pass


# =========================
# SERPER SEARCH
# =========================

def serper_search(
    query,
    api_key,
    num=10,
):

    response = requests.post(
        "https://google.serper.dev/search",

        headers={
            "X-API-KEY":
            api_key,

            "Content-Type":
            "application/json",
        },

        json={
            "q":
            query,

            "num":
            min(
                max(
                    int(
                        num
                    ),
                    1,
                ),
                20,
            ),
        },

        timeout=20,
    )

    response.raise_for_status()

    data = (
        response.json()
    )

    return [
        {
            "title":
            item.get(
                "title",
                "",
            ),

            "link":
            item.get(
                "link",
                "",
            ),

            "snippet":
            item.get(
                "snippet",
                "",
            ),
        }

        for item
        in data.get(
            "organic",
            [],
        )
    ]


# =========================
# WEB EXTRACTION
# =========================

def fetch_page(
    url
):

    response = requests.get(
        url,

        headers={
            "User-Agent":
            "Mozilla/5.0 (compatible; CharleysPRTracker/1.0)"
        },

        timeout=15,

        allow_redirects=True,
    )

    response.raise_for_status()

    if (
        "text/html"
        not in response.headers.get(
            "content-type",
            "",
        )
    ):

        return (
            "",
            response.url,
        )

    return (
        response.text[
            :1_500_000
        ],
        response.url,
    )


def extract_emails(
    text
):

    pattern = (
        r"\b"
        r"[A-Za-z0-9._%+-]+"
        r"@"
        r"[A-Za-z0-9.-]+"
        r"\."
        r"[A-Za-z]{2,}"
        r"\b"
    )

    return sorted(
        set(
            re.findall(
                pattern,
                text or "",
            )
        )
    )


def best_email(
    emails
):

    keywords = [
        "pr",
        "press",
        "creator",
        "influencer",
        "partnership",
        "collab",
        "marketing",
        "affiliate",
    ]

    for keyword in keywords:

        for email in emails:

            if keyword in email.lower():
                return email

    if emails:
        return emails[0]

    return ""


def classify(
    text
):

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
        or
        "gifted" in text
        or
        "pr list" in text
    ):
        return "PR / Gifting"

    if (
        "creator program"
        in text
        or
        "creator programme"
        in text
    ):
        return "Creator Program"

    if "influencer" in text:
        return "Influencer Program"

    if (
        "partnership"
        in text
        or
        "collaborat"
        in text
    ):
        return "Partnership"

    return "Unknown"


def infer_category(
    text
):

    text = (
        text or ""
    ).lower()

    checks = [

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

    for category, words in checks:

        if any(
            word in text
            for word
            in words
        ):

            return category

    return "Other"


def infer_region(
    text
):

    text = (
        text or ""
    ).lower()

    if any(
        word in text
        for word
        in [
            "worldwide",
            "global",
            "international",
        ]
    ):
        return "Worldwide"

    if any(
        word in text
        for word
        in [
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
        for word
        in [
            "united states",
            " usa ",
            " u.s.",
        ]
    ):
        return "US"

    if any(
        word in text
        for word
        in [
            "europe",
            "european",
        ]
    ):
        return "Europe"

    return "Unknown"


def min_followers(
    text
):

    text = (
        text or ""
    ).lower().replace(
        ",",
        "",
    )

    patterns = [

        r"(\d+)\s*k\+?\s*followers?",

        r"(?:minimum|at least)\s*(\d+)\s*k\s*followers?",

        r"(?:minimum|at least)\s*(\d{3,})\s*followers?",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
        )

        if match:

            number = int(
                match.group(1)
            )

            if (
                "k"
                in match.group(0)
            ):
                number *= 1000

            return number

    return 0


def brand_name(
    title,
    url,
):

    title = re.sub(
        r"\s+[-|–—]\s+.*$",
        "",
        title or "",
    ).strip()

    generic = {
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
        and
        title.lower()
        not in generic
        and
        len(title)
        <= 70
    ):
        return title

    host = (
        urlparse(
            url
        )
        .netloc
        .lower()
        .replace(
            "www.",
            "",
        )
    )

    return (
        host
        .split(".")[0]
        .replace(
            "-",
            " ",
        )
        .title()
    )


# =========================
# MATCH SCORE
# =========================

def score_match(
    profile,
    region,
    minimum,
    category,
    opp_type,
    text,
):

    score = 55
    reasons = []

    country = str(
        profile.get(
            "country",
            "",
        )
        or ""
    ).lower()

    niche = str(
        profile.get(
            "niche",
            "",
        )
        or ""
    ).lower()

    followers = max(

        int(
            profile.get(
                "instagram_followers"
            )
            or 0
        ),

        int(
            profile.get(
                "tiktok_followers"
            )
            or 0
        ),

        int(
            profile.get(
                "youtube_followers"
            )
            or 0
        ),
    )

    if region == "Worldwide":

        score += 12

        reasons.append(
            "Worldwide"
        )

    elif (
        region == "UK"
        and
        any(
            word in country
            for word
            in [
                "uk",
                "united kingdom",
                "england",
            ]
        )
    ):

        score += 15

        reasons.append(
            "UK location match"
        )

    elif (
        region
        in [
            "US",
            "Europe",
        ]
        and
        region.lower()
        not in country
    ):

        score -= 10

        reasons.append(
            f"{region} may not match"
        )

    if minimum == 0:

        score += 8

        reasons.append(
            "No public follower minimum"
        )

    elif followers >= minimum:

        score += 18

        reasons.append(
            "Follower requirement met"
        )

    else:

        score -= 28

        reasons.append(
            "Follower requirement may not be met"
        )

    if (
        category.lower()
        in niche
    ):

        score += 12

        reasons.append(
            "Niche match"
        )

    elif any(
        word in niche
        for word
        in [
            "beauty",
            "lifestyle",
            "fashion",
            "makeup",
            "skin",
            "hair",
        ]
    ):

        score += 7

        reasons.append(
            "Broad niche fit"
        )

    if (
        opp_type
        != "Unknown"
    ):

        score += 5

        reasons.append(
            f"Clear {opp_type} opportunity"
        )

    if (
        "applications are closed"
        in (
            text or ""
        ).lower()
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


def result_to_opportunity(
    result,
    profile,
    query,
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

    page_text = (
        f"{title}\n{snippet}"
    )

    final_url = link

    email = ""

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

        email = best_email(
            extract_emails(
                page_text
            )
        )

    except Exception:
        pass

    full_text = (
        f"{title}\n"
        f"{snippet}\n"
        f"{page_text}"
    )

    opp_type = classify(
        full_text
    )

    category = infer_category(
        full_text
    )

    region = infer_region(
        full_text
    )

    minimum = min_followers(
        full_text
    )

    (
        score,
        eligibility,
        reason,
    ) = score_match(

        profile,
        region,
        minimum,
        category,
        opp_type,
        full_text,
    )

    parsed = urlparse(
        final_url
    )

    website = ""

    if parsed.netloc:

        website = (
            f"{parsed.scheme}"
            f"://"
            f"{parsed.netloc}"
        )

    return {

        "brand":
        brand_name(
            title,
            final_url,
        ),

        "category":
        category,

        "opportunity_type":
        opp_type,

        "program_name":
        title[:180],

        "application_url":
        final_url,

        "contact_email":
        email,

        "brand_website":
        website,

        "region":
        region,

        "min_followers":
        minimum,

        "requirements":
        snippet[:700],

        "compensation":
        "Unknown",

        "eligibility":
        eligibility,

        "match_score":
        score,

        "score_reason":
        reason,

        "status":
        "Not Applied",

        "source_query":
        query,
    }


# =========================
# OPENAI
# =========================

def ai_client():

    key = st.secrets.get(
        "OPENAI_API_KEY",
        os.getenv(
            "OPENAI_API_KEY",
            "",
        ),
    )

    if (
        not key
        or
        not OpenAI
    ):
        return None

    return OpenAI(
        api_key=key
    )


def ai_generate(
    system,
    user,
):

    client = ai_client()

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

        return (
            response.output_text
        )

    except Exception as e:

        return (
            f"AI generation failed: {e}"
        )


# =========================
# DASHBOARD
# =========================

def dashboard_page():

    hero(
        APP_NAME,
        "Find, save and track PR, gifting and ambassador opportunities.",
    )

    df = get_opportunities()

    total = len(
        df
    )

    strong = (
        int(
            (
                df[
                    "match_score"
                ]
                >= 80
            ).sum()
        )
        if total
        else 0
    )

    applied = (
        int(
            df[
                "status"
            ]
            .isin(
                [
                    "Applied",
                    "Follow-up Due",
                    "Responded",
                    "Accepted",
                    "PR Received",
                    "Collaboration",
                ]
            )
            .sum()
        )
        if total
        else 0
    )

    wins = (
        int(
            df[
                "status"
            ]
            .isin(
                [
                    "Accepted",
                    "PR Received",
                    "Collaboration",
                ]
            )
            .sum()
        )
        if total
        else 0
    )

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

    columns = st.columns(
        4
    )

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
            df[
                "status"
            ]
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
            if row[
                "favorite"
            ]
            else "♡"
        )

        st.markdown(
            f"### {favorite} {row['brand']}"
        )

        st.write(
            f"**{row['opportunity_type']}** · "
            f"{int(row['match_score'])}/100 · "
            f"{row['region']}"
        )

        if row[
            "score_reason"
        ]:

            st.caption(
                row[
                    "score_reason"
                ]
            )

        st.divider()


# =========================
# CREATOR PROFILE
# =========================

def profile_page():

    hero(
        "Creator Profile",
        "Used to score which opportunities are a good fit.",
        "💖",
    )

    profile = load_profile()

    with st.form(
        "profile"
    ):

        left, right = st.columns(
            2
        )

        with left:

            name = st.text_input(
                "Creator name",
                profile.get(
                    "name",
                    "",
                ),
            )

            email = st.text_input(
                "Email",
                profile.get(
                    "email",
                    "",
                ),
            )

            country = st.text_input(
                "Country",
                profile.get(
                    "country",
                    "United Kingdom",
                ),
            )

            niche = st.text_input(
                "Niche",
                profile.get(
                    "niche",
                    "Beauty / Lifestyle",
                ),
            )

            instagram_url = (
                st.text_input(
                    "Instagram URL",
                    profile.get(
                        "instagram_url",
                        "",
                    ),
                )
            )

            instagram_followers = (
                st.number_input(
                    "Instagram followers",
                    min_value=0,
                    value=int(
                        profile.get(
                            "instagram_followers"
                        )
                        or 0
                    ),
                    step=100,
                )
            )

            tiktok_url = (
                st.text_input(
                    "TikTok URL",
                    profile.get(
                        "tiktok_url",
                        "",
                    ),
                )
            )

            tiktok_followers = (
                st.number_input(
                    "TikTok followers",
                    min_value=0,
                    value=int(
                        profile.get(
                            "tiktok_followers"
                        )
                        or 0
                    ),
                    step=100,
                )
            )

        with right:

            youtube_url = (
                st.text_input(
                    "YouTube URL",
                    profile.get(
                        "youtube_url",
                        "",
                    ),
                )
            )

            youtube_followers = (
                st.number_input(
                    "YouTube followers",
                    min_value=0,
                    value=int(
                        profile.get(
                            "youtube_followers"
                        )
                        or 0
                    ),
                    step=100,
                )
            )

            average_views = (
                st.number_input(
                    "Average views",
                    min_value=0,
                    value=int(
                        profile.get(
                            "average_views"
                        )
                        or 0
                    ),
                    step=100,
                )
            )

            engagement_rate = (
                st.number_input(
                    "Engagement rate (%)",
                    min_value=0.0,
                    value=float(
                        profile.get(
                            "engagement_rate"
                        )
                        or 0
                    ),
                    step=0.1,
                )
            )

            audience = (
                st.text_area(
                    "Audience demographics",
                    profile.get(
                        "audience",
                        "",
                    ),
                )
            )

            creator_bio = (
                st.text_area(
                    "Creator bio",
                    profile.get(
                        "creator_bio",
                        "",
                    ),
                )
            )

            media_kit_url = (
                st.text_input(
                    "Media kit URL",
                    profile.get(
                        "media_kit_url",
                        "",
                    ),
                )
            )

        submit = (
            st.form_submit_button(
                "Save creator profile 💾",
                type="primary",
                use_container_width=True,
            )
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


# =========================
# FIND OPPORTUNITIES
# =========================

def find_page(
    serper_key
):

    hero(
        "Find Opportunities",
        "Search for PR, gifting, ambassador, affiliate and creator programs.",
        "🔎",
    )

    profile = load_profile()

    mode = st.radio(

        "Discovery mode",

        [
            "Program Search",
            "Brand Discovery",
        ],

        horizontal=True,
    )

    c1, c2, c3 = st.columns(
        3
    )

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

        count = st.slider(
            "Results",
            3,
            20,
            10,
        )

    custom = st.text_input(
        "Optional keyword",
        placeholder="e.g. micro influencer, skincare, TikTok",
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

        query += (
            f" {custom}"
        )

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
                "Opportunity search is not connected."
            )

            return

        try:

            results = serper_search(
                query,
                serper_key,
                count,
            )

        except Exception as e:

            st.error(
                f"Search failed: {e}"
            )

            return

        added = 0
        duplicates = 0

        progress = st.progress(
            0
        )

        label = st.empty()

        for index, result in enumerate(
            results
        ):

            label.write(
                f"Checking {index + 1}/{len(results)} — {result.get('title', '')[:70]}"
            )

            try:

                opportunity = (
                    result_to_opportunity(
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
                    duplicates += 1

            except Exception:
                pass

            progress.progress(
                (
                    index + 1
                )
                /
                max(
                    1,
                    len(results),
                )
            )

        label.empty()

        st.success(
            f"Finished 💗 Added {added} new opportunities. "
            f"Skipped {duplicates} duplicates."
        )


# =========================
# DATABASE
# =========================

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

    c1, c2, c3, c4 = st.columns(
        4
    )

    with c1:

        categories = (
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

    with c2:

        types = (
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

    with c3:

        statuses = (
            st.multiselect(
                "Status",
                STATUSES,
            )
        )

    with c4:

        min_score = (
            st.slider(
                "Minimum score",
                0,
                100,
                0,
            )
        )

    filtered = df[
        df[
            "match_score"
        ]
        >= min_score
    ].copy()

    if categories:

        filtered = filtered[
            filtered[
                "category"
            ]
            .isin(
                categories
            )
        ]

    if types:

        filtered = filtered[
            filtered[
                "opportunity_type"
            ]
            .isin(
                types
            )
        ]

    if statuses:

        filtered = filtered[
            filtered[
                "status"
            ]
            .isin(
                statuses
            )
        ]

    filtered = (
        filtered.sort_values(
            [
                "favorite",
                "match_score",
            ],
            ascending=[
                False,
                False,
            ],
        )
    )

    for _, row in filtered.iterrows():

        favorite = (
            "💗"
            if row[
                "favorite"
            ]
            else "♡"
        )

        title = (
            f"{favorite} "
            f"{row['brand']} — "
            f"{row['opportunity_type']} — "
            f"{int(row['match_score'])}/100"
        )

        with st.expander(
            title
        ):

            left, right = st.columns(
                [
                    2,
                    1,
                ]
            )

            with left:

                st.write(
                    f"**Program:** {row['program_name'] or '—'}"
                )

                st.write(
                    f"**Category:** {row['category']} · "
                    f"**Region:** {row['region']}"
                )

                st.write(
                    f"**Eligibility:** {row['eligibility'] or '—'}"
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

                    key=
                    f"s{row['id']}",
                )

                favorite_checked = (
                    st.checkbox(
                        "Favorite 💗",
                        value=bool(
                            row[
                                "favorite"
                            ]
                        ),
                        key=
                        f"f{row['id']}",
                    )
                )

                notes = (
                    st.text_area(
                        "Notes",
                        value=(
                            row[
                                "notes"
                            ]
                            or ""
                        ),
                        key=
                        f"n{row['id']}",
                    )
                )

                if st.button(
                    "Save changes",
                    key=
                    f"save{row['id']}",
                    use_container_width=True,
                ):

                    fields = {
                        "status":
                        status,

                        "favorite":
                        1
                        if favorite_checked
                        else 0,

                        "notes":
                        notes,
                    }

                    if (
                        status
                        == "Applied"
                        and
                        not row[
                            "date_applied"
                        ]
                    ):

                        fields[
                            "date_applied"
                        ] = (
                            datetime.now()
                            .strftime(
                                "%Y-%m-%d"
                            )
                        )

                    update_opportunity(
                        int(
                            row[
                                "id"
                            ]
                        ),
                        fields,
                    )

                    st.rerun()

                if st.button(
                    "Delete",
                    key=
                    f"del{row['id']}",
                    use_container_width=True,
                ):

                    delete_opportunity(
                        int(
                            row[
                                "id"
                            ]
                        )
                    )

                    st.rerun()


# =========================
# MANUAL ADD
# =========================

def add_page():

    hero(
        "Add Opportunity",
        "Save something you found yourself.",
        "➕",
    )

    profile = load_profile()

    with st.form(
        "add"
    ):

        left, right = st.columns(
            2
        )

        with left:

            brand = st.text_input(
                "Brand *"
            )

            category = st.selectbox(
                "Category",
                CATEGORIES,
            )

            opp_type = st.selectbox(
                "Opportunity type",
                OPP_TYPES,
            )

            program = st.text_input(
                "Program name"
            )

            url = st.text_input(
                "Application URL"
            )

            email = st.text_input(
                "PR / influencer email"
            )

            website = st.text_input(
                "Brand website"
            )

        with right:

            region = st.selectbox(
                "Region",
                REGIONS,
            )

            minimum = st.number_input(
                "Minimum followers",
                min_value=0,
                step=100,
            )

            requirements = st.text_area(
                "Requirements"
            )

            compensation = (
                st.text_input(
                    "Compensation",
                    placeholder=
                    "Gifted / commission / paid",
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
        ) = score_match(

            profile,
            region,
            minimum,
            category,
            opp_type,
            requirements,
        )

        insert_opportunity(
            {
                "brand":
                brand.strip(),

                "category":
                category,

                "opportunity_type":
                opp_type,

                "program_name":
                program,

                "application_url":
                url,

                "contact_email":
                email,

                "brand_website":
                website,

                "region":
                region,

                "min_followers":
                minimum,

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


# =========================
# APPLICATION ASSISTANT
# =========================

def assistant_page():

    hero(
        "Application Assistant",
        "Generate answers and PR emails. Nothing is automatically submitted.",
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
            row[
                "id"
            ]
        )

        for _, row
        in df.iterrows()
    }

    choice = st.selectbox(
        "Choose opportunity",
        list(
            options.keys()
        ),
    )

    opportunity_id = (
        options[
            choice
        ]
    )

    row = (
        df[
            df[
                "id"
            ]
            == opportunity_id
        ]
        .iloc[0]
        .to_dict()
    )

    tab1, tab2 = st.tabs(
        [
            "Application Answers",
            "PR Email",
        ]
    )

    with tab1:

        questions = (
            st.text_area(
                "Paste application questions",
                height=160,
            )
        )

        if st.button(
            "Generate answers ✨",
            type="primary",
        ):

            if not ai_client():

                st.error(
                    "Add OPENAI_API_KEY to Streamlit Secrets first."
                )

            elif not questions.strip():

                st.warning(
                    "Paste some questions first."
                )

            else:

                prompt = f"""
CREATOR PROFILE:
{json.dumps(profile, indent=2)}

OPPORTUNITY:
{json.dumps(row, indent=2)}

QUESTIONS:
{questions}

Write natural, concise answers.

Do not invent facts.

Do not claim she uses the brand unless stated.
"""

                result = (
                    ai_generate(
                        "You help creators apply to brand partnership programs.",
                        prompt,
                    )
                )

                st.text_area(
                    "Generated answers",
                    result or "",
                    height=320,
                )

    with tab2:

        note = st.text_area(
            "Optional note"
        )

        if st.button(
            "Generate PR email 💕",
            type="primary",
        ):

            if not ai_client():

                st.error(
                    "Add OPENAI_API_KEY to Streamlit Secrets first."
                )

            else:

                prompt = f"""
CREATOR PROFILE:
{json.dumps(profile, indent=2)}

OPPORTUNITY:
{json.dumps(row, indent=2)}

NOTE:
{note}

Write a short personalised PR or collaboration email.

Include a subject line.

Do not invent facts.
"""

                result = (
                    ai_generate(
                        "You write concise creator outreach emails.",
                        prompt,
                    )
                )

                st.text_area(
                    "Generated email",
                    result or "",
                    height=280,
                )


# =========================
# APPLICATION TRACKER
# =========================

def tracker_page():

    hero(
        "Application Tracker",
        "Track applications, replies, PR packages and collaborations.",
        "📋",
    )

    df = get_opportunities()

    track = (
        df[
            df[
                "status"
            ]
            != "Not Applied"
        ]
        .copy()
    )

    if track.empty:

        st.info(
            "Mark something Want to Apply or Applied first."
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

        track[
            columns
        ],

        hide_index=True,

        disabled=[
            "id",
            "brand",
            "opportunity_type",
            "match_score",
        ],

        column_config={

            "status":
            st.column_config
            .SelectboxColumn(
                "Status",
                options=
                STATUSES,
            ),

            "match_score":
            st.column_config
            .ProgressColumn(
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
                row[
                    "id"
                ]
            )

            changes = {}

            for column in columns:

                if column in [
                    "id",
                    "brand",
                    "opportunity_type",
                    "match_score",
                ]:
                    continue

                new_value = (
                    ""
                    if pd.isna(
                        row[
                            column
                        ]
                    )
                    else row[
                        column
                    ]
                )

                old_value = (
                    ""
                    if pd.isna(
                        original.loc[
                            opportunity_id,
                            column,
                        ]
                    )
                    else
                    original.loc[
                        opportunity_id,
                        column,
                    ]
                )

                if str(
                    new_value
                ) != str(
                    old_value
                ):

                    changes[
                        column
                    ] = new_value

            if changes:

                update_opportunity(
                    opportunity_id,
                    changes,
                )

        st.rerun()


# =========================
# STATS
# =========================

def stats_page():

    hero(
        "Stats",
        "See what is turning into wins.",
        "📊",
    )

    df = get_opportunities()

    if df.empty:

        st.info(
            "Stats appear once opportunities are saved."
        )

        return

    st.subheader(
        "Status"
    )

    st.bar_chart(
        df[
            "status"
        ]
        .value_counts()
    )

    st.subheader(
        "Average match score"
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

    applied = df[
        df[
            "status"
        ]
        .isin(
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
        ]
        .isin(
            [
                "Accepted",
                "PR Received",
                "Collaboration",
            ]
        )
    ]

    rate = (
        len(
            wins
        )
        /
        len(
            applied
        )
        * 100

        if len(
            applied
        )

        else 0
    )

    c1, c2, c3 = st.columns(
        3
    )

    c1.metric(
        "Applications",
        len(
            applied
        ),
    )

    c2.metric(
        "Wins",
        len(
            wins
        ),
    )

    c3.metric(
        "Win rate",
        f"{rate:.1f}%",
    )


# =========================
# MAIN
# =========================

def main():

    st.set_page_config(
        page_title=
        APP_NAME,

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
            f"## 🎀 {APP_NAME}"
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

        if storage_ready():

            st.success(
                "Google Sheets connected ✅"
            )

        else:

            st.warning(
                "Google Sheets not connected"
            )

        serper_key = (
            st.secrets.get(
                "SERPER_API_KEY",
                os.getenv(
                    "SERPER_API_KEY",
                    "",
                ),
            )
        )

        if serper_key:

            st.success(
                "Opportunity search connected ✅"
            )

        else:

            st.warning(
                "Opportunity search not connected"
            )

        if ai_client():

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

    pages = {

        "Dashboard":
        dashboard_page,

        "Creator Profile":
        profile_page,

        "Find Opportunities":
        lambda:
        find_page(
            serper_key
        ),

        "Opportunity Database":
        database_page,

        "Add Opportunity":
        add_page,

        "Application Assistant":
        assistant_page,

        "Application Tracker":
        tracker_page,

        "Stats":
        stats_page,
    }

    pages[
        page
    ]()


if __name__ == "__main__":
    main()
