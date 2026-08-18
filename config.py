# ============================================================
# DTM Survey Alignment Tool — config.py
# Page config, default parameters, and CSS
# ============================================================

import streamlit as st

# ────────────────────────────────────────────────────────────
# Page configuration
# ────────────────────────────────────────────────────────────

PAGE_CONFIG = {
    "page_title": "DTM Survey Alignment Tool",
    "page_icon":  "🗺️",
    "layout":     "wide",
    "initial_sidebar_state": "expanded",
    "menu_items": {
        "About": (
            "**DTM Survey Alignment Tool**  \n"
            "Matches IOM DTM survey XLSForms against the Global Data Kit.  \n"
            "Produces a colour-coded Excel diagnostic report."
        ),
    },
}

# ────────────────────────────────────────────────────────────
# Default parameters
# ────────────────────────────────────────────────────────────

DEFAULTS = {
    "datakit_sheet":    "myWorkSheet",
    "survey_sheet":     "survey",
    "formcomponents":   "Baseline Sub-Area Assessment",
    "fuzzy_threshold":  0.20,
    "missing_categories": ["Core"],
    "df1_key_col":      "name",
    "df1_text_col":     "label::English(en)",
    "df1_type_col":     "type",
    "df2_key_col":      "FieldName",
    "df2_text_col":     "QuestionText(en)",
    "df2_id_col":       "FieldUniqueId",
    "df2_type_col":     "QuestionAnswerType",
    "df2_comp_col":     "QuestionComponent",
}

# ────────────────────────────────────────────────────────────
# Form components — DTM Global Data Kit component list
# (values as they appear in the Datakit's QuestionComponent column)
# ────────────────────────────────────────────────────────────

FORM_COMPONENTS = [
    "Flow Monitoring Survey",
    "Baseline Location Assessment",
    "Multi-Sectoral Location Assessment (MSLA)",
    "Baseline Sub-Area Assessment",
    "Emergency Event Tracking",
    "Migrant Presence Monitoring",
    "Departure Area Monitoring Tool",
    "Transhumance Early Warning Mechanism",
    "Household Needs Assessment",
    "Flow Monitoring Screener",
    "Displacement and Demographic Calculator (DDC)",
    "Flow Monitoring Registry",
    "Transhumance Flow Monitoring",
    "Drivers of Migration",
    "Intentions Survey - IND",
    "Intentions Survey - HH",
    "Transhumance Individual Surveys",
    "Flow Monitoring Counter",
    "Counter Trafficking Module",
    "Transhumance Presence and Profile",
]

# ────────────────────────────────────────────────────────────
# Status colour palette  (used in both UI and Excel export)
# ────────────────────────────────────────────────────────────

STATUS_COLOURS_HEX = {
    "LikelyAligned":       "#C6EFCE",
    "NeedsVerification":   "#FFEB9C",
    "FPReview":            "#FFC7CE",
    "DiscouragedQuestion": "#FCE4D6",
    "MissingCoreQuestion": "#E2D9F3",
    "DoesNotNeed2Align":   "#F5F5F5",
}

STATUS_FONT_HEX = {
    "LikelyAligned":       "#375623",
    "NeedsVerification":   "#9C6500",
    "FPReview":            "#9C0006",
    "DiscouragedQuestion": "#833C00",
    "MissingCoreQuestion": "#4B2E83",
    "DoesNotNeed2Align":   "#666666",
}

STATUS_EMOJI = {
    "LikelyAligned":       "✅",
    "NeedsVerification":   "🔍",
    "FPReview":            "⚠️",
    "DiscouragedQuestion": "🚫",
    "MissingCoreQuestion": "🔴",
    "DoesNotNeed2Align":   "—",
}

# ────────────────────────────────────────────────────────────
# CSS
# ────────────────────────────────────────────────────────────

CUSTOM_CSS = """
/* ═══════════════════════════════════════════════════════════
   DTM Survey Alignment Tool — stylesheet
   All colours are hardcoded hex (no CSS variables) so they
   survive Streamlit Cloud's theme injection reliably.
   ═══════════════════════════════════════════════════════════ */

/* ── App header ──────────────────────────────────────────── */
.dtm-title {
    color: #1F3864;
    font-size: 1.85rem;
    font-weight: 700;
    line-height: 1.2;
    margin-bottom: 0.1rem;
}
.dtm-sub {
    color: #555555;
    font-size: 0.88rem;
    margin-top: 0.1rem;
    margin-bottom: 0;
}

/* ── Sidebar background ──────────────────────────────────── */
[data-testid="stSidebar"] {
    background-color: #F0F4FA !important;
    border-right: 1px solid #CBD5E1;
}
[data-testid="stSidebar"] > div:first-child {
    background-color: #F0F4FA !important;
}

/* ── ALL text inside the sidebar — force dark ────────────── */
/* This is the main fix: Streamlit Cloud can render labels
   as near-white if not explicitly overridden.              */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] h4,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] div {
    color: #1A1A2E !important;
}

/* Widget labels specifically */
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] span,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] label {
    color: #1A1A2E !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
}

/* Caption / help text */
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
[data-testid="stSidebar"] small {
    color: #4A5568 !important;
    font-size: 0.78rem !important;
}

/* Slider value label */
[data-testid="stSidebar"] [data-testid="stSlider"] p,
[data-testid="stSidebar"] [data-testid="stSlider"] span {
    color: #1A1A2E !important;
}

/* Toggle label */
[data-testid="stSidebar"] [data-testid="stToggle"] p,
[data-testid="stSidebar"] [data-testid="stToggle"] span:not([data-baseweb]) {
    color: #1A1A2E !important;
}

/* Text input — label + input field */
[data-testid="stSidebar"] [data-testid="stTextInput"] label,
[data-testid="stSidebar"] [data-testid="stTextInput"] p {
    color: #1A1A2E !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] [data-testid="stTextInput"] input {
    background-color: #FFFFFF !important;
    color: #1A1A2E !important;
    border: 1px solid #94A3B8 !important;
    border-radius: 4px;
    font-size: 0.86rem;
}
[data-testid="stSidebar"] [data-testid="stTextInput"] input:focus {
    border-color: #2E5797 !important;
    box-shadow: 0 0 0 2px rgba(46, 87, 151, 0.18) !important;
}

/* Info / alert boxes inside sidebar */
[data-testid="stSidebar"] [data-testid="stAlert"] p,
[data-testid="stSidebar"] [data-testid="stAlert"] span {
    color: #1A1A2E !important;
}

/* ── Sidebar section headers (.sidebar-section class) ─────── */
.sidebar-section {
    color: #1F3864 !important;
    font-weight: 700;
    font-size: 0.88rem;
    letter-spacing: 0.02em;
    text-transform: uppercase;
    margin-top: 0.3rem;
    margin-bottom: 0.3rem;
    padding-bottom: 0.2rem;
    border-bottom: 2px solid #2E5797;
    display: block;
}

/* ── Expander (sidebar) ──────────────────────────────────── */
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    color: #1F3864 !important;
    font-weight: 600;
    font-size: 0.86rem;
    background-color: #E2ECF8;
    border-radius: 4px;
    padding: 0.3rem 0.5rem;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover {
    background-color: #D0E0F4;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary p,
[data-testid="stSidebar"] [data-testid="stExpander"] summary span {
    color: #1F3864 !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] p,
[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] label,
[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] span {
    color: #1A1A2E !important;
}

/* ── Divider ─────────────────────────────────────────────── */
[data-testid="stSidebar"] hr {
    border: none;
    border-top: 1px solid #CBD5E1;
    margin: 0.75rem 0;
}

/* ── Slider track + thumb ────────────────────────────────── */
[data-testid="stSlider"] [data-baseweb="slider"] [data-testid="stThumbValue"] {
    color: #1F3864 !important;
    background-color: #D6E4F7 !important;
}
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
    background-color: #2E5797 !important;
    border-color: #1F3864 !important;
}

/* ── Toggle active colour ────────────────────────────────── */
[data-baseweb="toggle"] div[data-checked="true"] {
    background-color: #2E5797 !important;
}

/* ── MAIN AREA — metric cards ────────────────────────────── */
div[data-testid="metric-container"] {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-left: 4px solid #2E5797;
    border-radius: 6px;
    padding: 0.7rem 0.9rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
div[data-testid="metric-container"] label,
div[data-testid="metric-container"] [data-testid="stMetricLabel"] p {
    font-size: 0.78rem !important;
    color: #555555 !important;
    font-weight: 500 !important;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 1.55rem !important;
    font-weight: 700 !important;
    color: #1F3864 !important;
}

/* ── Primary button (Run) ────────────────────────────────── */
div[data-testid="stButton"] > button[kind="primary"] {
    background-color: #1F3864 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px;
    font-weight: 600;
    letter-spacing: 0.02em;
    transition: background-color 0.18s ease;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    background-color: #2E5797 !important;
}
div[data-testid="stButton"] > button[kind="primary"]:disabled {
    background-color: #94A3B8 !important;
    color: #E2E8F0 !important;
}

/* ── Download button ─────────────────────────────────────── */
div[data-testid="stDownloadButton"] > button {
    background-color: #1F3864 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px;
    font-weight: 600;
    width: 100%;
    transition: background-color 0.18s ease;
}
div[data-testid="stDownloadButton"] > button:hover {
    background-color: #2E5797 !important;
}

/* ── File uploader ───────────────────────────────────────── */
[data-testid="stFileUploaderDropzone"] {
    border: 1.5px dashed #2E5797 !important;
    border-radius: 6px !important;
    background-color: #F8FAFD !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #1F3864 !important;
    background-color: #D6E4F7 !important;
}
[data-testid="stFileUploaderDropzone"] span,
[data-testid="stFileUploaderDropzone"] p,
[data-testid="stFileUploaderDropzone"] small {
    color: #2E5797 !important;
}

/* ── Main area text inputs ───────────────────────────────── */
[data-testid="stTextInput"] input {
    color: #1A1A2E;
    border-radius: 4px;
    font-size: 0.88rem;
}
[data-testid="stTextInput"] input:focus {
    border-color: #2E5797;
    box-shadow: 0 0 0 2px rgba(46, 87, 151, 0.15);
}

/* ── Progress bar ────────────────────────────────────────── */
[data-testid="stProgressBar"] > div > div {
    background-color: #2E5797 !important;
}

/* ── Dataframe ───────────────────────────────────────────── */
[data-testid="stDataFrame"] th {
    background-color: #1F3864 !important;
    color: #FFFFFF !important;
    font-weight: 600;
    font-size: 0.82rem;
}

/* ── Hide Streamlit chrome ───────────────────────────────── */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
"""


def inject_css() -> None:
    """Inject the CUSTOM_CSS block into the Streamlit app."""
    st.markdown(f"<style>{CUSTOM_CSS}</style>", unsafe_allow_html=True)
