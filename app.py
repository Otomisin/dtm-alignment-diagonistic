# ============================================================
# DTM Survey Alignment Tool — app.py
# IOM DTM | Streamlit front-end
# ============================================================
# Run locally:   streamlit run app.py
# Deploy:        push to GitHub → connect repo in Streamlit Cloud
# Default Datakit: place file at  data/datakit.xlsx
# ============================================================

from pathlib import Path

import pandas as pd
import streamlit as st

from config import (PAGE_CONFIG, DEFAULTS, FORM_COMPONENTS, STATUS_COLOURS_HEX,
                    STATUS_FONT_HEX, STATUS_EMOJI, inject_css)
from matcher import match_surveys, resolve_col
from exporter import export_alignment_excel

# ────────────────────────────────────────────────────────────
# Page setup  (must be first Streamlit call)
# ────────────────────────────────────────────────────────────

st.set_page_config(**PAGE_CONFIG)
inject_css()

DEFAULT_DATAKIT_PATH = Path("data/datakit.xlsx")

_CUSTOM_SENTINEL = "✏️ Custom…"
_NONE_SENTINEL = "— None (skip) —"


def _column_options(columns, default, allow_none=True):
    """
    Pure logic (no Streamlit calls, so it's unit-testable): build the
    dropdown option list and the index to preselect for a column-mapping
    field, given the real columns detected from a file and the configured
    default.
    """
    options = ([_NONE_SENTINEL] if allow_none else []) + \
        list(columns) + [_CUSTOM_SENTINEL]
    if default in columns:
        index = options.index(default)
    elif not default and allow_none:
        index = options.index(_NONE_SENTINEL)
    else:
        index = len(options) - 1  # Custom…
    return options, index


def _column_picker(label, columns, default, key, allow_none=True):
    """
    Render a column-mapping field.

    If `columns` (the real column names detected from the uploaded file)
    is available, render a dropdown of those columns plus a "Custom…"
    escape hatch. Otherwise (no file loaded yet) fall back to a plain
    free-text input pre-filled with `default`.
    """
    if columns:
        options, index = _column_options(columns, default, allow_none)
        choice = st.selectbox(label, options=options,
                               index=index, key=f"{key}_select")
        if choice == _CUSTOM_SENTINEL:
            return st.text_input(
                f"{label} (custom)",
                value=default if default not in columns else "",
                key=f"{key}_custom",
            )
        if choice == _NONE_SENTINEL:
            return ""
        return choice
    return st.text_input(label, value=default, key=f"{key}_text")


# ────────────────────────────────────────────────────────────
# Session state initialisation
# ────────────────────────────────────────────────────────────

for key, default in [
    ("result", None),
    ("excel_buf", None),
    ("last_survey_name", None),
    ("missing_categories_used", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ────────────────────────────────────────────────────────────
# Sidebar — Datakit + configuration
# ────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🗺️ DTM Alignment Tool")
    st.caption("IOM Displacement Tracking Matrix")
    st.divider()

    # ── Datakit source ──────────────────────────────────────
    st.markdown('<p class="sidebar-section">📂 Datakit</p>',
                unsafe_allow_html=True)

    has_default = DEFAULT_DATAKIT_PATH.exists()

    if has_default:
        use_default_dk = st.toggle(
            "Use stored Datakit",
            value=True,
            help=f"Loads `{DEFAULT_DATAKIT_PATH}` from the repository",
        )
    else:
        use_default_dk = False
        st.info(
            "ℹ️ No stored Datakit found.  \n"
            "Place your file at `data/datakit.xlsx` to enable auto-load.",
            icon=None,
        )

    uploaded_dk = None
    datakit_sheet = DEFAULTS["datakit_sheet"]

    if not use_default_dk:
        uploaded_dk = st.file_uploader(
            "Upload Datakit (.xlsx)",
            type=["xlsx"],
            key="dk_uploader",
        )
        datakit_sheet = st.text_input(
            "Datakit sheet name",
            value=DEFAULTS["datakit_sheet"],
        )
    st.divider()

    # ── Matching configuration ──────────────────────────────
    st.markdown('<p class="sidebar-section">⚙️ Configuration</p>',
                unsafe_allow_html=True)

    _default_component_idx = (
        FORM_COMPONENTS.index(DEFAULTS["formcomponents"])
        if DEFAULTS["formcomponents"] in FORM_COMPONENTS else 0
    )
    formcomponents = st.selectbox(
        "Component name",
        options=FORM_COMPONENTS,
        index=_default_component_idx,
        help="Component to match in the QuestionComponent column",
    )

    fuzzy_threshold = st.slider(
        "Fuzzy threshold (Jaro distance)",
        min_value=0.05,
        max_value=0.40,
        value=DEFAULTS["fuzzy_threshold"],
        step=0.01,
        help=(
            "Jaro *distance* threshold. \n"
            "Lower = stricter. Default 0.20 is roughly >= 80% similarity."
        ),
    )

    st.markdown(
        "Flag missing questions in:",
        help=(
            "A Datakit question tagged with this component that's absent "
            "from the uploaded survey is reported as \"missing\" — but only "
            "for the categories checked here. Core is on by default; check "
            "Optional/Recommended too if you also want those gaps surfaced."
        ),
    )
    _flag_core = st.checkbox(
        "Core", value="Core" in DEFAULTS["missing_categories"])
    _flag_recommended = st.checkbox(
        "Recommended", value="Recommended" in DEFAULTS["missing_categories"])
    _flag_optional = st.checkbox(
        "Optional", value="Optional" in DEFAULTS["missing_categories"])
    missing_categories = [
        cat for cat, checked in [
            ("Core", _flag_core),
            ("Recommended", _flag_recommended),
            ("Optional", _flag_optional),
        ] if checked
    ]

    # ── Column mappings ──────────────────────────────────────
    # The survey file itself is only uploaded further down in the main
    # area (later in script order), but Streamlit restores widget state
    # into session_state at the start of every rerun — so on any run
    # after a file's been uploaded, we can already "peek" at it here via
    # its widget key, before its own file_uploader call runs below.
    _survey_file_peek = st.session_state.get("survey_uploader")
    _survey_sheet_peek = st.session_state.get(
        "survey_sheet_input", DEFAULTS["survey_sheet"])
    survey_columns: list = []
    if _survey_file_peek is not None:
        try:
            _peek_survey = pd.read_excel(
                _survey_file_peek, sheet_name=_survey_sheet_peek, nrows=0)
            _survey_file_peek.seek(0)
            survey_columns = _peek_survey.columns.tolist()
        except Exception:
            survey_columns = []

    _dk_src = DEFAULT_DATAKIT_PATH if use_default_dk else uploaded_dk
    dk_columns: list = []
    if _dk_src is not None:
        try:
            _peek_dk = pd.read_excel(
                _dk_src, sheet_name=datakit_sheet, nrows=0)
            if hasattr(_dk_src, "seek"):
                _dk_src.seek(0)
            dk_columns = _peek_dk.columns.tolist()
        except Exception:
            dk_columns = []

    with st.expander("🗂 Column mappings"):
        st.caption(
            "Pick the matching column from your file, or choose "
            f"\"{_CUSTOM_SENTINEL}\" to type one in."
        )
        df1_key_col = _column_picker(
            "Survey — key field", survey_columns, DEFAULTS["df1_key_col"],
            "df1_key_col", allow_none=False)
        df1_text_col = _column_picker(
            "Survey — label field", survey_columns, DEFAULTS["df1_text_col"],
            "df1_text_col")
        df1_type_col = _column_picker(
            "Survey — type field", survey_columns, DEFAULTS["df1_type_col"],
            "df1_type_col")
        st.markdown("---")
        df2_key_col = _column_picker(
            "Datakit — key field", dk_columns, DEFAULTS["df2_key_col"],
            "df2_key_col", allow_none=False)
        df2_text_col = _column_picker(
            "Datakit — question text", dk_columns, DEFAULTS["df2_text_col"],
            "df2_text_col")
        df2_id_col = _column_picker(
            "Datakit — unique ID", dk_columns, DEFAULTS["df2_id_col"],
            "df2_id_col")
        df2_type_col = _column_picker(
            "Datakit — answer type", dk_columns, DEFAULTS["df2_type_col"],
            "df2_type_col")
        df2_comp_col = _column_picker(
            "Datakit — component column", dk_columns, DEFAULTS["df2_comp_col"],
            "df2_comp_col")


# ────────────────────────────────────────────────────────────
# Main area — upload + run
# ────────────────────────────────────────────────────────────

st.markdown('<p class="dtm-title">🗺️ DTM Survey Alignment Tool</p>',
            unsafe_allow_html=True)
st.markdown(
    '<p class="dtm-sub">Match survey forms against the Global Data Kit · IOM Displacement Tracking Matrix</p>',
    unsafe_allow_html=True,
)
st.divider()

col_upload, col_run = st.columns([3, 1], gap="large")

with col_upload:
    st.subheader("Upload survey form")
    survey_file = st.file_uploader(
        "Upload survey form (.xlsx)",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Expects a sheet called 'survey' by default (configurable below)",
        key="survey_uploader",
    )
    survey_sheet = st.text_input(
        "Survey sheet name",
        value=DEFAULTS["survey_sheet"],
        help="Name of the sheet containing the survey questions",
        key="survey_sheet_input",
    )

with col_run:
    st.subheader("Run")
    run_clicked = st.button(
        "▶ Run Alignment",
        type="primary",
        use_container_width=True,
        disabled=(survey_file is None),
    )

# `survey_columns` / `dk_columns` / `_dk_src` were already computed in
# the sidebar (needed there to populate the Column mappings dropdowns);
# reused here for the preview panel too.

# ── Column preview (shown once files are loaded) ─────────────
if survey_file is not None or not use_default_dk and uploaded_dk is not None:
    with st.expander("🔎 Preview column names — check before running", expanded=False):
        prev_col1, prev_col2 = st.columns(2)

        with prev_col1:
            st.markdown("**Survey form columns**")
            if survey_columns:
                st.dataframe(
                    pd.DataFrame({"Column name": survey_columns}),
                    use_container_width=True, hide_index=True, height=220,
                )
            elif survey_file is not None:
                st.warning("Could not preview this file.")
            else:
                st.caption("Upload a survey file to preview.")

        with prev_col2:
            st.markdown("**Datakit columns**")
            if dk_columns:
                st.dataframe(
                    pd.DataFrame({"Column name": dk_columns}),
                    use_container_width=True, hide_index=True, height=220,
                )
            elif _dk_src is not None:
                st.warning("Could not preview this file.")
            else:
                st.caption("Upload or enable stored Datakit to preview.")


# ────────────────────────────────────────────────────────────
# Execution logic
# ────────────────────────────────────────────────────────────

if run_clicked:
    # Validate inputs
    errors = []
    if survey_file is None:
        errors.append("Upload a survey form.")
    if not use_default_dk and uploaded_dk is None:
        errors.append("Upload a Datakit file or enable the stored default.")
    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # Load files
    with st.spinner("Loading files…"):
        try:
            survey_df = pd.read_excel(survey_file, sheet_name=survey_sheet)
        except Exception as exc:
            st.error(f"❌ Could not read survey form: {exc}")
            st.stop()

        try:
            dk_source = DEFAULT_DATAKIT_PATH if use_default_dk else uploaded_dk
            dk_df = pd.read_excel(dk_source, sheet_name=datakit_sheet)
        except Exception as exc:
            st.error(f"❌ Could not read Datakit: {exc}")
            st.stop()

    # ── Resolve text columns (handles language variants) ─────
    # e.g. QuestionText(en) → QuestionText(fr) when only French exists
    _resolved_dk_text = resolve_col(
        df2_text_col or None, dk_df, "Datakit", required=False)
    _resolved_survey_text = resolve_col(
        df1_text_col or None, survey_df, "survey", required=False)

    if df2_text_col and _resolved_dk_text and _resolved_dk_text != df2_text_col:
        st.warning(
            f'⚠️ Datakit text column **"{df2_text_col}"** not found — '
            f'using **"{_resolved_dk_text}"** instead. '
            f'You can update the configured name in **Column mappings** in the sidebar.',
            icon=None,
        )
    if df1_text_col and _resolved_survey_text and _resolved_survey_text != df1_text_col:
        st.warning(
            f'⚠️ Survey text column **"{df1_text_col}"** not found — '
            f'using **"{_resolved_survey_text}"** instead.',
            icon=None,
        )

    # Run matching
    progress_bar = st.progress(0, text="Starting…")

    def _on_progress(msg: str, pct: float):
        progress_bar.progress(min(pct, 1.0), text=msg)

    try:
        result = match_surveys(
            df1=survey_df,
            df2=dk_df,
            df1_name="survey",
            df2_name="Datakit",
            df1_key_col=df1_key_col,
            df2_key_col=df2_key_col,
            df1_text_col=df1_text_col or None,
            df2_text_col=df2_text_col or None,
            df2_id_col=df2_id_col or None,
            df1_type_col=df1_type_col or None,
            df2_type_col=df2_type_col or None,
            df2_component_col=df2_comp_col or None,
            fuzzy_threshold=fuzzy_threshold,
            formcomponents=formcomponents,
            missing_question_category=missing_categories or None,
            progress_cb=_on_progress,
        )
    except Exception as exc:
        progress_bar.empty()
        st.error(f"❌ Matching failed:\n\n```\n{exc}\n```")
        st.caption("💡 Open the **Preview column names** panel above to check exact column names, then update **Column mappings** in the sidebar.")
        st.stop()

    # Generate Excel in memory
    _on_progress("Generating Excel output…", 0.97)
    excel_buf = export_alignment_excel(
        result_df=result,
        df1_name="survey",
        df2_name="Datakit",
        df1_key_col=df1_key_col,
        df1_text_col=df1_text_col,
        formcomponents=formcomponents,
    )

    progress_bar.progress(1.0, text="Complete!")
    progress_bar.empty()

    # Persist to session state so re-runs don't reset results
    st.session_state.result = result
    st.session_state.excel_buf = excel_buf
    st.session_state.last_survey_name = survey_file.name
    st.session_state.missing_categories_used = missing_categories


# ────────────────────────────────────────────────────────────
# Results display (persists across re-runs via session_state)
# ────────────────────────────────────────────────────────────

if st.session_state.result is not None:
    result = st.session_state.result
    excel_buf = st.session_state.excel_buf

    st.success(
        f"✅ Alignment complete — **{st.session_state.last_survey_name}**",
        icon=None,
    )
    st.divider()

    # ── Summary metrics ─────────────────────────────────────
    st.subheader("📊 Summary")

    is_df1 = result[".RowSource"] == "df1"
    n_df1 = int(is_df1.sum())
    vc = result["AlignmentStatus"].value_counts()

    def _s(key: str) -> int:
        return int(vc.get(key, 0))

    n_missing_appended = int((~is_df1).sum())
    _missing_cats_used = st.session_state.missing_categories_used or []
    _missing_label = (
        "🔴 Missing (" + ", ".join(_missing_cats_used) + ")"
        if _missing_cats_used else "🔴 Missing"
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total survey rows",      n_df1)
    c2.metric("✅ Likely Aligned",       _s("LikelyAligned"))
    c3.metric("🔍 Needs Verification",   _s("NeedsVerification"))
    c4.metric("⚠️ FP Review",            _s("FPReview"))
    c5.metric("🚫 Discouraged",          _s("DiscouragedQuestion"))
    c6.metric(_missing_label,           n_missing_appended)

    # ── Alignment table ─────────────────────────────────────
    st.divider()
    st.subheader("Alignment Table")

    # Determine which columns to display
    survey_show_cols = [
        c for c in [df1_key_col, df1_text_col, df1_type_col]
        if c and c in result.columns
    ]
    dk_show_cols = [c for c in result.columns if c.startswith("Datakit_")]
    display_cols = (
        ["AlignmentStatus"] + survey_show_cols + ["Matching"] + dk_show_cols
    )
    display_cols = [c for c in display_cols if c in result.columns]

    view_df = result[display_cols].copy()

    # Colour the AlignmentStatus column via pandas Styler
    def _colour_status(val):
        if not isinstance(val, str):
            return ""
        if val.startswith("Missing"):
            bg = STATUS_COLOURS_HEX["MissingCoreQuestion"]
            fg = STATUS_FONT_HEX["MissingCoreQuestion"]
        else:
            bg = STATUS_COLOURS_HEX.get(val, "")
            fg = STATUS_FONT_HEX.get(val, "")
        if bg:
            return f"background-color: {bg}; color: {fg}; font-weight: 600"
        return ""

    styled = view_df.style.map(_colour_status, subset=["AlignmentStatus"])
    st.dataframe(styled, use_container_width=True, height=420)

    # ── Filter view ─────────────────────────────────────────
    with st.expander("🔎 Filter by status"):
        all_statuses = sorted(
            result["AlignmentStatus"].dropna().unique().tolist())
        selected = st.multiselect(
            "Show rows with status:",
            options=all_statuses,
            default=all_statuses,
        )
        filtered = result[result["AlignmentStatus"].isin(
            selected)][display_cols]
        st.dataframe(
            filtered.style.map(_colour_status, subset=["AlignmentStatus"]),
            use_container_width=True,
            height=350,
        )
        st.caption(f"{len(filtered):,} rows shown")

    # ── Download ─────────────────────────────────────────────
    st.divider()
    st.subheader("💾 Download")

    col_dl, col_info = st.columns([1, 2])
    with col_dl:
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_buf,
            file_name="Survey_Alignment_Diagnostic.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
            use_container_width=True,
        )
    with col_info:
        n_total = len(result)
        n_miss = n_missing_appended
        st.caption(
            f"**3 sheets:** Summary · Survey_Alignment_Diagnostics "
            f"({n_total} rows: {n_df1} survey + {n_miss} missing) "
            f"· Missing_Questions"
        )

else:
    # Placeholder when no results yet
    st.info(
        "Upload a survey form and click **▶ Run Alignment** to begin.",
        icon="👆",
    )
