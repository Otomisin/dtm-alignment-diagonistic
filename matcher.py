# ============================================================
# DTM Survey Alignment Tool — matcher.py
# Python port of match_surveys() from R (v10)
# Dependencies: pandas, rapidfuzz
# ============================================================

import re
from typing import Callable, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd
from rapidfuzz import process as rf_process
from rapidfuzz.distance import Jaro, JaroWinkler

# ────────────────────────────────────────────────────────────
# Defaults
# ────────────────────────────────────────────────────────────

DEFAULT_STRUCTURAL_TYPES = [
    "audit", "begin_repeat", "begin_group", "calculate",
    "end_repeat", "end_group", "note", "text", "trigger",
    "start", "end", "today",
]

STRING_DISTANCE_SCORERS: Dict[str, Callable] = {
    "Jaro-Winkler": JaroWinkler.distance,
    "Jaro": Jaro.distance,
}

# ────────────────────────────────────────────────────────────
# Public utility — column resolver
# ────────────────────────────────────────────────────────────

# Patterns tried in order when the requested column isn't found.
# Each is a regex matched against every column name in the dataframe.
_FALLBACK_PATTERNS: List[str] = [
    # any language variant: QuestionText(en), QuestionText(fr), …
    r"QuestionText\(",
    r"QuestionText",      # no parentheses at all
    r"Question.*Text",    # looser: QuestionnaireText, etc.
]


def resolve_col(
    requested: Optional[str],
    df: pd.DataFrame,
    df_label: str = "dataframe",
    required: bool = True,
) -> Optional[str]:
    """
    Return the best available column name for `requested`.

    Resolution order:
      1. Exact match                          → use as-is
      2. Walk _FALLBACK_PATTERNS              → return first hit
      3. required=True  → raise ValueError with available columns listed
         required=False → return None silently

    The returned value is always a real column name in `df`
    (or None if not required and not found).
    """
    if requested is None:
        return None

    # 1. Exact match
    if requested in df.columns:
        return requested

    # 2. Fallback patterns — only applied to QuestionText-style columns
    #    so we don't accidentally resolve unrelated columns.
    if re.search(r"QuestionText", requested, re.IGNORECASE):
        for pattern in _FALLBACK_PATTERNS:
            matches = [c for c in df.columns if re.search(
                pattern, c, re.IGNORECASE)]
            if matches:
                # first match wins (typically alphabetical from Excel)
                return matches[0]

    # 3. Not found
    if not required:
        return None

    available = ", ".join(f'"{c}"' for c in df.columns)
    raise ValueError(
        f'Column "{requested}" not found in {df_label}.\n'
        f'Available columns: {available}'
    )


# ────────────────────────────────────────────────────────────
# Internal helpers
# ────────────────────────────────────────────────────────────

def _safe(val) -> str:
    """Convert any value to a stripped string; NaN / None → ''."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()


def _parse_component_cell(cell_text: str) -> List[Dict]:
    """
    Parse a newline-separated QuestionComponent cell like:
        Baseline Sub-Area Assessment: Core
        Baseline Sub-Area Assessment: Discouraged
    into a list of {'comp': ..., 'cat': ...} dicts.
    """
    text = _safe(cell_text)
    if not text:
        return []
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = line.split(":", 1)
        rows.append({
            "comp": parts[0].strip(),
            "cat":  parts[1].strip() if len(parts) == 2 else "",
        })
    return rows


def _find_component_match(
    cell_text: str,
    search_comp,           # str or list[str]
    search_cats=None,      # str or list[str] or None
) -> Optional[str]:
    """
    Return the matched category string if the component cell contains an
    entry matching search_comp (and optionally search_cats), else None.
    Mirrors R's .find_component_match().
    """
    parsed = _parse_component_cell(cell_text)
    if not parsed:
        return None

    if isinstance(search_comp, str):
        search_comp = [search_comp]
    comp_pat = "|".join(re.escape(c) for c in search_comp)

    candidates = [
        r for r in parsed
        if re.search(comp_pat, r["comp"], re.IGNORECASE)
    ]
    if not candidates:
        return None

    if search_cats:
        if isinstance(search_cats, str):
            search_cats = [search_cats]
        cat_pat = "|".join(re.escape(c) for c in search_cats)
        candidates = [
            r for r in candidates
            if re.search(cat_pat, r["cat"], re.IGNORECASE)
        ]

    return candidates[0]["cat"] if candidates else None


# ────────────────────────────────────────────────────────────
# Main matching function
# ────────────────────────────────────────────────────────────

def match_surveys(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    # ── names (for output column prefixes) ──────────────────
    df1_name: str = "survey",
    df2_name: str = "Datakit",
    # ── column mappings ──────────────────────────────────────
    df1_key_col: str = "name",
    df2_key_col: str = "FieldName",
    df1_text_col=None,         # e.g. "label"
    df2_text_col=None,         # e.g. "QuestionText(en)"
    df2_id_col=None,         # e.g. "FieldUniqueId"
    df1_type_col="type",
    df2_type_col="QuestionAnswerType",
    df2_component_col="QuestionComponent",
    # ── matching behaviour ───────────────────────────────────
    structural_types=None,
    fuzzy_threshold: float = 0.20,
    # ── component / category logic ───────────────────────────
    formcomponents=None,   # e.g. "Baseline Sub-Area Assessment"
    question_category=None,   # e.g. "Core"  (informational only)
    missing_question_category=None,   # e.g. ["Core"]
    string_distance_algorithm: str = "Jaro",
    component_match_components=None,   # additional names for ComponentMatch check
    # ── progress callback ────────────────────────────────────
    progress_cb=None,   # callable(msg: str, pct: float)
) -> pd.DataFrame:
    """
    Match df1 (survey form) against df2 (Global Data Kit) in three stages:
      1. Exact key-field match
      2. Exact text-field match
      3. Fuzzy key / text match using the selected string distance algorithm

    Returns df1 with appended diagnostic columns plus any missing-core
    df2 rows appended at the bottom (AlignmentStatus = "Missing*Question").
    """

    # ── defaults / guards ────────────────────────────────────
    if structural_types is None:
        structural_types = DEFAULT_STRUCTURAL_TYPES
    struct_set = {s.lower() for s in structural_types}
    string_distance_scorer = STRING_DISTANCE_SCORERS[
        string_distance_algorithm
    ]

    def _nonnull(x):
        return x if x else None

    df1_text_col = _nonnull(df1_text_col)
    df2_text_col = _nonnull(df2_text_col)
    df2_id_col = _nonnull(df2_id_col)
    df1_type_col = _nonnull(df1_type_col)
    df2_type_col = _nonnull(df2_type_col)
    df2_component_col = _nonnull(df2_component_col)

    use_text = bool(df1_text_col and df2_text_col)
    use_id = bool(df2_id_col)
    use_type = bool(df1_type_col)
    use_type2 = bool(df2_type_col)
    use_component = bool(df2_component_col)
    use_comp_match = use_component and bool(formcomponents)
    use_missing_append = use_comp_match and bool(missing_question_category)

    # ── column resolution + validation ───────────────────────
    # resolve_col() returns the exact column that will be used,
    # falling back to language-variant alternatives for text cols.
    df1_key_col = resolve_col(df1_key_col,       df1, df1_name, required=True)
    df1_text_col = resolve_col(
        df1_text_col,      df1, df1_name, required=False)
    df1_type_col = resolve_col(
        df1_type_col,      df1, df1_name, required=False)
    df2_key_col = resolve_col(df2_key_col,       df2, df2_name, required=True)
    df2_text_col = resolve_col(
        df2_text_col,      df2, df2_name, required=False)
    df2_id_col = resolve_col(df2_id_col,        df2, df2_name, required=False)
    df2_type_col = resolve_col(
        df2_type_col,      df2, df2_name, required=False)
    df2_component_col = resolve_col(
        df2_component_col, df2, df2_name, required=False)

    # ── vectorise inputs ─────────────────────────────────────
    keys1 = [_safe(v) for v in df1[df1_key_col]]
    keys2 = [_safe(v) for v in df2[df2_key_col]]
    texts1 = [_safe(v) for v in df1[df1_text_col]
              ] if use_text else [""] * len(keys1)
    texts2 = [_safe(v) for v in df2[df2_text_col]] if use_text else []
    types1 = [_safe(v).lower() for v in df1[df1_type_col]
              ] if use_type else [""] * len(keys1)
    types2 = [_safe(v) for v in df2[df2_type_col]] if use_type2 else []
    ids2 = [_safe(v) for v in df2[df2_id_col]] if use_id else []
    comp2 = [_safe(v) for v in df2[df2_component_col]] if use_component else []

    n1 = len(keys1)
    matching = ["No Match"] * n1
    matched_k = [None] * n1
    matched_id = [None] * n1
    matched_t = [None] * n1
    matched_row = [None] * n1   # int index into df2

    def _record(i, idx, label):
        matching[i] = label
        matched_k[i] = keys2[idx]
        matched_row[i] = idx
        if use_id:
            matched_id[i] = ids2[idx]
        if use_text:
            matched_t[i] = texts2[idx]

    # ── Stage 1: exact key ───────────────────────────────────
    if progress_cb:
        progress_cb("Stage 1: Exact key matching…", 0.10)

    k2_map: Dict[str, int] = {}
    for idx, k in enumerate(keys2):
        if k and k not in k2_map:
            k2_map[k] = idx
    for i, k1 in enumerate(keys1):
        if k1 and k1 in k2_map:
            _record(i, k2_map[k1], "100% (Key Field Match)")

    # ── Stage 2: exact text ──────────────────────────────────
    if progress_cb:
        progress_cb("Stage 2: Exact text matching…", 0.25)

    if use_text:
        t2_map: Dict[str, int] = {}
        for idx, t in enumerate(texts2):
            if t and t not in t2_map:
                t2_map[t] = idx
        for i, t1 in enumerate(texts1):
            if matching[i] == "No Match" and t1 and t1 in t2_map:
                _record(i, t2_map[t1], "100% (Text Field Match)")

    # ── Stage 3: fuzzy (Jaro-Winkler) ───────────────────────
    if progress_cb:
        progress_cb("Stage 3: Fuzzy matching (Jaro-Winkler)…", 0.45)

    still_unmatched = [i for i in range(n1) if matching[i] == "No Match"]

    if still_unmatched and fuzzy_threshold > 0:
        # Pre-build valid-choice arrays (skip empty strings)
        valid_k2_idx = [idx for idx, k in enumerate(keys2) if k]
        valid_k2_val = [keys2[idx] for idx in valid_k2_idx]
        valid_t2_idx = [idx for idx, t in enumerate(
            texts2) if t] if use_text else []
        valid_t2_val = [texts2[idx]
                        for idx in valid_t2_idx] if use_text else []

        for i in still_unmatched:
            k1 = keys1[i]
            t1 = texts1[i] if use_text else ""

            # Best key match
            best_k = None
            if k1 and valid_k2_val:
                hit = rf_process.extractOne(
                    k1,
                    valid_k2_val,
                    scorer=string_distance_scorer,
                    score_cutoff=fuzzy_threshold,
                )
                # hit = (choice, distance, index_into_valid_k2_val) or None
                if hit and hit[1] < fuzzy_threshold:
                    best_k = (hit[1], valid_k2_idx[hit[2]])

            # Best text match
            best_t = None
            if t1 and valid_t2_val:
                hit = rf_process.extractOne(
                    t1,
                    valid_t2_val,
                    scorer=string_distance_scorer,
                    score_cutoff=fuzzy_threshold,
                )
                if hit and hit[1] < fuzzy_threshold:
                    best_t = (hit[1], valid_t2_idx[hit[2]])

            if not best_k and not best_t:
                continue

            # Prefer key match when distances are equal; otherwise take closer
            if best_k and (not best_t or best_k[0] <= best_t[0]):
                dist, idx = best_k
                _record(i, idx,
                        f"{round((1 - dist) * 100)}% (Fuzzy Key: {keys2[idx]})")
            else:
                dist, idx = best_t
                disp = texts2[idx]
                if len(disp) > 35:
                    disp = disp[:35] + "…"
                _record(i, idx,
                        f"{round((1 - dist) * 100)}% (Fuzzy Text: {disp})")

    # ── AlignmentStatus ──────────────────────────────────────
    if progress_cb:
        progress_cb("Computing alignment status…", 0.65)

    alignment_status: List[str] = []
    for i in range(n1):
        t = types1[i]
        m = matching[i]
        if use_type and t and t in struct_set:
            alignment_status.append("DoesNotNeed2Align")
        elif "100%" in m:
            alignment_status.append("LikelyAligned")
        elif "Fuzzy" in m:
            alignment_status.append("NeedsVerification")
        else:
            alignment_status.append("FPReview")

    # ── DiscouragedQuestion override ─────────────────────────
    if use_comp_match:
        fc = [formcomponents] if isinstance(
            formcomponents, str) else list(formcomponents)
        for i in range(n1):
            if matched_row[i] is None:
                continue
            if use_type and types1[i] in struct_set:
                continue
            cell = comp2[matched_row[i]]
            if _find_component_match(cell, fc, "Discouraged") is not None:
                alignment_status[i] = "DiscouragedQuestion"

    # ── ComponentMatch ───────────────────────────────────────
    component_match: list = [None] * n1
    if use_comp_match:
        fc = [formcomponents] if isinstance(
            formcomponents, str) else list(formcomponents)
        cm_names = fc + (list(component_match_components)
                         if component_match_components else [])
        cm_pat = "|".join(re.escape(n) for n in cm_names)
        for i in range(n1):
            if matched_row[i] is not None:
                cell = comp2[matched_row[i]]
                component_match[i] = (
                    "Yes" if re.search(cm_pat, cell, re.IGNORECASE) else "No"
                )

    # ── Assemble result DataFrame ────────────────────────────
    if progress_cb:
        progress_cb("Assembling result…", 0.80)

    result = df1.copy()
    result["Matching"] = matching
    result["AlignmentStatus"] = alignment_status
    if use_comp_match:
        result["ComponentMatch"] = component_match

    out_key = f"{df2_name}_{df2_key_col}"
    out_id = f"{df2_name}_{df2_id_col}" if use_id else None
    out_text = f"{df2_name}_{df2_text_col}" if use_text else None
    out_type2 = f"{df2_name}_{df2_type_col}" if use_type2 else None
    out_comp = f"{df2_name}_{df2_component_col}" if use_component else None

    result[out_key] = matched_k
    if use_id:
        result[out_id] = matched_id
    if use_text:
        result[out_text] = matched_t
    if use_type2:
        result[out_type2] = [
            types2[matched_row[i]] if matched_row[i] is not None else None
            for i in range(n1)
        ]
    if use_component:
        result[out_comp] = [
            comp2[matched_row[i]] if matched_row[i] is not None else None
            for i in range(n1)
        ]

    result[".RowSource"] = "df1"

    # ── Append missing df2 questions ─────────────────────────
    if use_missing_append:
        if progress_cb:
            progress_cb("Finding missing core questions…", 0.90)

        fc = [formcomponents] if isinstance(
            formcomponents, str) else list(formcomponents)
        mq_cats = (
            [missing_question_category]
            if isinstance(missing_question_category, str)
            else list(missing_question_category)
        )
        matched_idx_set = {
            matched_row[i] for i in range(n1) if matched_row[i] is not None
        }
        append_rows = []

        for j in range(len(df2)):
            if j in matched_idx_set:
                continue
            cell = comp2[j]
            if not cell:
                continue
            matched_cat = _find_component_match(cell, fc, mq_cats)
            if matched_cat is None:
                continue

            row: dict = {col: None for col in result.columns}
            row[df1_key_col] = keys2[j]
            if use_text and df1_text_col:
                row[df1_text_col] = texts2[j]
            row["Matching"] = "Missing from df1"
            cat_label = matched_cat[0].upper() + matched_cat[1:]
            row["AlignmentStatus"] = f"Missing{cat_label}Question"
            if use_comp_match:
                row["ComponentMatch"] = "Yes"
            row[out_key] = keys2[j]
            if use_id:
                row[out_id] = ids2[j]
            if use_text:
                row[out_text] = texts2[j]
            if use_type2:
                row[out_type2] = types2[j]
            if use_component:
                row[out_comp] = cell
            row[".RowSource"] = "df2_missing"
            append_rows.append(row)

        if append_rows:
            appended = pd.DataFrame(append_rows, columns=result.columns)
            result = pd.concat([result, appended], ignore_index=True)

    if progress_cb:
        progress_cb("Done.", 1.0)

    return result
