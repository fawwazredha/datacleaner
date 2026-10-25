"""
Local Auto Data Cleaner & Power BI-Style Analytics Dashboard

Install dependencies:
    pip install streamlit pandas openpyxl xlsxwriter plotly scikit-learn
    pip install rapidfuzz          # optional: faster fuzzy matching (falls back to difflib)

Run:
    streamlit run datacleaningweb.py
"""

import csv
import difflib
import io
import json
import math
import operator
import os
import re
import unicodedata
from collections import Counter
from functools import reduce

import numpy as np
import pandas as pd
import streamlit as st
from streamlit.errors import StreamlitAPIException

# Plotly & Scikit-Learn imports
import plotly.express as px
import plotly.graph_objects as go
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier

# Optional dependency for fuzzy matching
try:
    from rapidfuzz import fuzz as rf_fuzz
    from rapidfuzz import process as rf_process
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    rf_fuzz = None
    rf_process = None
    RAPIDFUZZ_AVAILABLE = False

st.set_page_config(page_title="Auto Data Cleaner & BI Dashboard", page_icon="🧹", layout="wide")

# =====================================================================
# CUSTOM MODERN CSS STYLING
# =====================================================================
st.markdown("""
<style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }
    [data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    @media (prefers-color-scheme: dark) {
        [data-testid="stMetric"] {
            background-color: #1e222d;
            border: 1px solid #2e3440;
        }
    }
    .streamlit-expanderHeader {
        font-weight: 600;
        font-size: 1.05rem;
        border-radius: 8px;
    }
    div.stButton > button:first-child {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    div.stButton > button[key*="btn_process"] {
        background-color: #2563eb;
        color: white;
        border: none;
        padding: 0.6rem 1.2rem;
        font-size: 1.1rem;
    }
    div.stButton > button[key*="btn_process"]:hover {
        background-color: #1d4ed8;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# CONSTANTS
# =====================================================================
MAX_HISTORY = 20
EXCEL_MAX_ROWS = 1_048_575

FMT_CURRENCY = "Currency / Number (Strip RM, $, commas)"
FMT_DATE_PRESET = "Date (Standard Preset)"
FMT_DATE_CUSTOM = "Date (Custom Manual Format)"
FMT_INT = "Integer (Whole Number)"
FMT_FLOAT = "Float (Decimal Number)"
FMT_TEXT = "Text / String"
FMT_CATEGORY = "Category"
FORMAT_TYPES = [
    FMT_CURRENCY, FMT_DATE_PRESET, FMT_DATE_CUSTOM,
    FMT_INT, FMT_FLOAT, FMT_TEXT, FMT_CATEGORY,
]

DATE_PRESETS = {
    "YYYY-MM-DD": "%Y-%m-%d",
    "DD/MM/YYYY": "%d/%m/%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "DD-MMM-YYYY": "%d-%b-%Y",
}

CASE_NONE = "No Change"
CASE_LOWER = "All Lowercase"
CASE_UPPER = "All UPPERCASE"
CASE_TITLE = "Title Case"

MISSING_NONE = "No Action"
MISSING_DROP = "Drop Rows"
MISSING_IMPUTE = "Impute (Mean/Mode)"
MISSING_CUSTOM = "Fill with Custom Value"

DELIMITERS = {
    "Whitespace (Space)": None,
    "Comma (,)": ",",
    "Semicolon (;)": ";",
    "Pipe (|)": "|",
    "Tab": "\t",
    "Hyphen (-)": "-",
    "Slash (/)": "/",
    "Dot (.)": ".",
    "Custom": "__custom__",
}

MODE_EXACT = "Exact Match"
MODE_CONTAINS = "Contains"
MODE_REGEX = "Regex"

OP_EQ = "equals (=)"
OP_NE = "does not equal (≠)"
OP_CONTAINS = "contains"
OP_NOT_CONTAINS = "does not contain"
OP_STARTS = "starts with"
OP_ENDS = "ends with"
OP_GT = "greater than (>)"
OP_GE = "greater than or equal (>=)"
OP_LT = "less than (<)"
OP_LE = "less than or equal (<=)"
OP_BLANK = "is blank"
OP_NOT_BLANK = "is not blank"
FILTER_OPS = [
    OP_EQ, OP_NE, OP_CONTAINS, OP_NOT_CONTAINS, OP_STARTS, OP_ENDS,
    OP_GT, OP_GE, OP_LT, OP_LE, OP_BLANK, OP_NOT_BLANK,
]
NO_VALUE_OPS = {OP_BLANK, OP_NOT_BLANK}
CMP_OPS = {OP_GT: operator.gt, OP_GE: operator.ge, OP_LT: operator.lt, OP_LE: operator.le}

ACTION_KEEP = "Keep rows matching condition"
ACTION_REMOVE = "Remove rows matching condition"

OUTLIER_IQR = "IQR (Interquartile Range)"
OUTLIER_Z = "Z-Score"
OUTLIER_CLIP = "Clip / Cap values to boundaries"
OUTLIER_DROP = "Drop rows containing outliers"

MATH_OPS = {
    "Addition (+)": operator.add,
    "Subtraction (-)": operator.sub,
    "Multiplication (*)": operator.mul,
    "Division (/)": operator.truediv,
}
ROUND_OPTIONS = ["No rounding", 0, 1, 2, 3, 4, 6]

DURATION_UNITS = ["Days", "Months", "Years", "Hours"]
DURATION_WHOLE = "Whole units (complete only)"
DURATION_DECIMAL = "Decimal (fractional)"
AVG_DAYS_PER_MONTH = 30.4375
AVG_DAYS_PER_YEAR = 365.25

FUZZY_MAX_UNIQUE = 5000 if RAPIDFUZZ_AVAILABLE else 1500

EMOJI_RE = re.compile(
    "["
    "\U0001F000-\U0001F2FF\U0001F300-\U0001FAFF\U0001F1E6-\U0001F1FF"
    "\u2300-\u23FF\u2600-\u27BF\u2B00-\u2BFF\u3030\u303D\u3297\u3299"
    "\u200D\u20E3\uFE0E\uFE0F"
    "]+"
)


# =====================================================================
# UI HELPERS & SESSION STATE MANAGEMENT
# =====================================================================
def stretch(func, *args, **kwargs):
    for extra in ({"width": "stretch"}, {"use_container_width": True}, {}):
        try:
            return func(*args, **kwargs, **extra)
        except (TypeError, StreamlitAPIException):
            continue
    return func(*args, **kwargs)


def init_state():
    defaults = {
        "df": None,
        "original_df": None,
        "data_sig": None,
        "sheets_sig": None,
        "sheet_names": [],
        "history": [],
        "steps": [],
        "lineage": {},
        "ver": 0,
        "cleaned_df": None,
        "csv_bytes": None,
        "xlsx_bytes": None,
        "xlsx_error": None,
        "process_stats": None,
        "process_warnings": [],
        "flash": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_outputs():
    for key in ("cleaned_df", "csv_bytes", "xlsx_bytes", "xlsx_error", "process_stats"):
        st.session_state[key] = None
    st.session_state["process_warnings"] = []


def load_new_dataframe(raw_df, sig):
    df = raw_df.copy()
    df.columns = [str(c) for c in df.columns]
    st.session_state["original_df"] = df
    st.session_state["df"] = df.copy()
    st.session_state["data_sig"] = sig
    st.session_state["history"] = []
    st.session_state["steps"] = []
    st.session_state["lineage"] = {c: c for c in df.columns}
    st.session_state["ver"] += 1
    st.session_state["flash"] = None
    clear_outputs()


def push_history():
    hist = st.session_state["history"]
    hist.append(
        {
            "df": st.session_state["df"].copy(),
            "steps": list(st.session_state["steps"]),
            "lineage": dict(st.session_state["lineage"]),
        }
    )
    if len(hist) > MAX_HISTORY:
        del hist[0]


def commit(new_df, step_msg, new_lineage):
    push_history()
    st.session_state["df"] = new_df.reset_index(drop=True)
    st.session_state["steps"].append(step_msg)
    st.session_state["lineage"] = new_lineage
    st.session_state["ver"] += 1
    clear_outputs()
    st.session_state["flash"] = ("success", f"✅ {step_msg}")
    st.rerun()


def undo_last():
    hist = st.session_state["history"]
    if not hist:
        return
    snap = hist.pop()
    st.session_state["df"] = snap["df"]
    st.session_state["steps"] = snap["steps"]
    st.session_state["lineage"] = snap["lineage"]
    st.session_state["ver"] += 1
    clear_outputs()
    st.session_state["flash"] = ("info", "↩️ Last action undone.")
    st.rerun()


def reset_all():
    original = st.session_state["original_df"]
    if original is None:
        return
    st.session_state["df"] = original.copy()
    st.session_state["steps"] = []
    st.session_state["history"] = []
    st.session_state["lineage"] = {c: c for c in original.columns}
    st.session_state["ver"] += 1
    clear_outputs()
    st.session_state["flash"] = ("info", "🔄 Data reset to original uploaded file.")
    st.rerun()


# =====================================================================
# FILE READING & DATA OPERATIONS
# =====================================================================
def read_csv_safely(uploaded_file):
    last_err = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            uploaded_file.seek(0)
            try:
                return pd.read_csv(uploaded_file, sep=None, engine="python", encoding=enc)
            except (csv.Error, pd.errors.ParserError):
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc)
        except UnicodeDecodeError as err:
            last_err = err
            continue
    raise ValueError(f"Failed to read CSV file encoding: {last_err}")


def unique_name(base, taken):
    name = str(base)
    i = 2
    while name in taken:
        name = f"{base}_{i}"
        i += 1
    taken.add(name)
    return name


def op_rename(df, edited, lineage):
    mapping = {}
    for _, row in edited.iterrows():
        old = row["Current Name"]
        raw_new = row["New Name"]
        new = "" if pd.isna(raw_new) else str(raw_new).strip()
        if new == "":
            raise ValueError(f"New column name for `{old}` cannot be empty.")
        if new != old:
            mapping[old] = new
    if not mapping:
        raise ValueError("No column names were modified.")

    new_cols = [mapping.get(c, c) for c in df.columns]
    dupes = sorted({c for c in new_cols if new_cols.count(c) > 1})
    if dupes:
        raise ValueError(f"Duplicate column names detected: {', '.join(dupes)}")

    out = df.rename(columns=mapping)
    new_lineage = {mapping.get(k, k): v for k, v in lineage.items()}
    msg = "Rename: " + ", ".join(f"`{o}` → `{n}`" for o, n in mapping.items())
    return out, msg, new_lineage


def op_drop(df, cols, lineage):
    if not cols:
        raise ValueError("Please select at least one column to drop.")
    if len(cols) >= df.shape[1]:
        raise ValueError("Cannot drop all columns.")
    out = df.drop(columns=cols)
    new_lineage = {k: v for k, v in lineage.items() if k not in cols}
    return out, "Drop columns: " + ", ".join(f"`{c}`" for c in cols), new_lineage


def compute_split(df, col, delim_label, custom_delim, max_split, strip_parts):
    delim = DELIMITERS[delim_label]
    if delim == "__custom__":
        if not custom_delim:
            raise ValueError("Please enter a custom delimiter.")
        delim = custom_delim
    s = df[col].astype("string")
    n = int(max_split) if int(max_split) > 0 else -1
    if delim is None:
        parts = s.str.split(None, n=n, expand=True)
    else:
        parts = s.str.split(delim, n=n, expand=True, regex=False)
    if parts.shape[1] <= 1:
        raise ValueError(f"Delimiter was not found in column `{col}`.")
    if strip_parts:
        parts = parts.apply(lambda c: c.str.strip())
    return parts


def op_split(df, col, delim_label, custom_delim, max_split, names_text, keep_original, strip_parts, lineage):
    parts = compute_split(df, col, delim_label, custom_delim, max_split, strip_parts)
    k = parts.shape[1]

    user_names = [x.strip() for x in names_text.split(",") if x.strip()] if names_text else []
    if len(user_names) > k:
        raise ValueError(f"You specified {len(user_names)} names, but only {k} parts were generated.")

    taken = set(df.columns)
    if not keep_original:
        taken.discard(col)
    final_names = []
    for i in range(k):
        base = user_names[i] if i < len(user_names) else f"{col}_{i + 1}"
        final_names.append(unique_name(base, taken))

    out = df.copy()
    pos = out.columns.get_loc(col) + 1
    for i, name in enumerate(final_names):
        out.insert(pos + i, name, parts.iloc[:, i].astype(object).values)
    new_lineage = dict(lineage)
    for name in final_names:
        new_lineage[name] = None
    if not keep_original:
        out = out.drop(columns=[col])
        new_lineage.pop(col, None)

    new_lineage = {c: new_lineage.get(c) for c in out.columns}
    msg = f"Split `{col}` → " + ", ".join(f"`{n}`" for n in final_names)
    return out, msg, new_lineage


def op_replace(df, cols, find, repl, mode, case_sensitive, lineage):
    if not cols:
        raise ValueError("Please select at least one column.")
    if mode != MODE_EXACT and find == "":
        raise ValueError("'Find' text cannot be empty for this mode.")
    if mode == MODE_REGEX:
        try:
            re.compile(find)
        except re.error as err:
            raise ValueError(f"Invalid Regex expression: {err}")

    out = df.copy()
    total = 0
    touched = []
    for c in cols:
        orig = out[c]
        s = orig.astype("string")

        if mode == MODE_EXACT:
            if find == "":
                mask = orig.isna() | (s.str.strip() == "")
            elif case_sensitive:
                mask = s == find
            else:
                mask = s.str.lower() == find.lower()
        elif mode == MODE_CONTAINS:
            mask = s.str.contains(find, case=case_sensitive, regex=False)
        else:
            mask = s.str.contains(find, case=case_sensitive, regex=True)
        mask = mask.fillna(False).astype(bool)

        n = int(mask.sum())
        if n == 0:
            continue

        if mode == MODE_EXACT:
            new_vals = pd.Series([repl] * len(out), index=out.index, dtype=object)
        elif mode == MODE_CONTAINS:
            safe_repl = repl.replace("\\", "\\\\")
            new_vals = s.str.replace(re.escape(find), safe_repl, case=case_sensitive, regex=True)
        else:
            new_vals = s.str.replace(find, repl, case=case_sensitive, regex=True)

        res = orig.astype(object)
        res[mask] = new_vals[mask]
        if pd.api.types.is_numeric_dtype(orig.dtype):
            try:
                res = pd.to_numeric(res)
            except (ValueError, TypeError):
                pass
        out[c] = res
        total += n
        touched.append(c)

    if total == 0:
        raise ValueError("No matches found — no changes applied.")
    shown_find = find if find != "" else "(blank)"
    msg = f"Find & Replace '{shown_find}' → '{repl}' ({total} cells across {len(touched)} columns)"
    return out, msg, dict(lineage)


def _bool(series):
    return series.fillna(False).astype(bool)


def to_datetime_safe(series, dayfirst=True):
    try:
        return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst, format="mixed")
    except (TypeError, ValueError):
        return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def build_mask(df, col, op, value, case_sensitive):
    series = df[col]
    s = series.astype("string")
    blank = _bool(series.isna() | (s.str.strip() == ""))

    if op == OP_BLANK:
        return blank
    if op == OP_NOT_BLANK:
        return ~blank

    raw = "" if value is None else str(value)
    if raw.strip() == "":
        raise ValueError("Please provide a value for the filter condition.")

    if op in CMP_OPS:
        func = CMP_OPS[op]
        raw = raw.strip()
        if pd.api.types.is_datetime64_any_dtype(series):
            target = pd.to_datetime(raw, errors="coerce", dayfirst=True)
            if pd.isna(target):
                raise ValueError(f"Value '{raw}' is not a valid date.")
            left = series
        else:
            try:
                target = float(raw.replace(",", ""))
                left = pd.to_numeric(series, errors="coerce")
            except ValueError:
                target = pd.to_datetime(raw, errors="coerce", dayfirst=True)
                if pd.isna(target):
                    raise ValueError(f"Value '{raw}' is not a valid number or date.")
                left = to_datetime_safe(series, True)
        return _bool(pd.Series(func(left, target), index=df.index))

    if op in (OP_EQ, OP_NE):
        m = None
        if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
            try:
                m = pd.to_numeric(series, errors="coerce") == float(raw.replace(",", "").strip())
            except ValueError:
                m = None
        if m is None:
            m = (s == raw) if case_sensitive else (s.str.lower() == raw.lower())
        m = _bool(pd.Series(m, index=df.index))
        return m if op == OP_EQ else ~m

    cmp_s = s if case_sensitive else s.str.lower()
    v = raw if case_sensitive else raw.lower()
    if op == OP_CONTAINS:
        return _bool(cmp_s.str.contains(v, regex=False))
    if op == OP_NOT_CONTAINS:
        return ~_bool(cmp_s.str.contains(v, regex=False))
    if op == OP_STARTS:
        return _bool(cmp_s.str.startswith(v))
    if op == OP_ENDS:
        return _bool(cmp_s.str.endswith(v))
    raise ValueError(f"Unrecognized operator: {op}")


def combine_masks(df, rules, use_and, case_sensitive):
    masks = [build_mask(df, c, o, v, case_sensitive) for c, o, v in rules]
    return reduce(operator.and_ if use_and else operator.or_, masks)


def op_filter(df, rules, use_and, keep, case_sensitive, lineage):
    combined = combine_masks(df, rules, use_and, case_sensitive)
    out = df[combined] if keep else df[~combined]
    if len(out) == len(df):
        raise ValueError("This filter does not alter any rows.")
    if out.empty:
        raise ValueError("This filter would remove ALL rows — operation aborted.")
    joiner = " AND " if use_and else " OR "
    desc = joiner.join(f"`{c}` {o}" + ("" if o in NO_VALUE_OPS else f" '{v}'") for c, o, v in rules)
    verb = "Keep" if keep else "Remove"
    msg = f"Filter ({verb}): {desc} → {len(df)} ➜ {len(out)} rows"
    return out, msg, dict(lineage)


def to_number(series):
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    txt = series.astype("string").str.strip()
    neg = txt.str.match(r"^\(.*\)$").fillna(False).astype(bool)
    cleaned = txt.str.replace(r"[^\d.\-]", "", regex=True).fillna("").astype(object)
    num = pd.to_numeric(cleaned, errors="coerce")
    return num.where(~neg, -num.abs())


def apply_format_rule(series, rule_type, config):
    if rule_type == FMT_CURRENCY:
        dec = int(config.get("decimals", 2))
        num = to_number(series).round(dec)
        if dec == 0:
            num = num.astype("Int64")
        return num

    if rule_type in (FMT_DATE_PRESET, FMT_DATE_CUSTOM):
        if rule_type == FMT_DATE_PRESET:
            target_fmt = DATE_PRESETS.get(config.get("date_fmt"), "%Y-%m-%d")
        else:
            target_fmt = str(config.get("custom_date_str", "")).strip()
            if not target_fmt:
                raise ValueError("Custom date format string is empty.")
            pd.Timestamp("2024-01-31 13:45:00").strftime(target_fmt)
        dt = to_datetime_safe(series, bool(config.get("dayfirst", True)))
        return dt.dt.strftime(target_fmt)

    if rule_type == FMT_INT:
        return to_number(series).round().astype("Int64")

    if rule_type == FMT_FLOAT:
        return to_number(series).round(int(config.get("decimals", 2)))

    if rule_type == FMT_TEXT:
        return series.map(lambda v: v if pd.isna(v) else str(v))

    if rule_type == FMT_CATEGORY:
        return series.astype("category")

    raise ValueError(f"Unrecognized format type: {rule_type}")


# =====================================================================
# ADVANCED TRANSFORMATIONS (Helpers)
# =====================================================================
def as_float(series):
    num = to_number(series)
    try:
        return num.astype("float64")
    except (TypeError, ValueError):
        return pd.Series(num.to_numpy(dtype="float64", na_value=np.nan), index=num.index)


def count_changed(old, new):
    changed = 0
    for a, b in zip(old.tolist(), new.tolist()):
        if pd.isna(a) and pd.isna(b):
            continue
        if a != b:
            changed += 1
    return changed


def strip_accents(text):
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def make_text_normalizer(opts, blank_as_missing):
    keep = str(opts.get("keep", "") or "")
    special_re = None
    if opts.get("special"):
        allowed = r"\w\s" + re.escape(keep)
        pattern = rf"[^{allowed}]" if "_" in keep else rf"[^{allowed}]|_"
        special_re = re.compile(pattern)
    ws_re = re.compile(r"\s+")

    def normalize(value):
        if not isinstance(value, str):
            return value
        text = value
        if opts.get("accents"):
            text = strip_accents(text)
        if opts.get("emoji"):
            text = EMOJI_RE.sub("", text)
        if special_re is not None:
            text = special_re.sub("", text)
        if opts.get("emoji") or special_re is not None:
            text = ws_re.sub(" ", text).strip()
        if blank_as_missing and text.strip() == "":
            return np.nan
        return text

    return normalize


def apply_text_normalization(out, tcfg, blank_as_missing, warnings, log):
    cols = [c for c in tcfg.get("cols", []) if c in out.columns]
    if not cols or not (tcfg.get("accents") or tcfg.get("emoji") or tcfg.get("special")):
        return out
    normalizer = make_text_normalizer(tcfg, blank_as_missing)
    parts = []
    if tcfg.get("accents"):
        parts.append("accents")
    if tcfg.get("emoji"):
        parts.append("emojis")
    if tcfg.get("special"):
        parts.append("special characters")
    for col in cols:
        try:
            old = out[col]
            new = old.map(normalizer)
            changed = count_changed(old, new)
            out[col] = new
            log.append(f"Text normalization `{col}`: removed {', '.join(parts)} ({changed} cells changed).")
        except Exception as err:
            warnings.append(f"Text normalization failed for column `{col}`: {err}")
    return out


def fuzzy_key(text):
    text = strip_accents(str(text)).lower()
    text = re.sub(r"[^\w\s]", " ", text).replace("_", " ")
    return re.sub(r"\s+", " ", text).strip()


def fuzzy_group_values(counts, threshold):
    values = sorted(counts.keys(), key=lambda v: (-counts[v], -len(v), v))
    keys = [fuzzy_key(v) for v in values]
    pending = [i for i in range(len(values)) if keys[i] != ""]
    mapping, clusters = {}, []

    while pending:
        i = pending.pop(0)
        canon, canon_key = values[i], keys[i]
        members = []
        if pending:
            if RAPIDFUZZ_AVAILABLE:
                cand_keys = [keys[j] for j in pending]
                hits = rf_process.extract(
                    canon_key, cand_keys, scorer=rf_fuzz.ratio, score_cutoff=threshold, limit=None
                )
                members = sorted({hit[2] for hit in hits})
                members = [pending[pos] for pos in members]
            else:
                matcher = difflib.SequenceMatcher(None, autojunk=False)
                matcher.set_seq2(canon_key)
                for j in pending:
                    matcher.set_seq1(keys[j])
                    if matcher.real_quick_ratio() * 100 < threshold:
                        continue
                    if matcher.quick_ratio() * 100 < threshold:
                        continue
                    if matcher.ratio() * 100 >= threshold:
                        members.append(j)
        if members:
            member_set = set(members)
            pending = [j for j in pending if j not in member_set]
            variants = [values[j] for j in members]
            for variant in variants:
                mapping[variant] = canon
            clusters.append((canon, variants, sum(counts[v] for v in variants)))
    return mapping, clusters


def fuzzy_cluster_series(series, threshold):
    counts = Counter(v for v in series.tolist() if isinstance(v, str) and v != "")
    if len(counts) > FUZZY_MAX_UNIQUE:
        raise ValueError(
            f"{len(counts)} distinct values exceed the safety limit of {FUZZY_MAX_UNIQUE} for fuzzy grouping."
        )
    return fuzzy_group_values(dict(counts), threshold)


def apply_fuzzy(out, fcfg, warnings, log):
    cols = [c for c in fcfg.get("cols", []) if c in out.columns]
    threshold = float(fcfg.get("threshold", 85))
    for col in cols:
        try:
            mapping, clusters = fuzzy_cluster_series(out[col], threshold)
            if not mapping:
                log.append(f"Fuzzy grouping `{col}`: no similar values found at {threshold:.0f}% threshold.")
                continue
            out[col] = out[col].map(lambda v, m=mapping: m.get(v, v) if isinstance(v, str) else v)
            cells = sum(c[2] for c in clusters)
            log.append(
                f"Fuzzy grouping `{col}`: merged {len(mapping)} variant(s) into {len(clusters)} "
                f"group(s) ({cells} cells updated) at {threshold:.0f}% threshold."
            )
        except Exception as err:
            warnings.append(f"Fuzzy grouping failed for column `{col}`: {err}")
    return out


def math_operand(out, operand, warnings):
    if operand["kind"] == "column":
        col = operand["value"]
        if col not in out.columns:
            raise ValueError(f"Column `{col}` was not found.")
        before = int(out[col].notna().sum())
        num = as_float(out[col])
        lost = before - int(num.notna().sum())
        if lost > 0:
            warnings.append(f"Column math: {lost} non-numeric value(s) in `{col}` were treated as NaN.")
        return num
    raw = str(operand["value"]).replace(",", "").strip()
    try:
        return float(raw)
    except ValueError:
        raise ValueError(f"Constant '{operand['value']}' is not a valid number.")


def describe_operand(operand):
    return f"`{operand['value']}`" if operand["kind"] == "column" else str(operand["value"]).strip()


def apply_math_rules(out, rules, warnings, log):
    for idx, rule in enumerate(rules, start=1):
        try:
            if rule["left"]["kind"] == "constant" and rule["right"]["kind"] == "constant":
                raise ValueError("At least one operand must be a column.")
            name = str(rule.get("name", "")).strip()
            if not name:
                raise ValueError("The new column name is empty.")
            left = math_operand(out, rule["left"], warnings)
            right = math_operand(out, rule["right"], warnings)
            with np.errstate(all="ignore"):
                result = pd.Series(MATH_OPS[rule["op"]](left, right), index=out.index, dtype="float64")
            inf_count = int(np.isinf(result).sum())
            if inf_count:
                result = result.replace([np.inf, -np.inf], np.nan)
                warnings.append(
                    f"Column math #{idx}: {inf_count} infinite result(s) (division by zero) were set to NaN."
                )
            if rule.get("round") not in (None, "No rounding"):
                result = result.round(int(rule["round"]))
            final_name = unique_name(name, set(out.columns))
            out[final_name] = result
            symbol = rule["op"].split("(")[-1].rstrip(")")
            note = "" if final_name == name else f" (renamed from `{name}` because the name already existed)"
            log.append(
                f"Column math: `{final_name}` = {describe_operand(rule['left'])} {symbol} "
                f"{describe_operand(rule['right'])}{note}."
            )
        except Exception as err:
            warnings.append(f"Column math #{idx} failed: {err}")
    return out


def parse_dates_naive(series, dayfirst):
    parsed = to_datetime_safe(series, dayfirst)
    try:
        if getattr(parsed.dt, "tz", None) is not None:
            parsed = parsed.dt.tz_localize(None)
    except AttributeError:
        pass
    return parsed


def whole_months_between(start, end):
    forward = start <= end
    sign = np.where(forward, 1, -1)
    lo = start.where(forward, end)
    hi = end.where(forward, start)
    raw = (hi.dt.year - lo.dt.year) * 12 + (hi.dt.month - lo.dt.month)
    time_hi = hi - hi.dt.normalize()
    time_lo = lo - lo.dt.normalize()
    incomplete = (hi.dt.day < lo.dt.day) | ((hi.dt.day == lo.dt.day) & (time_hi < time_lo))
    return (raw - incomplete.astype(int)) * sign


def compute_duration(start, end, unit, mode, decimals):
    delta = end - start
    whole = mode == DURATION_WHOLE
    if unit in ("Months", "Years") and whole:
        months = whole_months_between(start, end)
        values = months if unit == "Months" else np.trunc(months / 12)
    else:
        seconds = delta.dt.total_seconds()
        if unit == "Hours":
            values = seconds / 3600
        elif unit == "Days":
            values = seconds / 86400
        elif unit == "Months":
            values = seconds / 86400 / AVG_DAYS_PER_MONTH
        else:
            values = seconds / 86400 / AVG_DAYS_PER_YEAR
        if whole:
            values = np.trunc(values)
    values = pd.Series(values, index=start.index, dtype="float64")
    if whole:
        return values.round().astype("Int64")
    return values.round(int(decimals))


def apply_date_rules(out, rules, dayfirst, warnings, log):
    for idx, rule in enumerate(rules, start=1):
        try:
            start_col, end_col = rule["start"], rule["end"]
            for col in (start_col, end_col):
                if col not in out.columns:
                    raise ValueError(f"Column `{col}` was not found.")
            name = str(rule.get("name", "")).strip()
            if not name:
                raise ValueError("The new column name is empty.")

            parsed = {}
            for col in (start_col, end_col):
                before = int(out[col].notna().sum())
                dt = parse_dates_naive(out[col], dayfirst)
                valid = int(dt.notna().sum())
                if valid == 0:
                    raise ValueError(f"No valid dates could be parsed in `{col}`.")
                if before - valid > 0:
                    warnings.append(
                        f"Date duration #{idx}: {before - valid} value(s) in `{col}` are not valid dates (set to NaN)."
                    )
                parsed[col] = dt

            result = compute_duration(
                parsed[start_col], parsed[end_col], rule["unit"], rule["mode"], rule.get("decimals", 2)
            )
            final_name = unique_name(name, set(out.columns))
            out[final_name] = result
            note = "" if final_name == name else f" (renamed from `{name}` because the name already existed)"
            log.append(
                f"Date duration: `{final_name}` = {rule['unit'].lower()} from `{start_col}` to `{end_col}` "
                f"({rule['mode'].split(' (')[0].lower()}){note}."
            )
        except Exception as err:
            warnings.append(f"Date duration #{idx} failed: {err}")
    return out


def outlier_bounds(num, method, k):
    clean = num.dropna()
    if method == OUTLIER_IQR:
        if len(clean) < 4:
            return None
        q1, q3 = clean.quantile(0.25), clean.quantile(0.75)
        iqr = q3 - q1
        return float(q1 - k * iqr), float(q3 + k * iqr)
    if len(clean) < 3:
        return None
    std = clean.std(ddof=0)
    if pd.isna(std) or std == 0:
        return None
    mean = clean.mean()
    return float(mean - k * std), float(mean + k * std)


def apply_outliers(out, ocfg, warnings, log, stats):
    cols = [c for c in ocfg.get("cols", []) if c in out.columns]
    if not cols:
        return out
    method, k, action = ocfg["method"], float(ocfg["k"]), ocfg["action"]

    evaluated = {}
    for col in cols:
        try:
            num = as_float(out[col])
            bounds = outlier_bounds(num, method, k)
            if bounds is None:
                warnings.append(f"Outliers `{col}`: too few numeric values or zero variance — skipped.")
                continue
            evaluated[col] = (num, bounds)
        except Exception as err:
            warnings.append(f"Outlier detection failed for column `{col}`: {err}")
    if not evaluated:
        return out

    method_label = "IQR" if method == OUTLIER_IQR else "Z-Score"
    if action == OUTLIER_DROP:
        drop_mask = pd.Series(False, index=out.index)
        for col, (num, (lo, hi)) in evaluated.items():
            mask = (num < lo) | (num > hi)
            log.append(
                f"Outliers `{col}` ({method_label}, k={k:g}): {int(mask.sum())} outlier(s) outside "
                f"[{lo:.4g}, {hi:.4g}]."
            )
            drop_mask = drop_mask | mask
        before = len(out)
        out = out[~drop_mask]
        stats["outlier_rows_dropped"] = before - len(out)
        log.append(f"Outliers: dropped {before - len(out)} row(s) containing outliers.")
        return out

    for col, (num, (lo, hi)) in evaluated.items():
        try:
            outlier_count = int(((num < lo) | (num > hi)).sum())
            if outlier_count == 0:
                log.append(f"Outliers `{col}` ({method_label}, k={k:g}): none found.")
                continue
            if not pd.api.types.is_numeric_dtype(out[col]):
                warnings.append(
                    f"Outliers `{col}`: column is not numeric after formatting — format it as a number first."
                )
                continue
            if pd.api.types.is_integer_dtype(out[col]):
                lo_i, hi_i = math.ceil(lo), math.floor(hi)
                if lo_i <= hi_i:
                    clipped = num.clip(lower=lo_i, upper=hi_i).round().astype("Int64")
                else:
                    clipped = num.clip(lower=lo, upper=hi)
            else:
                clipped = num.clip(lower=lo, upper=hi)
            out[col] = clipped
            log.append(
                f"Outliers `{col}` ({method_label}, k={k:g}): clipped {outlier_count} value(s) to "
                f"[{lo:.4g}, {hi:.4g}]."
            )
        except Exception as err:
            warnings.append(f"Outlier clipping failed for column `{col}`: {err}")
    return out


def process_data(df, cfg, column_rules):
    out = df.copy()
    warnings = []
    stats = {"dup_removed": 0, "missing_rows_dropped": 0, "outlier_rows_dropped": 0, "advanced_log": []}
    adv = cfg.get("adv", {}) or {}
    adv_log = stats["advanced_log"]

    if cfg["trim"] or cfg["case"] != CASE_NONE or cfg["blank_missing"]:
        ws_re = re.compile(r"\s+")

        def transform(v):
            if not isinstance(v, str):
                return v
            if cfg["trim"]:
                v = ws_re.sub(" ", v).strip()
            if cfg["case"] == CASE_LOWER:
                v = v.lower()
            elif cfg["case"] == CASE_UPPER:
                v = v.upper()
            elif cfg["case"] == CASE_TITLE:
                v = v.title()
            if cfg["blank_missing"] and v.strip() == "":
                return np.nan
            return v

        text_cols = [
            c for c in out.columns
            if out[c].dtype == "object" or isinstance(out[c].dtype, pd.StringDtype)
        ]
        for c in text_cols:
            try:
                out[c] = out[c].map(transform)
            except Exception as err:
                warnings.append(f"Text cleaning failed for column `{c}`: {err}")

    try:
        out = apply_text_normalization(
            out, adv.get("text_norm", {}), cfg.get("blank_missing", False), warnings, adv_log
        )
        out = apply_fuzzy(out, adv.get("fuzzy", {}), warnings, adv_log)
        out = apply_date_rules(
            out, adv.get("durations", []), bool(adv.get("dur_dayfirst", True)), warnings, adv_log
        )
        out = apply_math_rules(out, adv.get("math", []), warnings, adv_log)
    except Exception as err:
        warnings.append(f"Advanced transformations stopped unexpectedly: {err}")

    for col_name, rule in column_rules.items():
        if col_name not in out.columns:
            continue
        try:
            before_na = int(out[col_name].isna().sum())
            out[col_name] = apply_format_rule(out[col_name], rule["type"], rule["config"])
            after_na = int(out[col_name].isna().sum())
            if after_na > before_na:
                warnings.append(
                    f"`{col_name}`: {after_na - before_na} values could not be converted and were set to NaN."
                )
        except Exception as err:
            warnings.append(f"Failed to process column `{col_name}`: {err}")

    try:
        out = apply_outliers(out, adv.get("outliers", {}), warnings, adv_log, stats)
    except Exception as err:
        warnings.append(f"Outlier handling stopped unexpectedly: {err}")

    if cfg["drop_dup"]:
        before = len(out)
        try:
            out = out.drop_duplicates()
            stats["dup_removed"] = before - len(out)
        except Exception as err:
            warnings.append(f"Failed to remove duplicates: {err}")

    missing_rules = cfg.get("column_missing_rules", {})
    if missing_rules:
        for col_name, rule in missing_rules.items():
            if col_name not in out.columns:
                continue

            action = rule["action"]
            custom_val = rule["custom_val"]

            try:
                if action == MISSING_DROP:
                    before_rows = len(out)
                    out = out.dropna(subset=[col_name])
                    stats["missing_rows_dropped"] += (before_rows - len(out))

                elif action == MISSING_IMPUTE:
                    s = out[col_name]
                    if s.isna().any():
                        if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
                            mean_val = s.mean()
                            if not pd.isna(mean_val):
                                val = int(round(mean_val)) if pd.api.types.is_integer_dtype(s) else mean_val
                                out[col_name] = s.fillna(val)
                        else:
                            mode_s = s.mode(dropna=True)
                            val = mode_s.iloc[0] if not mode_s.empty else "Unknown"
                            if isinstance(s.dtype, pd.CategoricalDtype) and val not in s.cat.categories:
                                out[col_name] = s.cat.add_categories([val])
                            out[col_name] = s.fillna(val)

                elif action == MISSING_CUSTOM and custom_val is not None:
                    raw_str = str(custom_val).strip()
                    try:
                        if "." in raw_str:
                            typed_val = float(raw_str)
                        else:
                            typed_val = int(raw_str)
                    except ValueError:
                        typed_val = raw_str

                    if isinstance(out[col_name].dtype, pd.CategoricalDtype):
                        if typed_val not in out[col_name].cat.categories:
                            out[col_name] = out[col_name].cat.add_categories([typed_val])

                    out[col_name] = out[col_name].fillna(typed_val)

            except Exception as err:
                warnings.append(f"Failed handling missing values for column `{col_name}`: {err}")

    return out.reset_index(drop=True), warnings, stats


def safe_dup_count(df):
    try:
        return int(df.duplicated().sum())
    except Exception:
        return 0


def build_audit(original, cleaned, lineage):
    rows = []
    for col in cleaned.columns:
        orig_name = lineage.get(col)
        after_dtype = str(cleaned[col].dtype)
        after_na = int(cleaned[col].isna().sum())
        if orig_name is not None and orig_name in original.columns:
            before_dtype = str(original[orig_name].dtype)
            before_na = int(original[orig_name].isna().sum())
            status = "Changed" if before_dtype != after_dtype else "Unchanged"
        else:
            orig_name = "(new column)"
            before_dtype = "—"
            before_na = None
            status = "New"
        rows.append(
            {
                "Current Column": col,
                "Original Column": orig_name,
                "Data Type (Before)": before_dtype,
                "Data Type (After)": after_dtype,
                "Missing (Before)": before_na,
                "Missing (After)": after_na,
                "Status": status,
            }
        )
    audit = pd.DataFrame(rows)
    audit["Missing (Before)"] = audit["Missing (Before)"].astype("Int64")

    kept = {v for v in lineage.values() if v is not None}
    dropped = [c for c in original.columns if c not in kept]

    summary = pd.DataFrame(
        {
            "Metric": ["Total Rows", "Total Columns", "Total Missing Values", "Duplicate Rows", "Memory (KB)"],
            "Before (Original File)": [
                original.shape[0],
                original.shape[1],
                int(original.isnull().sum().sum()),
                safe_dup_count(original),
                round(float(original.memory_usage(deep=True).sum()) / 1024, 1),
            ],
            "After (Cleaned Data)": [
                cleaned.shape[0],
                cleaned.shape[1],
                int(cleaned.isnull().sum().sum()),
                safe_dup_count(cleaned),
                round(float(cleaned.memory_usage(deep=True).sum()) / 1024, 1),
            ],
        }
    )
    return summary, audit, dropped


def to_excel_bytes(df):
    if len(df) > EXCEL_MAX_ROWS:
        return None, "Row count exceeds Excel maximum limit (1,048,575). Please export as CSV."
    last_err = None
    for engine in ("xlsxwriter", "openpyxl"):
        try:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine=engine) as writer:
                df.to_excel(writer, index=False, sheet_name="Cleaned Data")
                if engine == "xlsxwriter":
                    ws = writer.sheets["Cleaned Data"]
                    ws.freeze_panes(1, 0)
                    if len(df.columns) > 0:
                        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
                    for i, col in enumerate(df.columns):
                        try:
                            longest = int(df[col].astype(str).str.len().max()) if len(df) else 0
                        except Exception:
                            longest = 10
                        ws.set_column(i, i, min(max(len(str(col)), longest) + 2, 50))
            return buf.getvalue(), None
        except Exception as err:
            last_err = err
            continue
    return None, f"Failed to generate Excel file (please install `xlsxwriter` or `openpyxl`): {last_err}"


# =====================================================================
# MAIN APPLICATION SETUP
# =====================================================================
init_state()

st.title("🧹 Local Auto Data Cleaner & BI Dashboard")
st.write("Upload, clean, transform, and analyze your data with an interactive Power BI-style dashboard!")

uploaded_file = st.file_uploader("Choose a CSV or Excel file", type=["csv", "xlsx"], key="up_file")

if uploaded_file is None:
    st.info("⬆️ Please upload a CSV or Excel (.xlsx) file to get started.")
    st.stop()

file_sig = f"{uploaded_file.name}|{uploaded_file.size}"
load_error = None
try:
    if uploaded_file.name.lower().endswith(".csv"):
        if st.session_state["data_sig"] != file_sig:
            load_new_dataframe(read_csv_safely(uploaded_file), file_sig)
    else:
        if st.session_state["sheets_sig"] != file_sig:
            uploaded_file.seek(0)
            st.session_state["sheet_names"] = pd.ExcelFile(uploaded_file).sheet_names
            st.session_state["sheets_sig"] = file_sig
        sheet_names = st.session_state["sheet_names"]
        if len(sheet_names) > 1:
            sheet = st.selectbox("Select Sheet:", sheet_names, key=f"sb_sheet_{file_sig}")
        else:
            sheet = sheet_names[0]
        data_sig = f"{file_sig}|{sheet}"
        if st.session_state["data_sig"] != data_sig:
            uploaded_file.seek(0)
            load_new_dataframe(pd.read_excel(uploaded_file, sheet_name=sheet), data_sig)
except Exception as err:
    load_error = err

if load_error is not None:
    st.error(f"Failed to read file: {load_error}")
    st.stop()

df = st.session_state["df"]
if df is None or df.shape[1] == 0:
    st.error("The uploaded file contains no readable data.")
    st.stop()

ver = st.session_state["ver"]


def K(name):
    return f"{name}_v{ver}"


def operand_input(label, key_prefix, columns, default_kind):
    kinds = ["Column", "Constant"]
    kind = st.radio(
        f"{label} type", kinds, index=kinds.index(default_kind), horizontal=True, key=K(f"{key_prefix}_kind")
    )
    if kind == "Column":
        return {"kind": "column", "value": st.selectbox(f"{label} column", columns, key=K(f"{key_prefix}_col"))}
    return {"kind": "constant", "value": st.text_input(f"{label} constant value", value="1", key=K(f"{key_prefix}_const"))}


flash = st.session_state.get("flash")
if flash:
    getattr(st, flash[0])(flash[1])
    st.session_state["flash"] = None

# ---------------------------------------------------------------------
# WORKING DATA PREVIEW & TOP CONTROLS
# ---------------------------------------------------------------------
st.subheader("📊 Working Data Preview")
stretch(st.dataframe, df, height=300)

num_duplicates = safe_dup_count(df)
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Rows", df.shape[0])
m2.metric("Total Columns", df.shape[1])
m3.metric("Missing Values", int(df.isnull().sum().sum()))
m4.metric(
    label="Duplicate Rows",
    value=num_duplicates,
    delta=f"{num_duplicates} issues" if num_duplicates > 0 else "Clean",
    delta_color="inverse" if num_duplicates > 0 else "normal",
)

u1, u2, u3 = st.columns([1, 1, 3])
with u1:
    if stretch(st.button, "↩️ Undo", key=K("btn_undo"), disabled=len(st.session_state["history"]) == 0):
        undo_last()
with u2:
    if stretch(st.button, "🔄 Reset to Original", key=K("btn_reset")):
        reset_all()
with u3:
    with st.expander(f"📜 Applied Steps ({len(st.session_state['steps'])})", expanded=False):
        if st.session_state["steps"]:
            for i, step in enumerate(st.session_state["steps"], start=1):
                st.write(f"{i}. {step}")
        else:
            st.caption("No steps applied yet. Use the tabs below to modify your data.")

st.markdown("---")

# Navigation Tabs including Power BI Dashboard
tab_cols, tab_fr, tab_fmt, tab_out, tab_bi = st.tabs(
    [
        "🧱 Column Operations",
        "🔎 Find & Replace / Filter",
        "🧼 Cleaning & Format",
        "📋 Audit & Export",
        "📈 Power BI Dashboard",
    ]
)

# =====================================================================
# TAB 1: COLUMN OPERATIONS
# =====================================================================
with tab_cols:
    sub_rename, sub_drop, sub_split = st.tabs(["✏️ Rename", "🗑️ Drop", "✂️ Split Column"])

    with sub_rename:
        st.write("Edit new column names in the **New Name** column, then click Apply.")
        rename_input = pd.DataFrame({"Current Name": list(df.columns), "New Name": list(df.columns)})
        edited_rename = stretch(
            st.data_editor,
            rename_input,
            hide_index=True,
            disabled=["Current Name"],
            num_rows="fixed",
            key=K("de_rename"),
        )
        if stretch(st.button, "✅ Apply Rename", key=K("btn_rename")):
            result = None
            try:
                result = op_rename(df, edited_rename, st.session_state["lineage"])
            except ValueError as err:
                st.error(str(err))
            except Exception as err:
                st.error(f"Unexpected error during rename: {err}")
            if result is not None:
                commit(*result)

    with sub_drop:
        drop_cols = st.multiselect("Select columns to drop:", list(df.columns), key=K("ms_drop"))
        if stretch(st.button, "🗑️ Drop Selected Columns", key=K("btn_drop")):
            result = None
            try:
                result = op_drop(df, drop_cols, st.session_state["lineage"])
            except ValueError as err:
                st.error(str(err))
            except Exception as err:
                st.error(f"Unexpected error during drop: {err}")
            if result is not None:
                commit(*result)

    with sub_split:
        sp1, sp2 = st.columns(2)
        split_col = sp1.selectbox("Column to split:", list(df.columns), key=K("sb_split_col"))
        split_delim_label = sp2.selectbox("Delimiter:", list(DELIMITERS.keys()), key=K("sb_split_delim"))

        split_custom = ""
        if split_delim_label == "Custom":
            split_custom = st.text_input("Enter custom delimiter:", value="", key=K("ti_split_custom"))

        sp3, sp4 = st.columns(2)
        split_max = sp3.number_input(
            "Max split count (0 = split all; 1 = 2 parts):",
            min_value=0, max_value=20, value=0, step=1, key=K("ni_split_max"),
        )
        split_names = sp4.text_input(
            "New column names (comma-separated, optional):",
            value="", placeholder="e.g. First Name, Last Name", key=K("ti_split_names"),
        )
        sp5, sp6 = st.columns(2)
        split_keep = sp5.checkbox("Keep original column", value=False, key=K("chk_split_keep"))
        split_strip = sp6.checkbox("Trim whitespace on split parts", value=True, key=K("chk_split_strip"))

        try:
            preview_parts = compute_split(df, split_col, split_delim_label, split_custom, split_max, split_strip)
            st.caption(f"Preview: {preview_parts.shape[1]} parts will be created (showing first 5 rows).")
            preview_parts.columns = [f"part_{i + 1}" for i in range(preview_parts.shape[1])]
            stretch(st.dataframe, preview_parts.head(5).astype(object))
        except ValueError as err:
            st.caption(f"ℹ️ {err}")
        except Exception as err:
            st.caption(f"ℹ️ Preview unavailable: {err}")

        if stretch(st.button, "✂️ Apply Split", key=K("btn_split")):
            result = None
            try:
                result = op_split(
                    df, split_col, split_delim_label, split_custom, split_max,
                    split_names, split_keep, split_strip, st.session_state["lineage"],
                )
            except ValueError as err:
                st.error(str(err))
            except Exception as err:
                st.error(f"Unexpected error during split: {err}")
            if result is not None:
                commit(*result)

# =====================================================================
# TAB 2: FIND & REPLACE + FILTERING
# =====================================================================
with tab_fr:
    st.markdown("#### 🔁 Find & Replace")
    fr_all = st.checkbox("Apply to all columns", value=False, key=K("chk_fr_all"))
    if fr_all:
        fr_cols = list(df.columns)
    else:
        fr_cols = st.multiselect("Select columns:", list(df.columns), key=K("ms_fr_cols"))

    fr1, fr2 = st.columns(2)
    fr_find = fr1.text_input("Find:", value="", key=K("ti_fr_find"))
    fr_repl = fr2.text_input("Replace with:", value="", key=K("ti_fr_repl"))
    fr_mode = st.radio("Match mode:", [MODE_EXACT, MODE_CONTAINS, MODE_REGEX], horizontal=True, key=K("rd_fr_mode"))
    fr_case = st.checkbox("Match case", value=False, key=K("chk_fr_case"))
    st.caption("💡 Tip: Leave 'Find' empty with 'Exact Match' mode to fill all empty cells with 'Replace with'.")

    if stretch(st.button, "🔁 Apply Replace", key=K("btn_replace")):
        result = None
        try:
            result = op_replace(df, fr_cols, fr_find, fr_repl, fr_mode, fr_case, st.session_state["lineage"])
        except ValueError as err:
            st.error(str(err))
        except Exception as err:
            st.error(f"Unexpected error during replace: {err}")
        if result is not None:
            commit(*result)

    st.markdown("---")
    st.markdown("#### 🧮 Row Filter")
    n_rules = int(st.number_input("Number of conditions:", min_value=1, max_value=5, value=1, step=1, key=K("ni_rules")))
    filter_rules = []
    for i in range(n_rules):
        fc1, fc2, fc3 = st.columns([2, 2, 2])
        f_col = fc1.selectbox(f"Column (rule {i + 1}):", list(df.columns), key=K(f"flt_col_{i}"))
        f_op = fc2.selectbox("Operator:", FILTER_OPS, key=K(f"flt_op_{i}"))
        if f_op in NO_VALUE_OPS:
            f_val = ""
            fc3.caption("No value required.")
        else:
            f_val = fc3.text_input("Value:", value="", key=K(f"flt_val_{i}"))
        filter_rules.append((f_col, f_op, f_val))

    flt_use_and = True
    if n_rules > 1:
        logic_choice = st.radio(
            "Combine conditions with:", ["AND (all conditions)", "OR (any condition)"],
            horizontal=True, key=K("rd_flt_logic"),
        )
        flt_use_and = logic_choice.startswith("AND")
    flt_action = st.radio("Action:", [ACTION_KEEP, ACTION_REMOVE], key=K("rd_flt_action"))
    flt_case = st.checkbox("Match case", value=False, key=K("chk_flt_case"))

    try:
        preview_mask = combine_masks(df, filter_rules, flt_use_and, flt_case)
        st.caption(f"ℹ️ {int(preview_mask.sum())} out of {len(df)} rows match the conditions.")
    except ValueError:
        pass
    except Exception:
        pass

    if stretch(st.button, "🧮 Apply Filter", key=K("btn_filter")):
        result = None
        try:
            result = op_filter(df, filter_rules, flt_use_and, flt_action == ACTION_KEEP, flt_case, st.session_state["lineage"])
        except ValueError as err:
            st.error(str(err))
        except Exception as err:
            st.error(f"Unexpected error during filter: {err}")
        if result is not None:
            commit(*result)

# =====================================================================
# TAB 3: CLEANING & FORMATTING
# =====================================================================
with tab_fmt:
    with st.expander("💾 Recipe Preset System (Save / Load Config)", expanded=False):
        c_rec1, c_rec2 = st.columns(2)
        with c_rec1:
            st.markdown("** Export Current Settings**")
            st.caption("Save your configured cleaning rules to a JSON recipe file.")
        with c_rec2:
            st.markdown("** Import Saved Recipe**")
            uploaded_recipe = st.file_uploader("Upload Recipe File", type=["json"], key=K("up_recipe"))

    with st.expander("📝 1. Text Cleaning & Case Standardization", expanded=False):
        clean_text_spaces = st.checkbox("Trim excess whitespace", value=True, key="chk_spaces")
        blank_as_missing = st.checkbox("Treat empty strings / whitespace-only cells as Missing Value (NaN)", value=True, key="chk_blank_missing")
        text_case = st.selectbox("Change Text Case Format:", [CASE_NONE, CASE_LOWER, CASE_UPPER, CASE_TITLE], key="sb_textcase")

    with st.expander("🎯 2. Header-Specific Formatting (Custom Rules)", expanded=True):
        selected_custom_cols = st.multiselect("Select columns to format:", list(df.columns), key=K("ms_custom_cols"))
        column_rules = {}
        if selected_custom_cols:
            st.markdown("---")
            for col_name in selected_custom_cols:
                st.write(f"📌 **Settings for Header: `{col_name}`**")
                c1, c2, c3 = st.columns([2, 2, 2])
                fmt_type = c1.selectbox("Select Data Format Type:", FORMAT_TYPES, key=K(f"type_{col_name}"))
                extra_config = {}
                if fmt_type in (FMT_CURRENCY, FMT_FLOAT):
                    extra_config["decimals"] = c2.slider("Decimal Places:", 0, 4, 2, key=K(f"dec_{col_name}"))
                elif fmt_type == FMT_DATE_PRESET:
                    extra_config["date_fmt"] = c2.selectbox("Select Date Preset:", list(DATE_PRESETS.keys()), key=K(f"dfmt_{col_name}"))
                    extra_config["dayfirst"] = c3.checkbox("Parse as DD/MM first", value=True, key=K(f"dayfirst_{col_name}"))
                elif fmt_type == FMT_DATE_CUSTOM:
                    extra_config["custom_date_str"] = c2.text_input("Enter custom format string (e.g. %d/%m/%Y %H:%M):", value="%Y-%m-%d", key=K(f"custom_dfmt_{col_name}"))
                    extra_config["dayfirst"] = c3.checkbox("Parse as DD/MM first", value=True, key=K(f"dayfirst_c_{col_name}"))
                column_rules[col_name] = {"type": fmt_type, "config": extra_config}
                st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

    with st.expander("🧩 3. Missing Values & Duplicate Handling", expanded=True):
        st.markdown("#### 👯 Duplicate Rows Handling")
        drop_dup = st.checkbox("Drop Duplicate Rows", value=True, key="chk_dup")
        dup_mask = df.duplicated(keep=False)
        num_dups = int(df.duplicated().sum())
        if num_dups > 0:
            st.warning(f"⚠️ Found **{num_dups} duplicate row(s)** in your dataset.")
            if st.checkbox("🔍 Preview duplicate rows only", value=False, key=K("chk_preview_dups")):
                stretch(st.dataframe, df[dup_mask].sort_values(by=list(df.columns)), height=250)
        else:
            st.success("✅ No duplicate rows detected in the current dataset.")

        st.markdown("---")
        st.markdown("#### 🎯 Missing Values Breakdown & Handling")
        missing_series = df.isnull().sum()
        missing_cols_df = missing_series[missing_series > 0]
        if not missing_cols_df.empty:
            total_rows = len(df)
            info_list = [f"• **`{col_name}`**: {count} missing cells ({(count / total_rows) * 100:.1f}%)" for col_name, count in missing_cols_df.items()]
            st.info(f"📌 **Missing Values Detected in {len(missing_cols_df)} Column(s):**\n\n" + "\n".join(info_list))
        else:
            st.success("✅ No missing values found in any column!")

        default_missing_selection = list(missing_cols_df.index) if not missing_cols_df.empty else []
        selected_missing_cols = st.multiselect("Select column(s) to handle missing values:", list(df.columns), default=default_missing_selection, key=K("ms_missing_cols"))

        column_missing_rules = {}
        if selected_missing_cols:
            for col_name in selected_missing_cols:
                col_null_count = int(df[col_name].isnull().sum())
                st.write(f"📌 **Rule for Column: `{col_name}`** *(Currently {col_null_count} missing)*")
                c1, c2 = st.columns(2)
                action = c1.selectbox(f"Action for `{col_name}`:", [MISSING_DROP, MISSING_IMPUTE, MISSING_CUSTOM], key=K(f"sb_col_action_{col_name}"))
                custom_val = None
                if action == MISSING_CUSTOM:
                    custom_val = c2.text_input(f"Custom replacement value for `{col_name}`:", value="0", key=K(f"ti_col_custom_{col_name}"))
                column_missing_rules[col_name] = {"action": action, "custom_val": custom_val}
                st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)

    with st.expander("🚀 4. Advanced Transformations", expanded=False):
        all_cols = list(df.columns)
        text_cols_all = [c for c in all_cols if df[c].dtype == "object" or isinstance(df[c].dtype, pd.StringDtype)]
        adv_tab_out, adv_tab_math, adv_tab_text, adv_tab_date = st.tabs(["📈 Outliers", "➗ Column Math", "🔤 Text Normalization & Fuzzy", "📅 Date Duration"])

        math_rules = []
        with adv_tab_math:
            n_math = int(st.number_input("Number of calculated columns:", min_value=0, max_value=5, value=0, step=1, key=K("adv_math_n")))
            for i in range(n_math):
                mc1, mc2, mc3 = st.columns([3, 2, 3])
                with mc1:
                    math_left = operand_input("Left operand", f"adv_math_left_{i}", all_cols, "Column")
                math_op = mc2.selectbox("Operator:", list(MATH_OPS.keys()), key=K(f"adv_math_op_{i}"))
                with mc3:
                    math_right = operand_input("Right operand", f"adv_math_right_{i}", all_cols, "Constant")
                mn1, mn2 = st.columns(2)
                math_name = mn1.text_input("New column name:", value=f"calc_{i + 1}", key=K(f"adv_math_name_{i}"))
                math_round = mn2.selectbox("Round result:", ROUND_OPTIONS, format_func=lambda x: x if isinstance(x, str) else f"{x} decimal place(s)", key=K(f"adv_math_round_{i}"))
                math_rules.append({"left": math_left, "op": math_op, "right": math_right, "name": math_name, "round": math_round})

        duration_rules = []
        dur_dayfirst = True
        with adv_tab_date:
            n_dur = int(st.number_input("Number of duration columns:", min_value=0, max_value=5, value=0, step=1, key=K("adv_dur_n")))
            if n_dur > 0:
                dur_dayfirst = st.checkbox("Parse ambiguous dates as DD/MM first", value=True, key=K("adv_dur_dayfirst"))
            for i in range(n_dur):
                dc1, dc2, dc3 = st.columns(3)
                dur_start = dc1.selectbox("Start Date Column:", all_cols, key=K(f"adv_dur_start_{i}"))
                dur_end = dc2.selectbox("End Date Column:", all_cols, index=min(1, len(all_cols) - 1), key=K(f"adv_dur_end_{i}"))
                dur_unit = dc3.selectbox("Unit:", DURATION_UNITS, key=K(f"adv_dur_unit_{i}"))
                dd1, dd2, dd3 = st.columns(3)
                dur_mode = dd1.selectbox("Result type:", [DURATION_WHOLE, DURATION_DECIMAL], key=K(f"adv_dur_mode_{i}"))
                dur_decimals = 2
                if dur_mode == DURATION_DECIMAL:
                    dur_decimals = dd2.slider("Decimal places:", 0, 4, 2, key=K(f"adv_dur_dec_{i}"))
                dur_name = dd3.text_input("New column name:", value=f"duration_{i + 1}", key=K(f"adv_dur_name_{i}"))
                duration_rules.append({"start": dur_start, "end": dur_end, "unit": dur_unit, "mode": dur_mode, "decimals": dur_decimals, "name": dur_name})

        text_norm_cfg = {"cols": [], "accents": False, "emoji": False, "special": False, "keep": ""}
        fuzzy_cfg = {"cols": [], "threshold": 85}
        with adv_tab_text:
            norm_cols = st.multiselect("Columns to normalize:", text_cols_all, key=K("adv_norm_cols"))
            nc1, nc2, nc3 = st.columns(3)
            norm_accents = nc1.checkbox("Remove accents (é → e)", value=False, key=K("adv_norm_accents"))
            norm_emoji = nc2.checkbox("Remove emojis", value=False, key=K("adv_norm_emoji"))
            norm_special = nc3.checkbox("Remove special characters", value=False, key=K("adv_norm_special"))
            norm_keep = st.text_input("Extra characters to keep:", value="", placeholder="e.g. .-@/", key=K("adv_norm_keep")) if norm_special else ""
            text_norm_cfg = {"cols": norm_cols, "accents": norm_accents, "emoji": norm_emoji, "special": norm_special, "keep": norm_keep}

            st.markdown("---")
            fuzzy_cols = st.multiselect("Columns to group similar values:", text_cols_all, key=K("adv_fuzzy_cols"))
            fuzzy_threshold = st.slider("Similarity threshold (%):", 0, 100, 85, 1, key=K("adv_fuzzy_thr"))
            fuzzy_cfg = {"cols": fuzzy_cols, "threshold": fuzzy_threshold}

        outlier_cfg = {"cols": [], "method": OUTLIER_IQR, "k": 1.5, "action": OUTLIER_CLIP}
        with adv_tab_out:
            numeric_formats = (FMT_CURRENCY, FMT_INT, FMT_FLOAT)
            candidates = [c for c in all_cols if pd.api.types.is_numeric_dtype(df[c]) or column_rules.get(c, {}).get("type") in numeric_formats]
            out_cols = st.multiselect("Numeric columns to check:", candidates, key=K("adv_out_cols"))
            oc1, oc2 = st.columns(2)
            out_method = oc1.radio("Detection method:", [OUTLIER_IQR, OUTLIER_Z], key=K("adv_out_method"))
            out_k = oc2.slider("Multiplier / Threshold:", 0.5, 5.0, 1.5 if out_method == OUTLIER_IQR else 3.0, 0.1, key=K("adv_out_k"))
            out_action = st.radio("Action on outliers:", [OUTLIER_CLIP, OUTLIER_DROP], key=K("adv_out_action"))
            outlier_cfg = {"cols": out_cols, "method": out_method, "k": out_k, "action": out_action}

        adv_cfg = {
            "text_norm": text_norm_cfg,
            "fuzzy": fuzzy_cfg,
            "durations": duration_rules,
            "dur_dayfirst": dur_dayfirst,
            "math": math_rules,
            "outliers": outlier_cfg,
        }

    if stretch(st.button, "🚀 Process & Format Data Now", key=K("btn_process")):
        cfg = {
            "trim": clean_text_spaces,
            "blank_missing": blank_as_missing,
            "case": text_case,
            "drop_dup": drop_dup,
            "column_missing_rules": column_missing_rules,
            "adv": adv_cfg,
        }
        process_result = None
        with st.spinner("Processing data..."):
            try:
                process_result = process_data(df, cfg, column_rules)
            except Exception as err:
                st.error(f"Error while processing data: {err}")

        if process_result is not None:
            cleaned_df, proc_warnings, proc_stats = process_result
            csv_bytes = cleaned_df.to_csv(index=False).encode("utf-8")
            xlsx_bytes, xlsx_error = to_excel_bytes(cleaned_df)

            st.session_state["cleaned_df"] = cleaned_df
            st.session_state["csv_bytes"] = csv_bytes
            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.session_state["xlsx_error"] = xlsx_error
            st.session_state["process_stats"] = proc_stats
            st.session_state["process_warnings"] = proc_warnings
            st.success("✅ Data successfully cleaned and formatted! Check the **📋 Audit & Export** or **📈 Power BI Dashboard** tab.")

# =====================================================================
# TAB 4: AUDIT & EXPORT
# =====================================================================
with tab_out:
    cleaned_df = st.session_state.get("cleaned_df")
    if cleaned_df is None:
        st.info("No processed data available yet. Click **🚀 Process & Format Data Now** in the Cleaning & Format tab.")
    else:
        stats = st.session_state.get("process_stats") or {}
        st.success(
            f"✅ Result: {cleaned_df.shape[0]} rows × {cleaned_df.shape[1]} columns "
            f"| Duplicates removed: {stats.get('dup_removed', 0)} "
            f"| Rows dropped (missing): {stats.get('missing_rows_dropped', 0)} "
            f"| Outlier rows dropped: {stats.get('outlier_rows_dropped', 0)}"
        )
        stretch(st.dataframe, cleaned_df, height=350)
        base_name = os.path.splitext(uploaded_file.name)[0]
        d1, d2 = st.columns(2)
        with d1:
            if st.session_state.get("csv_bytes") is not None:
                stretch(st.download_button, "📥 Download Cleaned CSV", st.session_state["csv_bytes"], f"{base_name}_cleaned.csv", "text/csv", key=K("dl_csv"))
        with d2:
            if st.session_state.get("xlsx_bytes") is not None:
                stretch(st.download_button, "📥 Download Cleaned Excel (.xlsx)", st.session_state["xlsx_bytes"], f"{base_name}_cleaned.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=K("dl_xlsx"))

# =====================================================================
# TAB 5: POWER BI-STYLE ANALYTICS DASHBOARD
# =====================================================================
with tab_bi:
    # Guna cleaned_df jika ada, jika tidak guna df terkini
    active_df = st.session_state.get("cleaned_df") if st.session_state.get("cleaned_df") is not None else df
    
    st.markdown("### 📊 Interactive BI Canvas & Analytics Dashboard")
    st.caption("Cross-filtering, AI Key Influencers, Decomposition Treemap, and Smart Narratives powered by Plotly & Scikit-Learn.")
    
    # Identify Data Types for Auto-Slicers & Charts
    numeric_cols = [c for c in active_df.columns if pd.api.types.is_numeric_dtype(active_df[c]) and not pd.api.types.is_bool_dtype(active_df[c])]
    categorical_cols = [c for c in active_df.columns if active_df[c].dtype == "object" or isinstance(active_df[c].dtype, pd.CategoricalDtype) or isinstance(active_df[c].dtype, pd.StringDtype)]
    
    # -------------------------------------------------------------
    # 1. DYNAMIC SLICERS & FILTERS BAR (Power BI Style)
    # -------------------------------------------------------------
    st.markdown("#### 🎛️ Interactive Canvas Slicers")
    filter_container = st.container()
    filtered_df = active_df.copy()
    
    with filter_container:
        f1, f2 = st.columns(2)
        selected_cat_col = f1.selectbox("Filter Dimension (Category):", ["None"] + categorical_cols, key=K("bi_cat_filter_col"))
        
        if selected_cat_col != "None":
            unique_vals = list(active_df[selected_cat_col].dropna().unique())
            selected_vals = f2.multiselect(f"Select values for `{selected_cat_col}`:", unique_vals, default=unique_vals, key=K("bi_cat_filter_vals"))
            if selected_vals:
                filtered_df = filtered_df[filtered_df[selected_cat_col].isin(selected_vals)]
                
    st.markdown("---")
    
    # -------------------------------------------------------------
    # 2. TOP KPI CARDS
    # -------------------------------------------------------------
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("Active Rows Analyzed", len(filtered_df))
    kpi_col2.metric("Total Metrics Count", len(numeric_cols))
    
    if numeric_cols:
        target_kpi_col = st.sidebar.selectbox("🎯 Target Metric for Analysis:", numeric_cols, key=K("bi_kpi_target"))
        kpi_col3.metric(f"Total {target_kpi_col}", f"{filtered_df[target_kpi_col].sum():,.2f}")
        kpi_col4.metric(f"Average {target_kpi_col}", f"{filtered_df[target_kpi_col].mean():,.2f}")
    else:
        target_kpi_col = None
        kpi_col3.metric("Total Metric", "N/A")
        kpi_col4.metric("Average Metric", "N/A")
        
    st.markdown("---")
    
    # -------------------------------------------------------------
    # 3. INTERACTIVE VISUAL CANVAS (Rows of Plotly Charts)
    # -------------------------------------------------------------
    r1_col1, r1_col2 = st.columns(2)
    
    with r1_col1:
        st.markdown("##### 🌳 Decomposition Treemap (Hierarchy Analysis)")
        if categorical_cols and target_kpi_col:
            tree_cats = st.multiselect("Select Hierarchy Levels:", categorical_cols, default=categorical_cols[:min(2, len(categorical_cols))], key=K("bi_tree_cats"))
            if tree_cats:
                fig_tree = px.treemap(filtered_df, path=tree_cats, values=target_kpi_col, color=target_kpi_col, color_continuous_scale="Viridis")
                fig_tree.update_layout(margin=dict(t=10, l=10, r=10, b=10))
                st.plotly_chart(fig_tree, use_container_width=True)
            else:
                st.info("Select at least one category hierarchy level.")
        else:
            st.info("Treemap requires numeric & categorical columns.")

    with r1_col2:
        st.markdown("##### 📈 Distribution & Trend Comparison")
        if numeric_cols:
            dist_col = st.selectbox("Select Numeric Column:", numeric_cols, key=K("bi_dist_col"))
            fig_dist = px.histogram(filtered_df, x=dist_col, marginal="box", color_discrete_sequence=["#2563eb"])
            fig_dist.update_layout(margin=dict(t=10, l=10, r=10, b=10))
            st.plotly_chart(fig_dist, use_container_width=True)
        else:
            st.info("No numeric columns available.")
            
    st.markdown("---")
    
    # -------------------------------------------------------------
    # 4. ADVANCED AI KEY INFLUENCERS (Machine Learning Feature Importance)
    # -------------------------------------------------------------
    st.markdown("##### 🤖 AI Key Influencers (Root-Cause Analysis)")
    st.caption("Uses Machine Learning (Random Forest) to determine which columns drive performance or variance in your target metric.")
    
    if target_kpi_col and len(filtered_df) > 10:
        try:
            # Prepare Data for Machine Learning Model
            X_df = filtered_df.drop(columns=[target_kpi_col]).dropna()
            y = filtered_df.loc[X_df.index, target_kpi_col]
            
            # One-Hot Encode Categoricals
            X_encoded = pd.get_dummies(X_df, drop_first=True)
            
            if not X_encoded.empty and len(X_encoded) > 5:
                rf = RandomForestRegressor(n_estimators=50, random_state=42)
                rf.fit(X_encoded, y)
                
                importances = pd.Series(rf.feature_importances_, index=X_encoded.columns).sort_values(ascending=True).tail(10)
                
                fig_ai = px.bar(
                    x=importances.values,
                    y=importances.index,
                    orientation='h',
                    labels={'x': 'Relative Influence Score', 'y': 'Key Driver / Feature'},
                    color=importances.values,
                    color_continuous_scale="Blues"
                )
                fig_ai.update_layout(margin=dict(t=10, l=10, r=10, b=10))
                st.plotly_chart(fig_ai, use_container_width=True)
            else:
                st.info("Not enough valid feature columns to run AI Influencers.")
        except Exception as e:
            st.warning(f"Could not calculate AI Key Influencers: {e}")
    else:
        st.info("AI Key Influencers require a valid numeric target and at least 10 rows of data.")
        
    st.markdown("---")
    
    # -------------------------------------------------------------
    # 5. SMART NARRATIVE (Natural Language Summary)
    # -------------------------------------------------------------
    st.markdown("##### 📝 Smart Narrative (Executive Summary)")
    
    narrative_text = []
    narrative_text.append(f"• **Dataset Overview**: Current dataset contains **{len(filtered_df):,} active records** after applied filters.")
    
    if target_kpi_col:
        total_val = filtered_df[target_kpi_col].sum()
        avg_val = filtered_df[target_kpi_col].mean()
        max_val = filtered_df[target_kpi_col].max()
        narrative_text.append(f"• **Performance Insight**: The total accumulated **{target_kpi_col}** is **{total_val:,.2f}** with an average of **{avg_val:,.2f}** per entry.")
        narrative_text.append(f"• **Peak Observation**: The highest recorded value for **{target_kpi_col}** reaches **{max_val:,.2f}**.")
        
    if selected_cat_col != "None" and selected_cat_col in filtered_df.columns:
        top_cat = filtered_df[selected_cat_col].mode()
        if not top_cat.empty:
            narrative_text.append(f"• **Dominant Segment**: The most frequent category in **{selected_cat_col}** is **'{top_cat.iloc[0]}'**.")

    st.info("\n".join(narrative_text))