import os
import importlib
import json
import re
import zipfile
from xml.etree import ElementTree as ET
from datetime import date
from datetime import timedelta
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import excel_analysis

st.set_page_config(
    page_title="S관 생산 필요수량 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed",
)


DATA_DIR = "data"
KST = ZoneInfo("Asia/Seoul")
REPO_EXCEL_CANDIDATES = [
    "s관 부족수량.xlsx",
    os.path.join(DATA_DIR, "s관 부족수량.xlsx"),
]
TEMPLATE_XLSX_PATH = "업로드 양식.xlsx"
OUT_DIR = "out"
STREAMLIT_CONFIG_PATH = os.path.join(".streamlit", "config.toml")
DASHBOARD_LINKS_PATH = "dashboard_links.json"


def _today_kst() -> date:
    return datetime.now(tz=KST).date()


def _load_dashboard_links(path: str = DASHBOARD_LINKS_PATH) -> list[dict[str, str]]:
    """
    Load external dashboard links from a local json file.

    Recommended schema:
      [{"label": "...", "url": "https://..."}]

    Also accepts:
      {"links": [...]} and "name" instead of "label".
    """
    try:
        if not os.path.exists(path):
            return []
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("links"), list):
            data = data["links"]
        if not isinstance(data, list):
            return []
        out: list[dict[str, str]] = []
        for item in data:
            if not isinstance(item, dict):
                continue
            label = str(item.get("label") or item.get("name") or "").strip()
            url = str(item.get("url") or "").strip()
            if label and url:
                out.append({"label": label, "url": url})
        return out
    except Exception:
        return []


def _table_height_for_rows(
    n_rows: int,
    *,
    min_height: int,
    max_height: int,
    header_px: int = 110,
    row_px: int = 34,
) -> int:
    n = max(0, int(n_rows))
    h = int(header_px + (n * row_px))
    return max(min_height, min(max_height, h))


def _pre_widget_single_select_fix(*, key: str, default: str, options: list[str]) -> None:
    """
    Safe to call BEFORE the widget is instantiated in the current run.
    Fixes invalid/cleared value in session_state so the widget shows `default`.
    """
    v = st.session_state.get(key)
    if isinstance(v, list):
        v = v[0] if v else None
    if isinstance(v, str):
        v = v.strip()
    if v not in options:
        st.session_state[key] = default


def _on_change_single_select(key: str, default: str, options: list[str]) -> None:
    """
    Callback: safe to mutate session_state for the widget key.
    Used to snap back to default when user clears the selection.
    """
    v = st.session_state.get(key)
    if isinstance(v, list):
        v = v[0] if v else None
    if isinstance(v, str):
        v = v.strip()
    if v not in options:
        st.session_state[key] = default


def _coerce_single_value(value: str | None, *, default: str, options: list[str]) -> str:
    v = (value or "").strip()
    return v if v in options else default


def _find_repo_excel() -> str | None:
    for p in REPO_EXCEL_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def _file_mtime_label(path: str) -> str:
    try:
        def _xlsx_modified_ts(p: str) -> datetime | None:
            if not str(p).lower().endswith(".xlsx"):
                return None
            try:
                with zipfile.ZipFile(p) as zf:
                    core = zf.read("docProps/core.xml")
                root = ET.fromstring(core)
                ns = {
                    "dcterms": "http://purl.org/dc/terms/",
                }
                modified_el = root.find(".//dcterms:modified", ns)
                if modified_el is None or not (modified_el.text or "").strip():
                    return None
                raw = modified_el.text.strip()
                raw = raw.replace("Z", "+00:00")
                dt = datetime.fromisoformat(raw)
                return dt if dt.tzinfo is not None else dt.replace(tzinfo=ZoneInfo("UTC"))
            except Exception:
                return None

        # Prefer Excel's internal "modified" timestamp (prevents git checkout time confusion).
        ts = _xlsx_modified_ts(path)
        if ts is None:
            ts = datetime.fromtimestamp(os.path.getmtime(path), tz=KST)
        else:
            ts = ts.astimezone(KST)
        return ts.strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception:
        return "-"


def _read_bytes(path: str) -> bytes | None:
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None


def _safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def _ensure_latest_outputs(*, excel_path: str, out_dir: str) -> dict:
    due_csv = os.path.join(out_dir, "납기_제품군_공정별부족.csv")
    detail_csv = os.path.join(out_dir, "이니셜별_수주상세.csv")
    excel_mtime = os.path.getmtime(excel_path)

    if os.path.exists(due_csv) and os.path.exists(detail_csv):
        if os.path.getmtime(due_csv) >= excel_mtime and os.path.getmtime(detail_csv) >= excel_mtime:
            return {"ok": True, "regenerated": False, "due_csv": due_csv, "detail_csv": detail_csv}

    _safe_mkdir(out_dir)
    importlib.reload(excel_analysis)
    info = excel_analysis.export_due_process_shortage(file_path=excel_path, out_dir=out_dir)
    if not info.get("enabled"):
        return {"ok": False, "reason": info.get("reason") or "export failed"}
    return {"ok": True, "regenerated": True, "due_csv": due_csv, "detail_csv": detail_csv}


def _load_theme_from_config() -> dict:
    try:
        if not os.path.exists(STREAMLIT_CONFIG_PATH):
            return {}
        with open(STREAMLIT_CONFIG_PATH, "rb") as f:
            import tomllib  # py3.11+

            data = tomllib.load(f)
        return data.get("theme", {}) if isinstance(data, dict) else {}
    except Exception:
        return {}


def _style_dataframe_like_dashboard(df: pd.DataFrame) -> object:
    theme = _load_theme_from_config()
    bg = theme.get("backgroundColor", "#FBF7EE")
    sbg = theme.get("secondaryBackgroundColor", "#F2EBDD")
    text = theme.get("textColor", "#1B1B1B")
    try:
        return (
            df.style.set_properties(**{"background-color": bg, "color": text})
            .set_table_styles(
                [
                    {"selector": "th", "props": [("background-color", sbg), ("color", text)]},
                    {"selector": "td", "props": [("background-color", bg), ("color", text)]},
                ]
            )
        )
    except Exception:
        return df


def _apply_local_theme_css() -> None:
    theme = _load_theme_from_config()
    bg = theme.get("backgroundColor", "#FBF7EE")
    sbg = theme.get("secondaryBackgroundColor", "#F2EBDD")
    text = theme.get("textColor", "#1B1B1B")
    primary = theme.get("primaryColor", "#0A5C36")

    st.markdown(
        f"""
<style>
.stApp {{
  background-color: {bg} !important;
  color: {text} !important;
}}
[data-testid="stSidebar"] > div {{
  background-color: {sbg} !important;
}}
/* Make widgets/containers closer to local look across deployments */
div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stContainer"] {{
  background-color: transparent;
}}
a, a:visited {{
  color: {primary};
}}

/* Air / spacing */
.block-container {{
  padding-top: 2.2rem !important;
  padding-bottom: 2.2rem !important;
}}
div[data-testid="stVerticalBlock"] > div {{
  row-gap: 0.9rem;
}}
/* Pills & segmented controls spacing */
div[data-testid="stPills"] > div {{
  margin-top: 0.2rem;
  margin-bottom: 0.7rem;
}}
div[data-testid="stSegmentedControl"] > div {{
  margin-top: 0.2rem;
  margin-bottom: 0.7rem;
}}

/* Sidebar layout */
.sb-title {{
  font-size: 17px;
  font-weight: 800;
  margin: 0.15rem 0 0.65rem 0;
}}
.sb-hr {{
  border: 0;
  border-top: 1px solid rgba(0, 0, 0, 0.08);
  margin: 0.8rem 0;
}}
.sb-kv {{
  border: 1px solid rgba(0, 0, 0, 0.08);
  border-radius: 12px;
  padding: 0.9rem 0.95rem;
  background: rgba(255, 255, 255, 0.55);
}}
.sb-kv .row {{
  display: block;
  margin: 0.55rem 0;
}}
.sb-kv .k {{
  color: rgba(0, 0, 0, 0.55);
  font-size: 14px;
  font-weight: 700;
  white-space: nowrap;
  margin-bottom: 0.25rem;
}}
.sb-kv .v {{
  font-size: 14px;
  line-height: 1.45;
  color: rgba(0, 0, 0, 0.85);
  overflow-wrap: anywhere;
}}
.sb-kv code {{
  white-space: normal;
}}
.sb-dot {{
  display: inline-block;
  width: 8px;
  height: 8px;
  border-radius: 999px;
  margin-right: 6px;
  background: #9aa0a6;
}}
.sb-dot.ok {{ background: #1e8e3e; }}
.sb-dot.warn {{ background: #b06000; }}

div[data-testid="stDownloadButton"] button {{
  width: 100%;
  border-radius: 10px !important;
  border: 1px solid rgba(0, 0, 0, 0.12) !important;
  background: rgba(255, 255, 255, 0.75) !important;
  font-weight: 800 !important;
  font-size: 14px !important;
  padding: 0.65rem 0.75rem !important;
}}
div[data-testid="stDownloadButton"] button:hover {{
  border-color: rgba(0, 0, 0, 0.20) !important;
}}
 /* Make title breathe */
 h1 {{
   margin-bottom: 0.8rem !important;
 }}

 /* DataFrame: match dashboard background (Streamlit DataTable / BaseWeb) */
 div[data-testid="stDataFrame"] {{
   background-color: {bg} !important;
 }}
 .stDataFrame {{
   background-color: {bg} !important;
 }}
 div[data-testid="stDataFrame"] div[data-baseweb="data-table"] {{
   background-color: {bg} !important;
 }}
 .stDataFrame div[data-baseweb="data-table"] {{
   background-color: {bg} !important;
 }}
 div[data-testid="stDataFrame"] div[data-baseweb="data-table"] div[role="gridcell"] {{
   background-color: {bg} !important;
 }}
 .stDataFrame div[data-baseweb="data-table"] div[role="gridcell"] {{
   background-color: {bg} !important;
 }}
 div[data-testid="stDataFrame"] div[data-baseweb="data-table"] div[role="row"] {{
   background-color: {bg} !important;
 }}
 .stDataFrame div[data-baseweb="data-table"] div[role="row"] {{
   background-color: {bg} !important;
 }}
 div[data-testid="stDataFrame"] div[data-baseweb="data-table"] div[role="columnheader"] {{
   background-color: {sbg} !important;
 }}
 .stDataFrame div[data-baseweb="data-table"] div[role="columnheader"] {{
   background-color: {sbg} !important;
 }}
 div[data-testid="stDataFrame"] div[data-baseweb="data-table"] div[role="row"]:hover div[role="gridcell"] {{
   background-color: {sbg} !important;
 }}
 .stDataFrame div[data-baseweb="data-table"] div[role="row"]:hover div[role="gridcell"] {{
   background-color: {sbg} !important;
 }}

 /* DataFrame: fallback for table-based renderers */
 .stDataFrame table {{
   background-color: {bg} !important;
 }}
 .stDataFrame thead tr th {{
   background-color: {sbg} !important;
 }}
 .stDataFrame tbody tr td {{
   background-color: {bg} !important;
 }}
 .stDataFrame tbody tr:hover td {{
   background-color: {sbg} !important;
 }}
 </style>
         """,
         unsafe_allow_html=True,
     )


def _split_family(family: str) -> tuple[str, str]:
    # Expected format: "<제품명> + <POWER>"
    if family is None:
        return "", ""
    s = str(family)
    if " + " not in s:
        return s, ""
    left, right = s.rsplit(" + ", 1)
    return left.strip(), right.strip()


def _format_int(x) -> str:
    try:
        v = float(x)
    except Exception:
        return "0"
    if pd.isna(v):
        return "0"
    return f"{int(round(v)):,}"


def _normalize_signed_2dp(v, *, zero_sign: str = "+") -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return f"{zero_sign}00.00"
    s = str(v).strip()
    if not s or s.lower() == "nan" or s == "<NA>":
        return f"{zero_sign}00.00"
    try:
        x = float(s)
    except Exception:
        return s
    if pd.isna(x):
        return f"{zero_sign}00.00"
    sign = zero_sign if x == 0 else ("-" if x < 0 else "+")
    mag = abs(x)
    return f"{sign}{mag:.2f}"


def _parse_search_terms(raw: str) -> list[str]:
    if raw is None:
        return []
    terms = [t.strip() for t in str(raw).split(",")]
    return [t for t in terms if t]


def _filter_by_name_contains(df: pd.DataFrame, name_col: str, raw_terms: str) -> pd.DataFrame:
    terms = _parse_search_terms(raw_terms)
    if not terms or name_col not in df.columns:
        return df
    # OR-match: include row if any term is contained (case-insensitive).
    pattern = "|".join(re.escape(t) for t in terms)
    mask = df[name_col].astype("string").fillna("").str.contains(pattern, case=False, regex=True, na=False)
    return df.loc[mask].copy()


def _filter_by_any_contains(df: pd.DataFrame, cols: list[str], raw_terms: str) -> pd.DataFrame:
    terms = _parse_search_terms(raw_terms)
    cols = [c for c in cols if c in df.columns]
    if not terms or not cols:
        return df
    pattern = "|".join(re.escape(t) for t in terms)
    mask = pd.Series(False, index=df.index)
    for c in cols:
        mask = mask | df[c].astype("string").fillna("").str.contains(pattern, case=False, regex=True, na=False)
    return df.loc[mask].copy()


DEFAULT_STAGE_COLS = ["사출", "분리", "하이드레이션", "접착", "누수규격"]


@st.cache_data(show_spinner=False)
def _to_excel_bytes(df: pd.DataFrame, *, sheet_name: str = "data") -> bytes:
    output = BytesIO()
    xdf = df.copy()

    # Ensure date columns export as YYYY-MM-DD (no time component).
    for c in xdf.columns:
        if c == "납기일" or pd.api.types.is_datetime64_any_dtype(xdf[c]):
            dt = pd.to_datetime(xdf[c], errors="coerce")
            xdf[c] = dt.dt.strftime("%Y-%m-%d")
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            xdf.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    except Exception:
        # Fallback to xlsxwriter if available, otherwise raise.
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:  # type: ignore[call-arg]
            xdf.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def _format_item_code_list(codes: list[str], *, max_show: int = 12) -> str:
    codes = [str(c).strip() for c in codes if str(c).strip()]
    if not codes:
        return ""
    uniq = []
    seen = set()
    for c in codes:
        if c in seen:
            continue
        uniq.append(c)
        seen.add(c)
    if len(uniq) <= max_show:
        return ", ".join(uniq)
    head = ", ".join(uniq[:max_show])
    return f"{head} …(+{len(uniq) - max_show})"


def _attach_item_codes(
    df: pd.DataFrame,
    detail: pd.DataFrame | None,
    *,
    out_col: str = "제품코드",
    allowed_prefixes: list[str] | None = None,
) -> pd.DataFrame:
    if detail is None or detail.empty:
        return df
    if "제품 코드" not in detail.columns:
        return df

    # Join keys:
    # - Avoid ADD/CP/AXIS here because UI normalization turns missing values into ""
    #   while the detail CSV often has <NA>, which breaks equality joins.
    key_candidates = ["신규분류 요약코드", "제품군", "납기일"]
    key_cols = [c for c in key_candidates if c in df.columns and c in detail.columns]
    if not key_cols:
        return df

    d = detail.copy()
    if "납기일" in d.columns:
        d["납기일"] = pd.to_datetime(d["납기일"], errors="coerce")
    if "납기일" in df.columns:
        left = df.copy()
        left["납기일"] = pd.to_datetime(left["납기일"], errors="coerce")
    else:
        left = df.copy()

    # Normalize string keys to avoid mismatch due to NA vs "" vs whitespace.
    for c in key_cols:
        if c == "납기일":
            continue
        d[c] = d[c].astype("string").fillna("").str.strip()
        left[c] = left[c].astype("string").fillna("").str.strip()

    prefixes: list[str] | None = None
    if allowed_prefixes:
        prefixes = [str(p).strip().upper() for p in allowed_prefixes if str(p).strip()]

    def _codes_for_group(s: pd.Series) -> str:
        items = s.dropna().astype(str).map(lambda x: x.strip())
        if prefixes:
            items = items[items.map(lambda x: any(x.upper().startswith(p) for p in prefixes))]
        return _format_item_code_list(sorted(set(items.tolist())))

    codes = (
        d.groupby(key_cols, dropna=False)["제품 코드"]
        .apply(_codes_for_group)
        .reset_index()
        .rename(columns={"제품 코드": out_col})
    )
    merged = left.merge(codes, on=key_cols, how="left")
    return merged


@st.cache_data(show_spinner=False)
def _load_due_csv(path: str, mtime: float) -> pd.DataFrame:
    _ = mtime  # cache-buster when file changes
    header = pd.read_csv(path, nrows=0)
    dtype: dict[str, str] = {}
    for c in ["신규분류 요약코드", "제품군", "ADD", "CP", "AXIS"]:
        if c in header.columns:
            dtype[c] = "string"
    df = pd.read_csv(path, dtype=dtype if dtype else None)
    df["납기일"] = pd.to_datetime(df["납기일"], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def _load_order_detail_csv(path: str, mtime: float) -> pd.DataFrame:
    _ = mtime  # cache-buster when file changes
    header = pd.read_csv(path, nrows=0)
    dtype: dict[str, str] = {}
    for c in ["이니셜", "수주번호", "신규분류 요약코드", "수요 제품 이름", "제품군", "제품 코드", "ADD", "CP", "AXIS"]:
        if c in header.columns:
            dtype[c] = "string"
    df = pd.read_csv(path, dtype=dtype if dtype else None)
    if "납기일" in df.columns:
        df["납기일"] = pd.to_datetime(df["납기일"], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def _load_due_prepared(path: str, mtime: float) -> pd.DataFrame:
    # One-shot cache: load + prepare (avoids recomputing _prepare_lens_df on every rerun).
    base = _load_due_csv(path, mtime)
    return _prepare_lens_df(base)


@st.cache_data(show_spinner=False)
def _load_order_detail_prepared(path: str, mtime: float) -> pd.DataFrame:
    base = _load_order_detail_csv(path, mtime)
    out = _prepare_lens_df(base)
    # In order-detail, prefer the full product name if present.
    if "수요 제품 이름" in out.columns:
        out["품명"] = out["수요 제품 이름"].astype("string").fillna("")
    return out


@st.cache_data(show_spinner=False)
def _load_order_detail_grouped(path: str, mtime: float) -> pd.DataFrame:
    """
    Order view heavy step: group product-level rows once per source file.
    Filters (due-date/search/code) can be applied afterwards without regrouping.
    """
    df = _load_order_detail_prepared(path, mtime)
    numeric_cols = [c for c in DEFAULT_STAGE_COLS if c in df.columns]
    if numeric_cols:
        work = df.copy()
        for c in numeric_cols:
            work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0)
    else:
        work = df

    group_cols = [c for c in ["이니셜", "수주번호", "신규분류 요약코드", "품명", "납기일"] if c in work.columns]
    if group_cols and numeric_cols:
        work = work.groupby(group_cols, dropna=False, as_index=False)[numeric_cols].sum(numeric_only=True)
    return work


def _prepare_lens_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "제품군" in out.columns:
        out["품명"], out["POWER"] = zip(*out["제품군"].map(_split_family))
    else:
        out["품명"] = out.get("수요 제품 이름", out.get("품명", "")).astype("string").fillna("")
        out["POWER"] = ""

    out["POWER"] = out["POWER"].astype("string")
    out["POWER_num"] = pd.to_numeric(out["POWER"], errors="coerce")
    power_num = out["POWER_num"]

    # POWER sort:
    # - Positive first: +10.00 -> +0.00
    # - Then zero: -0.00 (business display)
    # - Then negative: -0.00 -> -12.00 (magnitude increasing)
    out["POWER_group"] = 3
    out.loc[power_num.notna() & (power_num > 0), "POWER_group"] = 0
    out.loc[power_num.notna() & (power_num == 0), "POWER_group"] = 1
    out.loc[power_num.notna() & (power_num < 0), "POWER_group"] = 2
    out["POWER_sort"] = float("inf")
    out.loc[power_num.notna(), "POWER_sort"] = -power_num

    for c in ["ADD", "CP", "AXIS"]:
        if c in out.columns:
            out[c] = out[c].astype("string")

    # Lens type rules for display:
    # - Spherical: POWER only
    # - Toric: POWER + CP + AXIS
    # - Multifocal: POWER + ADD
    cp_raw = out["CP"] if "CP" in out.columns else pd.Series([""] * len(out), dtype="string")
    add_raw = out["ADD"] if "ADD" in out.columns else pd.Series([""] * len(out), dtype="string")
    axis_raw = out["AXIS"] if "AXIS" in out.columns else pd.Series([""] * len(out), dtype="string")

    cp_has = cp_raw.fillna("").astype(str).str.strip() != ""
    add_has = add_raw.fillna("").astype(str).str.strip() != ""
    axis_has = axis_raw.fillna("").astype(str).str.strip() != ""

    is_toric = cp_has | axis_has
    is_multi = (~is_toric) & add_has

    def _norm_axis(x: str) -> str:
        s = str(x).strip()
        if not s or s.lower() == "nan" or s == "<NA>":
            return "000"
        try:
            n = int(float(s))
            n = max(0, min(999, n))
            return f"{n:03d}"
        except Exception:
            return str(s).zfill(3)[:3]

    if "CP" in out.columns:
        out["CP"] = ""
        out.loc[is_toric, "CP"] = cp_raw.loc[is_toric].map(lambda x: _normalize_signed_2dp(x, zero_sign="+"))
        out.loc[is_toric & (out["CP"].astype(str).str.strip() == ""), "CP"] = "+0.00"
        out["CP"] = out["CP"].astype("string")
        out["CP_num"] = pd.to_numeric(out["CP"], errors="coerce").fillna(0)
    else:
        out["CP_num"] = 0

    if "AXIS" in out.columns:
        out["AXIS"] = ""
        out.loc[is_toric, "AXIS"] = axis_raw.loc[is_toric].map(_norm_axis)
        out.loc[is_toric & (out["AXIS"].astype(str).str.strip() == ""), "AXIS"] = "000"
        out["AXIS"] = out["AXIS"].astype("string")
        out["AXIS_num"] = pd.to_numeric(out["AXIS"], errors="coerce").fillna(0)
    else:
        out["AXIS_num"] = 0

    if "ADD" in out.columns:
        out["ADD"] = ""
        out.loc[is_multi, "ADD"] = add_raw.loc[is_multi].map(lambda x: _normalize_signed_2dp(x, zero_sign="+"))
        out.loc[is_multi & (out["ADD"].astype(str).str.strip() == ""), "ADD"] = "+0.00"
        out["ADD"] = out["ADD"].astype("string")
        out["ADD_num"] = pd.to_numeric(out["ADD"], errors="coerce").fillna(0)
    else:
        out["ADD_num"] = 0

    return out


def _apply_due_date_end_filter(df: pd.DataFrame, end: date) -> pd.DataFrame:
    if "납기일" not in df.columns:
        return df
    due = pd.to_datetime(df["납기일"], errors="coerce")
    mask = due.dt.date.le(end)
    return df.loc[mask].copy()


def main() -> None:
    st.title("S관 생산 필요수량 대시보드")
    _apply_local_theme_css()

    dashboard_links = _load_dashboard_links()
    with st.sidebar:
        if dashboard_links:
            st.markdown("<div class='sb-title'>대시보드 링크</div>", unsafe_allow_html=True)
            for item in dashboard_links:
                st.link_button(item["label"], item["url"], use_container_width=True)
            st.markdown("<div class='sb-hr'></div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='sb-title'>대시보드 링크</div>", unsafe_allow_html=True)
            st.caption(f"`{DASHBOARD_LINKS_PATH}`에 링크를 추가하면 여기서 새 탭으로 열 수 있어요.")
            st.markdown("<div class='sb-hr'></div>", unsafe_allow_html=True)

    excel_path = _find_repo_excel()
    if excel_path:
        st.caption(f"업데이트(엑셀 저장시각): `{_file_mtime_label(excel_path)}`")
    with st.sidebar:
        st.markdown("<div class='sb-title'>자료 다운로드</div>", unsafe_allow_html=True)
        b = _read_bytes(TEMPLATE_XLSX_PATH)
        if b is not None:
            st.download_button(
                "업로드 양식(.xlsx) 다운로드",
                data=b,
                file_name=os.path.basename(TEMPLATE_XLSX_PATH),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_tpl_xlsx",
            )
        else:
            st.caption(f"양식 파일이 없습니다: `{TEMPLATE_XLSX_PATH}`")

        if excel_path:
            st.caption(f"읽는 파일: `{os.path.basename(excel_path)}`")
        else:
            st.caption("읽는 파일: -")

    if not excel_path:
        st.error("`s관 부족수량.xlsx` 파일을 찾지 못했습니다. 저장소에 커밋해 두고 다시 실행하세요.")
        st.caption("기대 위치: `s관 부족수량.xlsx` 또는 `data/s관 부족수량.xlsx`")
        st.stop()

    with st.spinner("엑셀에서 데이터 생성/로딩 중..."):
        out_dir = OUT_DIR
        ensure = _ensure_latest_outputs(excel_path=excel_path, out_dir=out_dir)
        if not ensure.get("ok"):
            st.error(f"데이터 생성 실패: {ensure.get('reason')}")
            st.stop()
        if ensure.get("regenerated"):
            st.cache_data.clear()

        due_csv = str(ensure["due_csv"])
        detail_csv = str(ensure["detail_csv"])

    detail_for_map: pd.DataFrame | None = None
    try:
        detail_for_map = _load_order_detail_prepared(detail_csv, os.path.getmtime(detail_csv))
    except Exception:
        detail_for_map = None

    def _tab_sort_key(code: str) -> tuple[int, int, str]:
        s = str(code).strip()
        sl = s.lower()
        is_color = 1 if "color" in sl else 0  # non-color first, color later
        if "toric" in sl:
            lens_rank = 2
        elif "m/f" in sl or "multifocal" in sl or "multi" in sl:
            lens_rank = 1
        elif "sph" in sl or "spherical" in sl:
            lens_rank = 0
        else:
            lens_rank = 9
        return (is_color, lens_rank, sl)

    try:
        base = _load_due_prepared(due_csv, os.path.getmtime(due_csv))
    except Exception as e:
        st.error("데이터 로딩 실패")
        st.caption(str(e))
        st.stop()

    df = base
    new_code_col = "신규분류 요약코드" if "신규분류 요약코드" in df.columns else None

    def render(
        filtered: pd.DataFrame,
        *,
        ui_key_prefix: str,
        process_only: str | None = None,
        selected_code: str | None = None,
    ) -> None:
        stage_cols_raw = [process_only] if process_only else DEFAULT_STAGE_COLS
        numeric_cols = [c for c in stage_cols_raw if c in filtered.columns]
        total_label = f"{process_only} 필요수량" if process_only else "총 필요수량"
        header_ph = st.empty()
        metric_ph = st.empty()

        search_raw = st.text_input(
            "검색 (품명)",
            placeholder="예: O2O2, SEPIA, ASH",
            key=f"{ui_key_prefix}_name_search",
        )
        filtered = _filter_by_name_contains(filtered, "품명", search_raw)

        if process_only and process_only in filtered.columns:
            proc_v = pd.to_numeric(filtered[process_only], errors="coerce").fillna(0)
            filtered = filtered.loc[proc_v.ne(0)].copy()

        df_num = filtered.copy()
        for c in numeric_cols:
            df_num[c] = pd.to_numeric(df_num[c], errors="coerce").fillna(0)

        total_col = process_only if process_only in df_num.columns else ("누수규격" if "누수규격" in df_num.columns else None)
        header_ph.subheader(total_label)
        metric_ph.metric(label="", value=_format_int(df_num[total_col].sum()) if total_col else "0")

        st.subheader("납기별 상세")

        view = filtered.copy()
        sort_cols: list[str] = []
        for c in [
            "납기일",
            "신규분류 요약코드",
            "품명",
            "POWER_group",
            "POWER_sort",
            "CP_num",
            "AXIS_num",
            "ADD_num",
        ]:
            if c in view.columns:
                sort_cols.append(c)
        if sort_cols:
            view = view.sort_values(sort_cols, ascending=[True] * len(sort_cols), na_position="last")

        allowed_prefixes: list[str] | None = None
        if process_only:
            prefix_map: dict[str, list[str]] = {
                "사출": ["R"],
                "분리": ["Q"],
                "하이드레이션": ["P"],
                "접착": ["P"],
                "누수규격": ["P"],
            }
            allowed_prefixes = prefix_map.get(process_only, None)
            view = _attach_item_codes(view, detail_for_map, allowed_prefixes=allowed_prefixes)
            if "제품코드" not in view.columns:
                view["제품코드"] = ""

        # Export dataframe (keep numeric) BEFORE display formatting.
        export_df = view.copy()
        for s in stage_cols_raw:
            if s in export_df.columns:
                export_df[s] = pd.to_numeric(export_df[s], errors="coerce").fillna(0).astype(int)

        has_toric = False
        has_multi = False
        if "CP" in view.columns or "AXIS" in view.columns:
            cp_has = (
                view["CP"].astype("string").fillna("").astype(str).str.strip().ne("").any()
                if "CP" in view.columns
                else False
            )
            axis_has = (
                view["AXIS"].astype("string").fillna("").astype(str).str.strip().ne("").any()
                if "AXIS" in view.columns
                else False
            )
            has_toric = bool(cp_has or axis_has)
        if "ADD" in view.columns:
            has_multi = bool(view["ADD"].astype("string").fillna("").astype(str).str.strip().ne("").any())

        # 공정별 보기에서 분류가 명확한 경우 규격 컬럼을 강제 결정
        if selected_code and selected_code != "전체":
            sl = str(selected_code).lower()
            if "toric" in sl:
                has_toric, has_multi = True, False
            elif "m/f" in sl or "multifocal" in sl or "multi" in sl:
                has_toric, has_multi = False, True
            elif "sph" in sl or "spherical" in sl:
                has_toric, has_multi = False, False

        cols = ["신규분류 요약코드", "품명"]
        if process_only:
            cols.append("제품코드")
        cols += ["POWER", "납기일"] + stage_cols_raw
        if has_toric:
            power_idx = cols.index("POWER")
            cols[power_idx + 1 : power_idx + 1] = ["CP", "AXIS"]
        if has_multi:
            power_idx = cols.index("POWER")
            insert_at = power_idx + 1 + (2 if has_toric else 0)
            cols[insert_at:insert_at] = ["ADD"]

        cols = [c for c in cols if c in view.columns]

        export_cols = cols.copy()
        export_df2 = export_df.copy()

        # Download rules:
        # - 납기별 상세: 화면 그대로(제품코드 없음)
        # - 공정별 보기: 제품코드 포함 + 공정별 prefix 필터
        if process_only:
            if "제품코드" in export_df2.columns and "제품코드" not in export_cols:
                # 품명과 POWER 사이에 제품코드
                if "품명" in export_cols and "POWER" in export_cols:
                    export_cols.insert(export_cols.index("POWER"), "제품코드")
                elif "품명" in export_cols:
                    export_cols.insert(export_cols.index("품명") + 1, "제품코드")
                else:
                    export_cols.insert(0, "제품코드")

        # Display formatting (comma) after export snapshot.
        for s in stage_cols_raw:
            if s in view.columns:
                view[s] = pd.to_numeric(view[s], errors="coerce").fillna(0).astype(int)

        column_config = {
            "신규분류 요약코드": st.column_config.TextColumn(width="medium"),
            "품명": st.column_config.TextColumn(width="large"),
            "제품코드": st.column_config.TextColumn(width="medium"),
            "POWER": st.column_config.TextColumn(width="small"),
            "CP": st.column_config.TextColumn(width="small"),
            "AXIS": st.column_config.TextColumn(width="small"),
            "ADD": st.column_config.TextColumn(width="small"),
            "납기일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
        }
        for c in stage_cols_raw:
            if c in cols:
                column_config[c] = st.column_config.NumberColumn(format="localized", width="small")
        column_config = {k: v for k, v in column_config.items() if k in cols}

        # Download button should not push the totals row away from the table header,
        # so render it BEFORE totals grid.
        xlsx_bytes = _to_excel_bytes(export_df2[export_cols], sheet_name="다운로드")
        st.download_button(
            "엑셀 다운로드",
            data=xlsx_bytes,
            file_name=f"{'공정' if process_only else '납기'}_{selected_code or '전체'}_{process_only or '전체'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{ui_key_prefix}_download",
        )

        stage_totals = {c: _format_int(df_num[c].sum()) for c in stage_cols_raw if c in df_num.columns}
        view_show = view[cols].copy()
        for c in stage_cols_raw:
            if c in view_show.columns:
                view_show[c] = pd.to_numeric(view_show[c], errors="coerce").fillna(0).map(_format_int)
        # Display totals as a clean HTML row above the columns
        if stage_totals:
            totals_html = " ".join([
                f"<span style='margin-right: 20px; font-size: 15px;'>{c}: <strong style='color: #0066cc;'>{stage_totals[c]}</strong></span>"
                for c in stage_cols_raw if c in stage_totals
            ])
            st.markdown(f"<div style='margin-bottom: 8px; padding: 4px 8px;'>{totals_html}</div>", unsafe_allow_html=True)

        view_show.columns = cols

        table_h = _table_height_for_rows(len(view), min_height=280, max_height=720)
        st.dataframe(
            _style_dataframe_like_dashboard(view_show),
            use_container_width=True,
            height=table_h,
            hide_index=True,
            column_config=column_config,
        )

    if new_code_col is None:
        render(df, ui_key_prefix="all")
        return

    view_options = ["납기별 상세", "공정별 보기", "수주별 현황"]
    _pre_widget_single_select_fix(key="view_mode", default="납기별 상세", options=view_options)
    view_mode_raw = st.segmented_control(
        "보기",
        options=view_options,
        default="납기별 상세",
        key="view_mode",
        on_change=_on_change_single_select,
        args=("view_mode", "납기별 상세", view_options),
        label_visibility="collapsed",
    )
    view_mode = _coerce_single_value(view_mode_raw, default="납기별 상세", options=view_options)

    prev_mode = st.session_state.get("_prev_view_mode")
    if prev_mode != view_mode:
        st.session_state["code_pill"] = "전체"
        if view_mode == "납기별 상세":
            # Reset due-date filter when entering due view.
            st.session_state["due_due_quick"] = "해제"
            st.session_state["due_due_end"] = _today_kst()
            st.session_state["_prev_due_due_quick"] = "해제"
        if view_mode == "공정별 보기":
            st.session_state["process_pill"] = "사출"
            # Reset due-date filter when entering process view.
            st.session_state["proc_due_quick"] = "해제"
            st.session_state["proc_due_end"] = _today_kst()
            st.session_state["_prev_proc_due_quick"] = "해제"
        if view_mode == "수주별 현황":
            # Always reset due-date filter when entering order view.
            st.session_state["order_due_quick"] = "해제"
            st.session_state["order_due_end"] = _today_kst()
            st.session_state["_prev_order_due_quick"] = "해제"
        st.session_state["_prev_view_mode"] = view_mode

    process_only = None
    if view_mode == "납기별 상세":
        # Due date end quick-picks for due view.
        due_quick_options = ["해제", "직접", "+7일", "+14일"]
        _pre_widget_single_select_fix(key="due_due_quick", default="해제", options=due_quick_options)
        due_quick_raw = st.pills(
            "납기일 종료 (빠른 선택)",
            options=due_quick_options,
            default="해제",
            key="due_due_quick",
            selection_mode="single",
            on_change=_on_change_single_select,
            args=("due_due_quick", "해제", due_quick_options),
            label_visibility="collapsed",
        )
        due_quick = _coerce_single_value(due_quick_raw, default="해제", options=due_quick_options)
        if due_quick == "+7일":
            due_default_end = _today_kst() + timedelta(days=7)
        elif due_quick == "+14일":
            due_default_end = _today_kst() + timedelta(days=14)
        else:
            due_default_end = _today_kst()

        prev_due_quick = st.session_state.get("_prev_due_due_quick")
        if prev_due_quick != due_quick:
            st.session_state["due_due_end"] = due_default_end
            st.session_state["_prev_due_due_quick"] = due_quick

        due_end_date = st.date_input(
            "납기일 종료",
            value=st.session_state.get("due_due_end", due_default_end),
            key="due_due_end",
            disabled=(due_quick == "해제"),
        )

    if view_mode == "공정별 보기":
        _pre_widget_single_select_fix(key="process_pill", default="사출", options=DEFAULT_STAGE_COLS)
        process_only_raw = st.pills(
            "공정",
            options=DEFAULT_STAGE_COLS,
            default="사출",
            key="process_pill",
            on_change=_on_change_single_select,
            args=("process_pill", "사출", DEFAULT_STAGE_COLS),
            label_visibility="collapsed",
        )
        process_only = _coerce_single_value(process_only_raw, default="사출", options=DEFAULT_STAGE_COLS)

        # Due date end quick-picks (same idea as order view).
        proc_quick_options = ["해제", "직접", "+7일", "+14일"]
        _pre_widget_single_select_fix(key="proc_due_quick", default="해제", options=proc_quick_options)
        proc_quick_raw = st.pills(
            "납기일 종료 (빠른 선택)",
            options=proc_quick_options,
            default="해제",
            key="proc_due_quick",
            selection_mode="single",
            on_change=_on_change_single_select,
            args=("proc_due_quick", "해제", proc_quick_options),
            label_visibility="collapsed",
        )
        proc_quick = _coerce_single_value(proc_quick_raw, default="해제", options=proc_quick_options)
        if proc_quick == "+7일":
            proc_default_end = _today_kst() + timedelta(days=7)
        elif proc_quick == "+14일":
            proc_default_end = _today_kst() + timedelta(days=14)
        else:
            proc_default_end = _today_kst()

        prev_proc_quick = st.session_state.get("_prev_proc_due_quick")
        if prev_proc_quick != proc_quick:
            st.session_state["proc_due_end"] = proc_default_end
            st.session_state["_prev_proc_due_quick"] = proc_quick

        proc_end_date = st.date_input(
            "납기일 종료",
            value=st.session_state.get("proc_due_end", proc_default_end),
            key="proc_due_end",
            disabled=(proc_quick == "해제"),
        )

    # 분류 pills (view-mode별로 totals 계산 데이터가 다름)
    if view_mode == "수주별 현황":
        if detail_for_map is None:
            st.error("수주별 현황 데이터가 없습니다. 엑셀 시트/컬럼을 확인하세요.")
            st.stop()
        order_df = _load_order_detail_grouped(detail_csv, os.path.getmtime(detail_csv))

        # Due date end quick-picks
        order_quick_options = ["해제", "직접", "+7일", "+14일"]
        _pre_widget_single_select_fix(key="order_due_quick", default="해제", options=order_quick_options)
        quick_raw = st.pills(
            "납기일 종료 (빠른 선택)",
            options=order_quick_options,
            default="해제",
            key="order_due_quick",
            selection_mode="single",
            on_change=_on_change_single_select,
            args=("order_due_quick", "해제", order_quick_options),
            label_visibility="collapsed",
        )
        quick = _coerce_single_value(quick_raw, default="해제", options=order_quick_options)
        if quick == "+7일":
            default_end = _today_kst() + timedelta(days=7)
        elif quick == "+14일":
            default_end = _today_kst() + timedelta(days=14)
        else:
            default_end = _today_kst()

        # Ensure quick pick actually updates the date_input (Streamlit keeps widget state by key).
        prev_quick = st.session_state.get("_prev_order_due_quick")
        if prev_quick != quick:
            st.session_state["order_due_end"] = default_end
            st.session_state["_prev_order_due_quick"] = quick

        end_date = st.date_input(
            "납기일 종료",
            value=st.session_state.get("order_due_end", default_end),
            key="order_due_end",
            disabled=(quick == "해제"),
        )
        if quick != "해제":
            order_df = _apply_due_date_end_filter(order_df, end_date)
        codes_src = order_df
        value_col = "누수규격"
    else:
        codes_src = df
        if view_mode == "납기별 상세":
            due_quick_state = st.session_state.get("due_due_quick", "해제")
            if due_quick_state != "해제":
                codes_src = _apply_due_date_end_filter(
                    codes_src,
                    st.session_state.get("due_due_end", _today_kst()),
                )
        if view_mode == "공정별 보기":
            proc_quick_state = st.session_state.get("proc_due_quick", "해제")
            if proc_quick_state != "해제":
                codes_src = _apply_due_date_end_filter(codes_src, st.session_state.get("proc_due_end", _today_kst()))
        value_col = process_only if process_only else "누수규격"

    totals_base: dict[str, float] = {}
    total_all = 0.0
    if value_col in codes_src.columns:
        tmp = codes_src.copy()
        tmp[value_col] = pd.to_numeric(tmp[value_col], errors="coerce").fillna(0)
        totals_base = tmp.groupby(new_code_col, dropna=False)[value_col].sum(numeric_only=True).to_dict()
        total_all = float(tmp[value_col].sum())

    codes = (
        codes_src[new_code_col]
        .astype("string")
        .fillna("")
        .map(lambda x: x.strip() if isinstance(x, str) else "")
    )
    code_options = sorted([c for c in codes.unique().tolist() if c], key=_tab_sort_key)

    def _code_label(code: str) -> str:
        if code == "전체":
            return f"전체 ({_format_int(total_all)})"
        return f"{code} ({_format_int(totals_base.get(code, 0.0))})"

    code_all_options = ["전체"] + code_options
    _pre_widget_single_select_fix(key="code_pill", default="전체", options=code_all_options)
    code_raw = st.pills(
        "분류",
        options=code_all_options,
        default="전체",
        key="code_pill",
        format_func=_code_label,
        selection_mode="single",
        on_change=_on_change_single_select,
        args=("code_pill", "전체", code_all_options),
        label_visibility="collapsed",
    )
    code = _coerce_single_value(code_raw, default="전체", options=code_all_options)

    if view_mode == "수주별 현황":
        subset = order_df if code == "전체" else order_df[order_df[new_code_col].astype("string") == code].copy()
        stage_cols_raw = DEFAULT_STAGE_COLS
        numeric_cols = [c for c in stage_cols_raw if c in subset.columns]

        search_raw = st.text_input(
            "검색 (품명/이니셜/수주번호)",
            placeholder="예: 해외, 202601, SEPIA",
            key=f"order_{code}_search",
        )
        subset = _filter_by_any_contains(subset, ["품명", "이니셜", "수주번호"], search_raw)

        detail_num = subset.copy()
        for c in numeric_cols:
            if c in detail_num.columns:
                detail_num[c] = pd.to_numeric(detail_num[c], errors="coerce").fillna(0)

        stage_sum = 0
        for c in numeric_cols:
            stage_sum = stage_sum + detail_num[c].fillna(0)
        detail_num = detail_num.loc[stage_sum.fillna(0).gt(0)].copy()

        stage_totals = {
            c: _format_int(pd.to_numeric(detail_num[c], errors="coerce").fillna(0).sum()) for c in numeric_cols
        }

        # Summary rows: order-level (initial + order number), summed across products.
        summary_base = detail_num.copy()
        if "납기일" in summary_base.columns:
            due_dt = pd.to_datetime(summary_base["납기일"], errors="coerce")
            summary_base["_due_date"] = due_dt.dt.date
        else:
            summary_base["_due_date"] = pd.NaT

        group_key = [c for c in ["이니셜", "수주번호"] if c in summary_base.columns]
        if not group_key:
            group_key = ["수주번호"] if "수주번호" in summary_base.columns else []
        if code == "전체" and new_code_col in summary_base.columns:
            group_key = [c for c in [*group_key, new_code_col] if c in summary_base.columns]

        agg_spec: dict[str, str] = {c: "sum" for c in numeric_cols}
        if "품명" in summary_base.columns:
            agg_spec["품목수"] = "nunique"
        else:
            agg_spec["품목수"] = "size"
        if "_due_date" in summary_base.columns:
            agg_spec["납기일"] = "min"

        # Pandas needs existing columns for named aggs; create working cols.
        work = summary_base.copy()
        if "품명" in work.columns:
            work["품목수"] = work["품명"]
        else:
            work["품목수"] = 1
        if "_due_date" in work.columns:
            work["납기일"] = work["_due_date"]

        order_num = work.groupby(group_key, dropna=False, as_index=False).agg(agg_spec)
        sort_cols = [c for c in ["납기일", new_code_col, "이니셜", "수주번호"] if c in order_num.columns]
        if sort_cols:
            order_num = order_num.sort_values(sort_cols, ascending=[True] * len(sort_cols), na_position="last")
        order_num.insert(0, "우선순위", range(1, len(order_num) + 1))

        order_view = order_num.copy()
        for c in numeric_cols:
            if c in order_view.columns:
                order_view[c] = pd.to_numeric(order_view[c], errors="coerce").fillna(0).astype(int)
        if "품목수" in order_view.columns:
            order_view["품목수"] = pd.to_numeric(order_view["품목수"], errors="coerce").fillna(0).astype(int)

        base_cols = ["우선순위", "이니셜", "수주번호"]
        if code == "전체":
            base_cols.append(new_code_col)
        base_cols += ["품목수", "납기일"]
        summary_cols = [c for c in base_cols if c in order_view.columns] + numeric_cols

        col_cfg_summary: dict[str, object] = {
            "우선순위": st.column_config.NumberColumn(format="%d", width="small"),
            "이니셜": st.column_config.TextColumn(width="small"),
            "수주번호": st.column_config.TextColumn(width="medium"),
            "신규분류 요약코드": st.column_config.TextColumn(width="medium"),
            "품목수": st.column_config.NumberColumn(format="%d", width="small"),
            "납기일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
        }
        for c in numeric_cols:
            col_cfg_summary[c] = st.column_config.NumberColumn(format="localized", width="small")
        col_cfg_summary = {k: v for k, v in col_cfg_summary.items() if k in summary_cols}

        xlsx_bytes_sum = _to_excel_bytes(order_view[summary_cols], sheet_name="수주요약")
        st.download_button(
            "엑셀 다운로드 (요약)",
            data=xlsx_bytes_sum,
            file_name=f"수주요약_{code}_{_today_kst().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"order_{code}_download_sum",
        )

        order_show = order_view[summary_cols].copy()
        order_show.columns = summary_cols

        # Display totals as HTML above the table
        if stage_totals:
            totals_html = " ".join([
                f"<span style='margin-right: 20px; font-size: 15px;'>{c}: <strong style='color: #0066cc;'>{stage_totals[c]}</strong></span>"
                for c in numeric_cols if c in stage_totals
            ])
            st.markdown(f"<div style='margin-bottom: 8px; padding: 4px 8px;'>{totals_html}</div>", unsafe_allow_html=True)

        sum_h = _table_height_for_rows(len(order_view), min_height=260, max_height=520)
        st.dataframe(
            _style_dataframe_like_dashboard(order_show),
            use_container_width=True,
            height=sum_h,
            hide_index=True,
            column_config=col_cfg_summary,
        )

        st.divider()

        view = detail_num.copy()
        sort_cols = [c for c in ["납기일", "이니셜", "수주번호", "품명"] if c in view.columns]
        if sort_cols:
            view = view.sort_values(sort_cols, ascending=[True] * len(sort_cols), na_position="last")
        view.insert(0, "우선순위", range(1, len(view) + 1))

        for c in numeric_cols:
            if c in view.columns:
                view[c] = pd.to_numeric(view[c], errors="coerce").fillna(0).astype(int)

        cols = [c for c in ["우선순위", "이니셜", "수주번호", "신규분류 요약코드", "품명", "납기일"] if c in view.columns] + numeric_cols

        column_config = {
            "우선순위": st.column_config.NumberColumn(format="%d", width="small"),
            "이니셜": st.column_config.TextColumn(width="small"),
            "수주번호": st.column_config.TextColumn(width="medium"),
            "신규분류 요약코드": st.column_config.TextColumn(width="medium"),
            "품명": st.column_config.TextColumn(width="large"),
            "납기일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
        }
        for c in numeric_cols:
            column_config[c] = st.column_config.NumberColumn(format="localized", width="small")
        column_config = {k: v for k, v in column_config.items() if k in cols}

        xlsx_bytes_det = _to_excel_bytes(view[cols], sheet_name="수주상세")
        st.download_button(
            "엑셀 다운로드 (상세)",
            data=xlsx_bytes_det,
            file_name=f"수주상세_{code}_{_today_kst().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"order_{code}_download_det",
        )

        detail_h = _table_height_for_rows(len(view), min_height=320, max_height=720)
        st.dataframe(
            _style_dataframe_like_dashboard(view[cols]),
            use_container_width=True,
            height=detail_h,
            hide_index=True,
            column_config=column_config,
        )
        return

    base_df = df
    if view_mode == "납기별 상세":
        due_quick_state = st.session_state.get("due_due_quick", "해제")
        if due_quick_state != "해제":
            base_df = _apply_due_date_end_filter(base_df, st.session_state.get("due_due_end", _today_kst()))
    if view_mode == "공정별 보기":
        proc_quick_state = st.session_state.get("proc_due_quick", "해제")
        if proc_quick_state != "해제":
            base_df = _apply_due_date_end_filter(base_df, st.session_state.get("proc_due_end", _today_kst()))

    subset = base_df if code == "전체" else base_df[base_df[new_code_col].astype("string") == code].copy()
    page_key = "due" if process_only is None else f"proc_{process_only}"
    render(
        subset,
        ui_key_prefix=f"{page_key}_{code}",
        process_only=process_only,
        selected_code=code,
    )


if __name__ == "__main__":
    main()
