import os
import math
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

# Risk tab business defaults (fixed; no user settings)
RISK_CAPA_RUN_DAYS = 7
RISK_YELLOW_BUFFER_DAYS = 1.0
# Fixed WIP/queue lead time per process (calendar days)
RISK_WIP_DAYS_PER_PROCESS = 1.0
# Scheduling start offset (calendar days). Use 0 to allow same-day completion dates like 4/22.
RISK_SCHED_START_OFFSET_DAYS = 0.0


@st.cache_data(show_spinner=False)
def _load_theme_from_config_cached(_mtime: float) -> dict:
    _ = _mtime  # cache-buster when file changes
    try:
        if not os.path.exists(STREAMLIT_CONFIG_PATH):
            return {}
        with open(STREAMLIT_CONFIG_PATH, "rb") as f:
            import tomllib  # py3.11+

            data = tomllib.load(f)
        return data.get("theme", {}) if isinstance(data, dict) else {}
    except Exception:
        return {}


def _today_kst() -> date:
    return datetime.now(tz=KST).date()


def _end_of_month(d: date) -> date:
    if d.month == 12:
        first_next = date(d.year + 1, 1, 1)
    else:
        first_next = date(d.year, d.month + 1, 1)
    return first_next - timedelta(days=1)


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


def _on_change_risk_grade_pills(*, key: str, grade_options: list[str]) -> None:
    """
    Risk tab grade filter callback.
    Selecting '해제' sets selection to all grades.
    """
    v = st.session_state.get(key)
    if not isinstance(v, list):
        v = [v] if v else []
    v = [str(x).strip() for x in v if str(x).strip()]
    st.session_state[key] = [g for g in v if g in grade_options]


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


@st.cache_data(show_spinner=False)
def _xlsx_sheet_names_cached(path: str, _mtime: float) -> set[str]:
    _ = _mtime  # cache-buster when file changes
    try:
        if not path or (not os.path.exists(path)):
            return set()
        with zipfile.ZipFile(path) as zf:
            wb = zf.read("xl/workbook.xml")
        root = ET.fromstring(wb)
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets_el = root.find("m:sheets", ns)
        if sheets_el is None:
            return set()
        out: set[str] = set()
        for sh in sheets_el.findall("m:sheet", ns):
            name = (sh.attrib.get("name") or "").strip()
            if name:
                out.add(name)
        return out
    except Exception:
        # Fallback (should be rare): use pandas engine if workbook.xml is unavailable.
        try:
            xl = pd.ExcelFile(path)
            return set(str(x).strip() for x in xl.sheet_names if str(x).strip())
        except Exception:
            return set()


def _outputs_status(*, excel_path: str, out_dir: str) -> dict:
    """
    Fast up-to-date check for derived CSV outputs.

    Important UX: tab switches cause reruns; keep this check lightweight so the
    app feels instant when nothing changed.
    """
    due_csv = os.path.join(out_dir, "납기_제품군_공정별부족.csv")
    detail_csv = os.path.join(out_dir, "이니셜별_수주상세.csv")
    equip_code_target_csv = os.path.join(out_dir, "설비별_공정_제품코드_최소목표일.csv")
    prod_daily_csv = os.path.join(out_dir, "생산실적_공정별_일별양품.csv")

    try:
        excel_mtime = float(os.path.getmtime(excel_path))
    except Exception:
        return {"ok": False, "reason": "엑셀 파일을 찾을 수 없습니다."}

    # Base outputs must exist and be newer than the Excel file.
    base_ok = (
        os.path.exists(due_csv)
        and os.path.exists(detail_csv)
        and (os.path.getmtime(due_csv) >= excel_mtime)
        and (os.path.getmtime(detail_csv) >= excel_mtime)
    )
    if not base_ok:
        return {
            "ok": True,
            "needs_regen": True,
            "due_csv": due_csv,
            "detail_csv": detail_csv,
            "equip_code_target_csv": None,
            "prod_daily_csv": None,
        }

    # Optional outputs are required only if the corresponding sheets exist.
    sheet_names = _xlsx_sheet_names_cached(excel_path, excel_mtime)
    has_equip = "설비별" in sheet_names
    has_prod = "생산실적" in sheet_names

    equip_ok = os.path.exists(equip_code_target_csv) and (os.path.getmtime(equip_code_target_csv) >= excel_mtime)
    prod_ok = os.path.exists(prod_daily_csv) and (os.path.getmtime(prod_daily_csv) >= excel_mtime)

    if (has_equip and (not equip_ok)) or (has_prod and (not prod_ok)):
        return {
            "ok": True,
            "needs_regen": True,
            "due_csv": due_csv,
            "detail_csv": detail_csv,
            "equip_code_target_csv": equip_code_target_csv if has_equip else None,
            "prod_daily_csv": prod_daily_csv if has_prod else None,
        }

    return {
        "ok": True,
        "needs_regen": False,
        "regenerated": False,
        "due_csv": due_csv,
        "detail_csv": detail_csv,
        "equip_code_target_csv": equip_code_target_csv if (has_equip and equip_ok) else None,
        "prod_daily_csv": prod_daily_csv if (has_prod and prod_ok) else None,
    }


def _ensure_latest_outputs(*, excel_path: str, out_dir: str) -> dict:
    due_csv = os.path.join(out_dir, "납기_제품군_공정별부족.csv")
    detail_csv = os.path.join(out_dir, "이니셜별_수주상세.csv")
    equip_code_target_csv = os.path.join(out_dir, "설비별_공정_제품코드_최소목표일.csv")
    prod_daily_csv = os.path.join(out_dir, "생산실적_공정별_일별양품.csv")
    excel_mtime = os.path.getmtime(excel_path)

    if os.path.exists(due_csv) and os.path.exists(detail_csv):
        if os.path.getmtime(due_csv) >= excel_mtime and os.path.getmtime(detail_csv) >= excel_mtime:
            try:
                sheet_names = _xlsx_sheet_names_cached(excel_path, float(excel_mtime))
                has_equip = "설비별" in sheet_names
                has_prod = "생산실적" in sheet_names
            except Exception:
                has_equip = False
                has_prod = False

            equip_ok = os.path.exists(equip_code_target_csv) and os.path.getmtime(equip_code_target_csv) >= excel_mtime
            prod_ok = os.path.exists(prod_daily_csv) and os.path.getmtime(prod_daily_csv) >= excel_mtime

            # All required/available outputs exist and are up-to-date.
            if (not has_equip or equip_ok) and (not has_prod or prod_ok):
                return {
                    "ok": True,
                    "regenerated": False,
                    "needs_regen": False,
                    "due_csv": due_csv,
                    "detail_csv": detail_csv,
                    "equip_code_target_csv": equip_code_target_csv if has_equip and equip_ok else None,
                    "prod_daily_csv": prod_daily_csv if has_prod and prod_ok else None,
                }

    _safe_mkdir(out_dir)
    importlib.reload(excel_analysis)
    info = excel_analysis.export_due_process_shortage(file_path=excel_path, out_dir=out_dir)
    if not info.get("enabled"):
        return {"ok": False, "reason": info.get("reason") or "export failed"}
    # Optional: production actuals (for risk tab)
    try:
        prod_info = excel_analysis.export_production_daily_good_qty(file_path=excel_path, out_dir=out_dir)
        _ = prod_info
    except Exception:
        pass
    return {
        "ok": True,
        "regenerated": True,
        "needs_regen": False,
        "due_csv": due_csv,
        "detail_csv": detail_csv,
        "equip_code_target_csv": equip_code_target_csv if os.path.exists(equip_code_target_csv) else None,
        "prod_daily_csv": prod_daily_csv if os.path.exists(prod_daily_csv) else None,
    }


def _load_theme_from_config() -> dict:
    """
    Load Streamlit theme config (cached by file mtime).

    This is called frequently (table styling / CSS injection), so we cache to reduce
    redundant disk reads while keeping behavior identical when config changes.
    """

    try:
        mtime = os.path.getmtime(STREAMLIT_CONFIG_PATH) if os.path.exists(STREAMLIT_CONFIG_PATH) else -1.0
    except Exception:
        mtime = -1.0
    return _load_theme_from_config_cached(float(mtime))


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
        if c == "납기일":
            dt = pd.to_datetime(xdf[c], errors="coerce")
            xdf[c] = dt.dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_datetime64_any_dtype(xdf[c]):
            xdf[c] = xdf[c].dt.strftime("%Y-%m-%d")
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
    # Use lens-spec columns (ADD/CP/AXIS) so Toric/Multi rows map to the correct item code.
    # When a column is missing on either side, create an empty normalized column to avoid
    # accidental broad joins (which can cause mixed/duplicated codes in one cell).
    key_cols = ["신규분류 요약코드", "제품군", "ADD", "CP", "AXIS", "납기일"]
    required = [c for c in key_cols if c != "납기일"]

    d = detail.copy()
    if "납기일" in d.columns:
        d["납기일"] = pd.to_datetime(d["납기일"], errors="coerce")
    left = df.copy()
    if "납기일" in left.columns:
        left["납기일"] = pd.to_datetime(left["납기일"], errors="coerce")

    # Ensure required string keys exist on both sides.
    for c in required:
        if c not in d.columns:
            d[c] = ""
        if c not in left.columns:
            left[c] = ""

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
        # If there are still multiple codes, try to match the group AXIS (last 3 digits in code).
        try:
            group_key = s.name if isinstance(s.name, tuple) else (s.name,)
            axis_idx = key_cols.index("AXIS")
            axis_val = str(group_key[axis_idx]).strip()
        except Exception:
            axis_val = ""
        if axis_val.isdigit() and len(axis_val) == 3 and len(items) > 1:
            items_axis = items[items.map(lambda x: str(x)[-3:].isdigit() and str(x)[-3:] == axis_val)]
            if len(items_axis) > 0:
                items = items_axis

        uniq = sorted(set([x for x in items.tolist() if str(x).strip()]))
        if not uniq:
            return ""
        # 공정별 보기에서는 한 행=한 제품코드가 되도록 1개로 고정 표시
        return uniq[0]

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
def _load_equip_code_min_target_csv(path: str, mtime: float) -> pd.DataFrame:
    _ = mtime  # cache-buster when file changes
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    header = pd.read_csv(path, nrows=0)
    dtype: dict[str, str] = {}
    for c in ["공정", "제품 코드"]:
        if c in header.columns:
            dtype[c] = "string"
    df = pd.read_csv(path, dtype=dtype if dtype else None)
    if "최소목표일" in df.columns:
        df["최소목표일"] = pd.to_datetime(df["최소목표일"], errors="coerce")
    for c in ["공정", "제품 코드"]:
        if c in df.columns:
            df[c] = df[c].astype("string").fillna("").str.strip()
    return df


@st.cache_data(show_spinner=False)
def _load_prod_daily_csv(path: str | None, mtime: float) -> pd.DataFrame:
    _ = mtime  # cache-buster when file changes
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    header = pd.read_csv(path, nrows=0)
    dtype: dict[str, str] = {}
    for c in ["공정"]:
        if c in header.columns:
            dtype[c] = "string"
    df = pd.read_csv(path, dtype=dtype if dtype else None)
    if "생산일자" in df.columns:
        df["생산일자"] = pd.to_datetime(df["생산일자"], errors="coerce")
    for c in ["공정"]:
        if c in df.columns:
            df[c] = df[c].astype("string").fillna("").str.strip()
    if "양품" in df.columns:
        df["양품"] = pd.to_numeric(df["양품"], errors="coerce").fillna(0)
    return df


@st.cache_data(show_spinner=False)
def _load_injection_sheet_cached(path: str, mtime: float) -> dict[str, pd.DataFrame]:
    _ = mtime  # cache-buster when file changes
    try:
        raw = pd.read_excel(path, sheet_name="사출")
    except Exception:
        return {"equip": pd.DataFrame(), "arrange": pd.DataFrame()}

    equip_cols = ["위치", "사출 호기", "구분", "구분2", "생산 제품", "비고"]
    if "사출 호기" in raw.columns:
        equip = raw.loc[raw["사출 호기"].notna(), [c for c in equip_cols if c in raw.columns]].copy()
    else:
        equip = pd.DataFrame(columns=[c for c in equip_cols if c in raw.columns])
    if not equip.empty and "사출 호기" in equip.columns:
        equip["설비코드"] = (
            equip["사출 호기"]
            .astype("string")
            .fillna("")
            .astype(str)
            .str.strip()
            .str.split()
            .str[0]
        )
    else:
        equip["설비코드"] = ""

    def _is_blank(v: object) -> bool:
        if v is None:
            return True
        s = str(v).strip()
        return (s == "") or (s.lower() == "nan")

    if "생산 제품" not in equip.columns:
        equip["생산 제품"] = ""
    if "비고" not in equip.columns:
        equip["비고"] = ""
    equip["생산 제품"] = equip["생산 제품"].astype("string").fillna("").astype(str).str.strip()
    equip["비고"] = equip["비고"].astype("string").fillna("").astype(str).str.strip()

    # Rule: E(생산 제품) 공란 + F(비고) 기입 => 배정 불가
    equip["배정가능"] = ~equip.apply(lambda r: _is_blank(r.get("생산 제품")) and (not _is_blank(r.get("비고"))), axis=1)

    arrange_cols = ["제품명코드", "제품명", "구분.1"]
    if "제품명코드" in raw.columns:
        arrange = raw.loc[raw["제품명코드"].notna(), [c for c in arrange_cols if c in raw.columns]].copy()
    else:
        arrange = pd.DataFrame(columns=[c for c in arrange_cols if c in raw.columns])
    if arrange.empty:
        arrange = pd.DataFrame(columns=["제품명코드", "제품명", "구분.1"])
    for c in ["제품명코드", "제품명", "구분.1"]:
        if c not in arrange.columns:
            arrange[c] = ""
        arrange[c] = arrange[c].astype("string").fillna("").astype(str).str.strip()

    def _to_base_r(v: object) -> str:
        s = str(v or "").strip().upper()
        m = re.match(r"^(R\d+)", s)
        return m.group(1) if m else s

    # Allow users to paste full item code (e.g., R1026+03.75...) but normalize to base R (R1026)
    arrange["제품명코드"] = arrange["제품명코드"].map(_to_base_r)
    return {"equip": equip, "arrange": arrange}


@st.cache_data(show_spinner=False)
def _load_injection_machine_medians_cached(path: str, mtime: float) -> dict[str, object]:
    _ = mtime  # cache-buster when file changes
    try:
        df = pd.read_excel(path, sheet_name="생산실적", usecols=["생산일자", "공정코드", "양품수량", "기계코드"])
    except Exception:
        return {"med_by_machine": {}, "global_median": 0.0}

    df["공정코드"] = df["공정코드"].astype("string").fillna("").astype(str)
    df = df.loc[df["공정코드"].str.contains("사출", na=False)].copy()
    if df.empty:
        return {"med_by_machine": {}, "global_median": 0.0}

    df["생산일자"] = pd.to_datetime(df["생산일자"], errors="coerce").dt.date
    df["양품수량"] = pd.to_numeric(df["양품수량"], errors="coerce").fillna(0.0)
    df["기계코드"] = df["기계코드"].astype("string").fillna("").astype(str).str.strip()
    df = df.loc[df["생산일자"].notna() & df["기계코드"].ne("")].copy()
    if df.empty:
        return {"med_by_machine": {}, "global_median": 0.0}

    daily = df.groupby(["기계코드", "생산일자"], dropna=False, as_index=False)["양품수량"].sum()
    med_by_machine = daily.groupby("기계코드", dropna=False)["양품수량"].median().to_dict()
    global_median = float(daily["양품수량"].median()) if not daily.empty else 0.0
    return {"med_by_machine": med_by_machine, "global_median": global_median}


def _infer_machine_code_from_equip(equip_code: str) -> str | None:
    s = str(equip_code or "").strip().upper()
    m = re.match(r"^([A-Z])(\d+)$", s)
    if not m:
        return None
    letter, n = m.group(1), int(m.group(2))
    if letter == "A":
        return f"A형 인라인{n}호기 - 조립중합"
    if letter == "B":
        return f"B형 조립중합{n}호기"
    if letter == "C":
        return f"C형 조립중합{n}호기"
    return None


def _extract_base_r(code: object) -> str:
    s = str(code or "").strip()
    if not s:
        return ""
    m = re.match(r"^(R\d+)", s, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip().upper()
    return s.split("-", 1)[0].strip().upper()


def _coerce_date_value(x: object) -> date | None:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    t = pd.to_datetime(x, errors="coerce")
    if pd.isna(t):
        return None
    try:
        return t.date()
    except Exception:
        return None


def _choose_power_slots(powers_rem: dict[float, int], *, slots: int, slot_qty: int) -> tuple[list[float], list[int]]:
    """
    Return (slot_powers, slot_qtys) where each slot is assigned to one POWER (duplicates allowed).
    slot_qtys are capped by remaining qty; unused capacity is 0.
    """
    out_p: list[float] = []
    out_q: list[int] = []
    if slots <= 0 or slot_qty <= 0 or not powers_rem:
        return out_p, out_q

    def _best_candidates() -> list[tuple[float, int]]:
        items = [(p, int(q)) for p, q in powers_rem.items() if int(q) > 0]
        items.sort(key=lambda t: (-t[1], t[0]))
        return items

    for _ in range(int(slots)):
        cand = _best_candidates()
        if not cand:
            break
        p, rem = cand[0]
        q = int(min(rem, slot_qty))
        out_p.append(float(p))
        out_q.append(int(q))
        powers_rem[p] = int(rem - q)

    return out_p, out_q


def _build_injection_schedule(
    *,
    demand: pd.DataFrame,
    inj_equip: pd.DataFrame,
    arrange: pd.DataFrame,
    excel_path: str,
    excel_mtime: float,
    start_date: date,
    horizon_days: int,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Build 4~5 day injection schedule (2 blocks/day/equipment, 8 power slots/block).
    Returns (schedule_df, remaining_df, warnings).
    """
    warnings: list[str] = []
    if demand is None or demand.empty:
        return (
            pd.DataFrame(
                columns=[
                    "날짜",
                    "설비명",
                    "제품명코드",
                    "제품명",
                    "Block",
                    "POWER 리스트",
                    "POWER 개수",
                    "배정수량",
                    "잔여수량",
                    "세팅구분",
                ]
            ),
            pd.DataFrame(columns=["제품명코드", "제품명", "납기일", "잔여수량"]),
            ["사출 부족수량(수요)이 없습니다."],
        )

    inj_equip = inj_equip.copy()
    arrange = arrange.copy()

    if "설비코드" not in inj_equip.columns:
        warnings.append("사출 시트 설비 테이블에서 설비코드를 찾지 못했습니다.")
        return (pd.DataFrame(), pd.DataFrame(), warnings)

    arrange_map: dict[str, str] = {}
    if (not arrange.empty) and ("제품명코드" in arrange.columns) and ("구분.1" in arrange.columns):
        for _, r in arrange.iterrows():
            k = str(r.get("제품명코드") or "").strip().upper()
            v = str(r.get("구분.1") or "").strip()
            if k and v and k not in arrange_map:
                arrange_map[k] = v

    arrange_labels = sorted({v for v in arrange_map.values() if v}, key=lambda x: -len(x))
    inj_equip["라인구분"] = ""
    if arrange_labels and ("사출 호기" in inj_equip.columns):
        text = inj_equip["사출 호기"].astype("string").fillna("").astype(str)
        for lbl in arrange_labels:
            mask = inj_equip["라인구분"].eq("") & text.str.contains(re.escape(lbl), na=False)
            inj_equip.loc[mask, "라인구분"] = lbl

    def _parse_running_base(v: object) -> str:
        s = str(v or "").strip()
        if not s:
            return ""
        if re.match(r"^R\d{3,}", s, flags=re.IGNORECASE):
            return _extract_base_r(s).upper()
        # 제품명(판매명) 기반 매칭은 금지: 엑셀 사출 시트 E열에는 base R코드를 입력해야 함
        return ""

    inj_equip["현재제품"] = inj_equip.get("생산 제품", "").map(_parse_running_base) if "생산 제품" in inj_equip.columns else ""
    inj_equip["설비명"] = inj_equip["설비코드"].astype("string").fillna("").astype(str).str.strip().str.upper()

    if "배정가능" in inj_equip.columns:
        usable = inj_equip.loc[inj_equip["배정가능"].fillna(False)].copy()
    else:
        usable = inj_equip.copy()
    usable = usable.loc[usable["설비명"].ne("")].copy()

    if usable.empty:
        warnings.append("사출 시트에서 배정 가능한 설비가 없습니다(E 공란 + F 기입 설비는 배정 불가).")
        return (pd.DataFrame(), pd.DataFrame(), warnings)

    med_info = _load_injection_machine_medians_cached(excel_path, float(excel_mtime))
    med_by_machine: dict[str, float] = med_info.get("med_by_machine", {}) if isinstance(med_info, dict) else {}
    global_med = float(med_info.get("global_median", 0.0)) if isinstance(med_info, dict) else 0.0
    if global_med <= 0:
        global_med = 24000.0
        warnings.append("생산실적 기반 설비 생산량 추정에 실패하여 기본 CAPA(일 24000ea/설비)를 사용합니다.")

    def _equip_day_capa(equip_name: str) -> int:
        mc = _infer_machine_code_from_equip(equip_name)
        if mc and mc in med_by_machine:
            v = float(med_by_machine.get(mc) or 0.0)
            return int(max(0, round(v)))
        return int(max(0, round(global_med)))

    horizon_days = int(horizon_days)
    if horizon_days <= 0:
        horizon_days = 5
    days = [start_date + timedelta(days=i) for i in range(horizon_days)]

    work = demand.copy()
    for c in ["제품코드", "품명", "POWER", "납기일"]:
        if c not in work.columns:
            work[c] = ""
    if "사출" not in work.columns:
        work["사출"] = 0

    work["제품코드"] = work["제품코드"].astype("string").fillna("").astype(str).str.strip()
    work = work.loc[work["제품코드"].ne("")].copy()
    work = work.loc[work["제품코드"].str.upper().str.startswith("R")].copy()
    work["제품명코드"] = work["제품코드"].map(_extract_base_r).astype("string").fillna("").astype(str).str.strip().str.upper()
    work["사출"] = pd.to_numeric(work["사출"], errors="coerce").fillna(0).astype(int)
    work = work.loc[work["사출"].gt(0) & work["제품명코드"].ne("")].copy()
    if work.empty:
        return (pd.DataFrame(), pd.DataFrame(), ["R코드(사출) 부족수량(수요)이 없습니다."])

    work["POWER_num"] = pd.to_numeric(work["POWER"], errors="coerce")
    work["_due"] = work["납기일"].map(_coerce_date_value)

    product_info: dict[str, dict[str, object]] = {}
    for _, r in work.iterrows():
        base_r = str(r["제품명코드"] or "").strip().upper()
        if not base_r:
            continue
        due = r.get("_due", None)
        name = str(r.get("품명") or "").strip()
        p = r.get("POWER_num", None)
        if p is None or (isinstance(p, float) and math.isnan(p)):
            continue
        need = int(r.get("사출") or 0)
        if need <= 0:
            continue

        info = product_info.get(base_r)
        if info is None:
            info = {
                "제품명코드": base_r,
                "제품명": name,
                "납기일": due,
                "powers": {},
                "라인구분": arrange_map.get(base_r, ""),
            }
            product_info[base_r] = info
        if name and (not str(info.get("제품명") or "").strip()):
            info["제품명"] = name
        if due is not None:
            cur_due = info.get("납기일")
            if (cur_due is None) or (isinstance(cur_due, date) and due < cur_due):
                info["납기일"] = due
        powers: dict[float, int] = info["powers"]  # type: ignore[assignment]
        powers[float(p)] = int(powers.get(float(p), 0) + need)

    for base_r in list(product_info.keys()):
        if not str(product_info[base_r].get("라인구분") or "").strip():
            warnings.append(f"어레인지 누락: {base_r} (사출 시트 I~K에 제품명코드 매핑 필요)")
            product_info.pop(base_r, None)

    if not product_info:
        return (pd.DataFrame(), pd.DataFrame(), warnings or ["어레인지 매칭 가능한 제품이 없습니다."])

    def _product_remaining(base_r: str) -> int:
        info = product_info.get(base_r) or {}
        powers = info.get("powers") or {}
        return int(sum(int(v) for v in powers.values()))

    def _product_due(base_r: str) -> date:
        d = product_info.get(base_r, {}).get("납기일")
        return d if isinstance(d, date) else date(2099, 12, 31)

    def _eligible_products_for_equipment(equip_row: pd.Series) -> list[str]:
        line = str(equip_row.get("라인구분") or "").strip()
        if not line:
            return []
        out = [
            k
            for k, v in product_info.items()
            if str(v.get("라인구분") or "").strip() == line and _product_remaining(k) > 0
        ]
        out.sort(key=lambda k: (_product_due(k), -_product_remaining(k), k))
        return out

    equip_last: dict[str, str] = {
        str(r["설비명"]): str(r.get("현재제품") or "").strip().upper() for _, r in usable.iterrows()
    }
    equip_affinity: dict[str, str] = {}

    rows: list[dict[str, object]] = []

    for day in days:
        for _, er in usable.iterrows():
            equip_name = str(er.get("설비명") or "").strip().upper()
            if not equip_name:
                continue
            day_capa = _equip_day_capa(equip_name)
            block_capa = max(0, int(round(day_capa / 2.0)))
            slot_qty = max(1, int(round(block_capa / 8.0))) if block_capa > 0 else 0

            prev_prod = str(equip_last.get(equip_name, "") or "").strip().upper()
            affinity = str(equip_affinity.get(equip_name, "") or "").strip().upper()
            candidates = _eligible_products_for_equipment(er)

            def _pick_product(prefer: str | None) -> str:
                if prefer and prefer in candidates and _product_remaining(prefer) > 0:
                    return prefer
                if affinity and affinity in candidates and _product_remaining(affinity) > 0:
                    return affinity
                return candidates[0] if candidates else ""

            prod1 = _pick_product(prev_prod)
            if prod1:
                equip_affinity.setdefault(equip_name, prod1)

            def _setting_label(*, block: int, cur: str, prev_day: str, prev_block: str) -> str:
                cur = str(cur or "").strip().upper()
                prev_day = str(prev_day or "").strip().upper()
                prev_block = str(prev_block or "").strip().upper()
                if not cur:
                    return "유휴"
                if block == 1:
                    if not prev_day:
                        return "신규세팅"
                    return "세팅유지" if cur == prev_day else "잡체인지"
                if prev_block and (cur == prev_block):
                    return "세팅유지"
                if not prev_block and prev_day and (cur == prev_day):
                    return "세팅유지"
                if not prev_block and (not prev_day):
                    return "신규세팅"
                return "잡체인지"

            prev_block_prod = ""
            for block in (1, 2):
                if block == 1:
                    cur_prod = prod1
                else:
                    if prev_block_prod and _product_remaining(prev_block_prod) > 0:
                        cur_prod = prev_block_prod
                    else:
                        candidates = _eligible_products_for_equipment(er)
                        cur_prod = _pick_product(prev_prod if not prev_block_prod else None)
                        if cur_prod:
                            equip_affinity.setdefault(equip_name, cur_prod)

                powers_list: list[str] = []
                assign_qty = 0
                rem_after = 0
                prod_name = ""
                if cur_prod:
                    info = product_info.get(cur_prod) or {}
                    prod_name = str(info.get("제품명") or "").strip()
                    powers: dict[float, int] = info.get("powers") or {}
                    slot_p, slot_q = _choose_power_slots(powers, slots=8, slot_qty=slot_qty)
                    assign_qty = int(sum(slot_q))
                    rem_after = _product_remaining(cur_prod)
                    powers_list = [f"{p:+.2f}" for p in slot_p]
                    if len(powers_list) < 8:
                        powers_list += [""] * (8 - len(powers_list))
                else:
                    powers_list = [""] * 8

                setting = _setting_label(block=block, cur=cur_prod, prev_day=prev_prod, prev_block=prev_block_prod)
                rows.append(
                    {
                        "날짜": day,
                        "설비명": equip_name,
                        "제품명코드": cur_prod,
                        "제품명": prod_name,
                        "Block": block,
                        "POWER 리스트": ", ".join([p for p in powers_list if str(p).strip()]),
                        "POWER 개수": int(sum(1 for p in powers_list if str(p).strip())),
                        "배정수량": int(assign_qty),
                        "잔여수량": int(rem_after) if cur_prod else 0,
                        "세팅구분": setting,
                    }
                )
                if cur_prod:
                    prev_block_prod = cur_prod

            equip_last[equip_name] = prev_block_prod or prev_prod

    sched = pd.DataFrame(rows)

    rem_rows: list[dict[str, object]] = []
    for base_r, info in sorted(product_info.items(), key=lambda kv: (_product_due(kv[0]), kv[0])):
        rem = _product_remaining(base_r)
        if rem <= 0:
            continue
        rem_rows.append(
            {
                "제품명코드": base_r,
                "제품명": str(info.get("제품명") or "").strip(),
                "납기일": info.get("납기일"),
                "잔여수량": rem,
            }
        )
    remaining = pd.DataFrame(rem_rows)
    return (sched, remaining, warnings)


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
    numeric_cols = [c for c in [*DEFAULT_STAGE_COLS, "필요수량"] if c in df.columns]
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
    col = df["납기일"]
    due = col if pd.api.types.is_datetime64_any_dtype(col) else pd.to_datetime(col, errors="coerce")
    mask = due.dt.date.le(end)
    return df.loc[mask].copy()


def _compute_capa_table_from_prod_daily(
    prod_daily: pd.DataFrame,
    *,
    n_run_days: int,
    as_of: date,
) -> pd.DataFrame:
    """
    CAPA 산정: 최근 N개 '가동일'(실적>0) 기준 일평균 양품수량.
    - prod_daily는 (공정, 생산일자, 양품) 일별 집계 CSV 기준
    - as_of는 기준일(전일까지)로 자른다
    """
    if prod_daily is None or prod_daily.empty:
        return pd.DataFrame(columns=["공정", "CAPA", "capa_days", "last_run_date"])

    df = prod_daily
    if "생산일자" not in df.columns or "공정" not in df.columns or "양품" not in df.columns:
        return pd.DataFrame(columns=["공정", "CAPA", "capa_days", "last_run_date"])

    df = df.copy()
    df["생산일자"] = pd.to_datetime(df["생산일자"], errors="coerce")
    df["양품"] = pd.to_numeric(df["양품"], errors="coerce").fillna(0)
    df["공정"] = df["공정"].astype("string").fillna("").str.strip()

    # Filter once (avoid repeated .copy()).
    mask = df["공정"].ne("") & df["생산일자"].notna() & df["양품"].gt(0)
    if mask.any():
        df = df.loc[mask]
    else:
        return pd.DataFrame(columns=["공정", "CAPA", "capa_days", "last_run_date"])
    df = df.loc[df["생산일자"].dt.date.le(as_of)]
    if df.empty:
        return pd.DataFrame(columns=["공정", "CAPA", "capa_days", "last_run_date"])

    df = df.sort_values(["공정", "생산일자"], ascending=[True, True])

    rows: list[dict] = []
    for proc, g in df.groupby("공정", dropna=False):
        # 일별 집계 형태를 다시 보장(동일일자 다건 대비)
        by_day = (
            g.groupby(g["생산일자"].dt.date, dropna=False)["양품"]
            .sum()
            .sort_index()
        )
        if by_day.empty:
            continue
        last_days = by_day.tail(max(1, int(n_run_days)))
        rows.append(
            {
                "공정": str(proc),
                "CAPA": float(last_days.mean()),
                "capa_days": int(last_days.shape[0]),
                "last_run_date": pd.to_datetime(last_days.index[-1]),
            }
        )

    out = pd.DataFrame(rows)
    if not out.empty:
        out["CAPA"] = pd.to_numeric(out["CAPA"], errors="coerce").fillna(0)
    return out


def _grade_from_days(*, required_days: float, remaining_days: float, buffer_days: float) -> str:
    if required_days is None or remaining_days is None:
        return "NO_DUE"
    try:
        r = float(required_days)
        rem = float(remaining_days)
        buf = float(buffer_days)
    except Exception:
        return "NO_DUE"
    if rem < 0:
        rem = 0.0
    if r == 0.0:
        return "GREEN"
    if not pd.notna(r):
        return "NO_CAPA"
    if r == float("inf"):
        return "NO_CAPA"
    if r > rem:
        return "RED"
    if (rem - r) <= max(0.0, buf):
        return "YELLOW"
    return "GREEN"


def _build_order_risk_table(
    order_df: pd.DataFrame,
    capa_table: pd.DataFrame,
    *,
    today: date,
    buffer_days: float,
    start_offset_days: int = 1,
) -> pd.DataFrame:
    """
    수주별 리스크 산출.
    - 부족수량(order_df의 공정 컬럼) vs CAPA(공정별)
    - 수주 등급은 최악 공정 기준
    - 완료예상/지연은 누수규격 기준(연속 24/7 가정)
    """
    if order_df is None or order_df.empty:
        return pd.DataFrame()

    base = order_df.copy()
    if "납기일" in base.columns:
        due_col = base["납기일"]
        base["납기일"] = (
            due_col if pd.api.types.is_datetime64_any_dtype(due_col) else pd.to_datetime(due_col, errors="coerce")
        )

    stage_cols = [c for c in DEFAULT_STAGE_COLS if c in base.columns]
    if not stage_cols:
        return pd.DataFrame()

    for c in stage_cols:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    # Remaining calendar days (24/7 운영 전제)
    due_dt = base["납기일"] if "납기일" in base.columns else pd.Series([pd.NaT] * len(base))
    base["_due_date"] = due_dt.dt.date
    base["남은일수_raw"] = base["_due_date"].map(lambda d: (d - today).days if isinstance(d, date) else None)
    base["남은일수"] = base["남은일수_raw"].map(lambda x: max(0, int(x)) if isinstance(x, int) else (max(0, int(x)) if isinstance(x, float) else None))

    capa_map = {}
    capa_days_map: dict[str, int] = {}
    if capa_table is not None and (not capa_table.empty) and "공정" in capa_table.columns:
        tmp = capa_table.copy()
        tmp["공정"] = tmp["공정"].astype("string").fillna("").str.strip()
        if "CAPA" in tmp.columns:
            tmp["CAPA"] = pd.to_numeric(tmp["CAPA"], errors="coerce").fillna(0)
        if "capa_days" in tmp.columns:
            tmp["capa_days"] = pd.to_numeric(tmp["capa_days"], errors="coerce").fillna(0).astype(int)
        capa_map = {str(r["공정"]): float(r.get("CAPA", 0) or 0) for r in tmp.to_dict("records")}
        capa_days_map = {str(r["공정"]): int(r.get("capa_days", 0) or 0) for r in tmp.to_dict("records")}

    meta_cols = [c for c in ["이니셜", "수주번호", "신규분류 요약코드", "품명", "납기일"] if c in base.columns]
    melt = base[meta_cols + ["남은일수_raw", "남은일수"] + stage_cols].melt(
        id_vars=meta_cols + ["남은일수_raw", "남은일수"],
        value_vars=stage_cols,
        var_name="공정",
        value_name="부족수량",
    )
    melt["부족수량"] = pd.to_numeric(melt["부족수량"], errors="coerce").fillna(0)
    melt = melt.loc[melt["부족수량"].gt(0)].copy()
    if melt.empty:
        return pd.DataFrame()

    melt["CAPA"] = melt["공정"].map(lambda x: capa_map.get(str(x), 0.0))
    melt["CAPA_days"] = melt["공정"].map(lambda x: capa_days_map.get(str(x), 0))
    melt["필요일수"] = melt.apply(
        lambda r: (float(r["부족수량"]) / float(r["CAPA"])) if float(r["CAPA"]) > 0 else float("inf"),
        axis=1,
    )
    melt["슬랙"] = melt.apply(
        lambda r: (float(r["남은일수"]) - float(r["필요일수"])) if r["남은일수"] is not None else float("nan"),
        axis=1,
    )
    melt["등급"] = melt.apply(
        lambda r: _grade_from_days(required_days=r["필요일수"], remaining_days=r["남은일수"], buffer_days=buffer_days),
        axis=1,
    )

    grade_rank = {"RED": 3, "YELLOW": 2, "GREEN": 1, "NO_CAPA": 4, "NO_DUE": 0}
    melt["_grade_rank"] = melt["등급"].map(lambda x: grade_rank.get(str(x), 0))

    group_key = [c for c in ["이니셜", "수주번호", "신규분류 요약코드"] if c in melt.columns]
    if not group_key:
        group_key = ["수주번호"] if "수주번호" in melt.columns else []
    if not group_key:
        return pd.DataFrame()

    # Worst grade per order: NO_CAPA > RED > YELLOW > GREEN
    worst = (
        melt.groupby(group_key, dropna=False)["_grade_rank"]
        .max()
        .reset_index()
        .rename(columns={"_grade_rank": "_worst_rank"})
    )

    proc_order = {"사출": 0, "분리": 1, "하이드레이션": 2, "접착": 3, "누수규격": 4}
    melt["_proc_order"] = melt["공정"].map(lambda x: proc_order.get(str(x), 999))

    # Start process: 가장 앞 공정(사출->...->누수규격) 중 부족수량>0인 공정
    idx_start = melt.groupby(group_key, dropna=False)["_proc_order"].idxmin()
    start_proc = melt.loc[idx_start, group_key + ["공정", "부족수량", "CAPA", "필요일수", "남은일수", "등급"]].copy()
    start_proc = start_proc.rename(
        columns={
            "공정": "시작공정",
            "부족수량": "시작_부족수량",
            "CAPA": "시작_CAPA",
            "필요일수": "시작_필요일수",
            "남은일수": "시작_남은일수",
            "등급": "시작_등급",
        }
    )

    # Bottleneck: (등급 심각도) -> (슬랙 최소) -> (공정순서) 우선
    def _pick_bn_index(g: pd.DataFrame) -> int:
        gg = g.sort_values(
            by=["_grade_rank", "슬랙", "_proc_order"],
            ascending=[False, True, True],
            na_position="last",
        )
        return int(gg.index[0])

    try:
        idx_bn = melt.groupby(group_key, dropna=False, group_keys=False).apply(_pick_bn_index, include_groups=False)
    except TypeError:
        # pandas<2.2 compatibility
        idx_bn = melt.groupby(group_key, dropna=False, group_keys=False).apply(_pick_bn_index)
    bottleneck = melt.loc[idx_bn, group_key + ["공정", "부족수량", "CAPA", "CAPA_days", "필요일수", "남은일수", "등급", "슬랙"]].copy()
    bottleneck = bottleneck.rename(
        columns={
            "공정": "병목공정",
            "부족수량": "병목_부족수량",
            "CAPA": "병목_CAPA",
            "CAPA_days": "병목_CAPA_days",
            "필요일수": "병목_필요일수",
            "남은일수": "병목_남은일수",
            "등급": "병목_등급",
            "슬랙": "병목_슬랙",
        }
    )

    # Join back representative meta columns
    rep = base.copy()
    rep_key = group_key.copy()
    rep_cols = [c for c in ["품명", "납기일"] if c in rep.columns]
    if rep_cols:
        rep = rep.groupby(rep_key, dropna=False, as_index=False).agg({c: "first" for c in rep_cols})
    else:
        rep = rep[group_key].drop_duplicates()

    out = (
        rep.merge(worst, on=group_key, how="left")
        .merge(start_proc, on=group_key, how="left")
        .merge(bottleneck, on=group_key, how="left")
    )

    def _rank_to_grade(n) -> str:
        try:
            n = int(n)
        except Exception:
            return "GREEN"
        inv = {v: k for k, v in grade_rank.items()}
        # Prefer business order: NO_CAPA as RED-equivalent alert
        g = inv.get(n, "GREEN")
        return "RED" if g == "NO_CAPA" else g

    out["리스크등급"] = out["_worst_rank"].map(_rank_to_grade)

    # Order-level shortage summary for risk display/forecast:
    # - 후공정 타관(사출/분리까지만 부족이고, 하이드/접착/누수는 0)
    # - 누수규격(표시용 필요수량): 기본은 누수규격 부족수량, 타관이면 분리(없으면 사출)
    # - 막힘공정(표시용): 공정간 부족 "증분"과 CAPA를 반영해 선택
    # - 완료예정일: 게이트공정 기준(타관이면 분리, 아니면 누수규격; 24/7 연속운영 가정)
    stage_agg = base[group_key + stage_cols].copy()
    stage_agg = stage_agg.groupby(group_key, dropna=False, as_index=False)[stage_cols].sum(numeric_only=True)
    for c in stage_cols:
        stage_agg[c] = pd.to_numeric(stage_agg[c], errors="coerce").fillna(0)
    s10 = stage_agg["사출"] if "사출" in stage_agg.columns else 0
    s20 = stage_agg["분리"] if "분리" in stage_agg.columns else 0
    s45 = stage_agg["하이드레이션"] if "하이드레이션" in stage_agg.columns else 0
    s55 = stage_agg["접착"] if "접착" in stage_agg.columns else 0
    s80 = stage_agg["누수규격"] if "누수규격" in stage_agg.columns else 0
    stage_agg["후공정_타관"] = ((s10 > 0) | (s20 > 0)) & (s45 <= 0) & (s55 <= 0) & (s80 <= 0)

    # 표시용 "누수규격"(필요수량)
    base_need = stage_agg["누수규격"] if "누수규격" in stage_agg.columns else 0
    alt_need = stage_agg["분리"] if "분리" in stage_agg.columns else (stage_agg["사출"] if "사출" in stage_agg.columns else 0)
    stage_agg["누수규격"] = base_need.where(~stage_agg["후공정_타관"], alt_need)

    # 부족 "증분" 계산(부족이 실제로 새로 생기는 구간)
    d10 = s10
    d20 = (s20 - s10).clip(lower=0)
    d45 = (s45 - s20).clip(lower=0)
    d55 = (s55 - s45).clip(lower=0)
    d80 = (s80 - s55).clip(lower=0)
    stage_agg["_d10"] = d10
    stage_agg["_d20"] = d20
    stage_agg["_d45"] = d45
    stage_agg["_d55"] = d55
    stage_agg["_d80"] = d80

    # 막힘공정 선택: (증분/CAPA) 최대 (CAPA=0이면 inf), 동률이면 증분 큰 공정, 그 다음 공정순서
    def _pick_block_stage(row) -> str:
        candidates = [
            ("사출", float(row.get("_d10") or 0)),
            ("분리", float(row.get("_d20") or 0)),
            ("하이드레이션", float(row.get("_d45") or 0)),
            ("접착", float(row.get("_d55") or 0)),
            ("누수규격", float(row.get("_d80") or 0)),
        ]
        best = ""
        best_score = -1.0
        best_delta = -1.0
        best_ord = 999
        for proc, delta in candidates:
            if delta <= 0:
                continue
            capa = float(capa_map.get(proc, 0.0) or 0.0)
            if capa <= 0:
                score = float("inf")
            else:
                score = delta / capa
            ordv = proc_order.get(proc, 999)
            if (
                (score > best_score)
                or (score == best_score and delta > best_delta)
                or (score == best_score and delta == best_delta and ordv < best_ord)
            ):
                best = proc
                best_score = score
                best_delta = delta
                best_ord = ordv
        # fallback: no incremental diff -> treat as first shortage stage
        if not best:
            for proc in DEFAULT_STAGE_COLS:
                try:
                    if float(row.get(proc, 0) or 0) > 0:
                        return proc
                except Exception:
                    continue
        return best

    stage_agg["막힘공정"] = stage_agg.apply(_pick_block_stage, axis=1)

    # 완료예정일(요청수량 기반 + 공정반영형):
    # - 공정별 필요수량(사출/분리/하이드/접착/누수)을 그대로 사용
    # - 공정 간 WIP/대기일은 공정당 1일로 고정
    # - 납기순(EDD)으로 수주를 정렬해 공정별 1라인(flow-shop) 스케줄링으로 완료예정일 산정
    due_agg = base[group_key + (["납기일"] if "납기일" in base.columns else [])].copy()
    if "납기일" in due_agg.columns:
        due_agg["납기일"] = pd.to_datetime(due_agg["납기일"], errors="coerce")
        due_agg = due_agg.groupby(group_key, dropna=False, as_index=False)["납기일"].min()
    else:
        due_agg = base[group_key].drop_duplicates()
        due_agg["납기일"] = pd.NaT

    stage_agg = stage_agg.merge(due_agg, on=group_key, how="left")

    # Gate stage for completion date: transfer -> 분리(없으면 사출), else 누수규격(없으면 마지막 부족 공정)
    def _gate_for_done(row) -> str:
        if bool(row.get("후공정_타관")):
            if "분리" in stage_cols and float(row.get("분리", 0) or 0) > 0:
                return "분리"
            if "사출" in stage_cols and float(row.get("사출", 0) or 0) > 0:
                return "사출"
        if "누수규격" in stage_cols and float(row.get("누수규격", 0) or 0) > 0:
            return "누수규격"
        last = ""
        last_ord = -1
        for p in DEFAULT_STAGE_COLS:
            if p in stage_cols and float(row.get(p, 0) or 0) > 0 and proc_order.get(p, -1) >= last_ord:
                last = p
                last_ord = proc_order.get(p, -1)
        return last

    stage_agg["게이트공정"] = stage_agg.apply(_gate_for_done, axis=1)

    if "납기일" in stage_agg.columns:
        due_col2 = stage_agg["납기일"]
        stage_agg["납기일"] = (
            due_col2 if pd.api.types.is_datetime64_any_dtype(due_col2) else pd.to_datetime(due_col2, errors="coerce")
        )

    sched = stage_agg.copy()
    sched = sched.sort_values(["납기일"], ascending=[True], na_position="last")

    proc_path_all = ["사출", "분리", "하이드레이션", "접착", "누수규격"]
    proc_path_all = [p for p in proc_path_all if p in stage_cols]

    completion: dict[str, list[float]] = {p: [] for p in proc_path_all}
    available: dict[str, float] = {p: 0.0 for p in proc_path_all}

    for _, row in sched.iterrows():
        is_transfer = bool(row.get("후공정_타관"))
        path = ["사출", "분리"] if is_transfer else proc_path_all
        prev_done = 0.0
        for p in proc_path_all:
            qty = float(row.get(p, 0) or 0)
            capa = float(capa_map.get(p, 0.0) or 0.0)
            # If process not in this order's path or has no qty, treat as no-op.
            if p not in path or qty <= 0:
                done = max(available[p], prev_done)
                completion[p].append(done)
                prev_done = done
                continue
            if capa <= 0:
                done = float("inf")
                completion[p].append(done)
                prev_done = done
                available[p] = max(available[p], done)
                continue
            start = max(available[p], prev_done + float(RISK_WIP_DAYS_PER_PROCESS))
            dur = qty / capa
            done = start + dur
            completion[p].append(done)
            available[p] = done
            prev_done = done

    for p in proc_path_all:
        sched[f"_done_{p}"] = completion[p]

    def _done_day(row) -> float:
        gate = str(row.get("게이트공정") or "").strip()
        if not gate:
            return float("nan")
        try:
            return float(row.get(f"_done_{gate}", float("nan")))
        except Exception:
            return float("nan")

    sched["_done_days"] = sched.apply(_done_day, axis=1)

    def _to_date(x) -> pd.Timestamp:
        try:
            v = float(x)
        except Exception:
            return pd.NaT
        if not pd.notna(v) or v == float("inf"):
            return pd.NaT
        d = int(math.ceil(max(0.0, v + float(start_offset_days))))
        return pd.Timestamp(today + timedelta(days=d))

    sched["완료예정일"] = sched["_done_days"].map(_to_date)
    # NOTE: do not pre-create stage_agg["완료예정일"] before merge, otherwise pandas will suffix
    # columns (완료예정일_x/완료예정일_y) and downstream column selection will break.
    stage_agg = stage_agg.merge(sched[group_key + ["완료예정일"]], on=group_key, how="left")

    out = out.merge(
        stage_agg[group_key + ["후공정_타관", "누수규격", "막힘공정", "게이트공정", "완료예정일"]],
        on=group_key,
        how="left",
    )
    out["후공정_타관"] = out["후공정_타관"].fillna(False).astype(bool)
    if "누수규격" in out.columns:
        out["누수규격"] = pd.to_numeric(out["누수규격"], errors="coerce").fillna(0).astype(int)

    # 리스크가 없는(GREEN) 경우 완료예정일은 납기일로 표시(혼란 방지).
    if "납기일" in out.columns:
        out["납기일"] = pd.to_datetime(out["납기일"], errors="coerce")
    if "완료예정일" in out.columns:
        out["완료예정일"] = pd.to_datetime(out["완료예정일"], errors="coerce")
    if "리스크등급" in out.columns and "완료예정일" in out.columns and "납기일" in out.columns:
        mask_green = out["리스크등급"].astype(str).eq("GREEN") & out["납기일"].notna()
        out.loc[mask_green, "완료예정일"] = out.loc[mask_green, "납기일"]

    def _reason_row(r) -> str:
        if pd.isna(r.get("병목공정")) and pd.isna(r.get("시작공정")):
            return ""
        bn_proc_name = str(r.get("막힘공정") or r.get("병목공정") or "").strip()
        if not bn_proc_name:
            return ""

        bn_capa = float(capa_map.get(bn_proc_name, 0.0) or 0.0)
        grade = str(r.get("리스크등급") or "").strip()

        if grade == "GREEN":
            return ""

        if bn_capa <= 0:
            msg = f"{bn_proc_name} 병목(CAPA=0)"
        else:
            if grade == "RED":
                msg = f"{bn_proc_name} 병목(납기내 불가)"
            elif grade == "YELLOW":
                msg = f"{bn_proc_name} 주의(여유부족)"
            else:
                msg = f"{bn_proc_name} 정상"

        if bool(r.get("후공정_타관")):
            return f"{msg} / 후공정 타관"
        return msg

    out["리스크사유"] = out.apply(_reason_row, axis=1)

    # Priority sort: grade, due date, delay
    sort_grade = {"RED": 0, "YELLOW": 1, "GREEN": 2}
    out["_g"] = out["리스크등급"].map(lambda x: sort_grade.get(str(x), 9))
    if "납기일" in out.columns:
        out = out.sort_values(["_g", "납기일"], ascending=[True, True], na_position="last")
    else:
        out = out.sort_values(["_g"], ascending=[True], na_position="last")
    out.insert(0, "우선순위", range(1, len(out) + 1))
    out = out.drop(columns=[c for c in ["_g", "_worst_rank"] if c in out.columns])
    return out


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
    try:
        st.caption(f"앱 코드 수정시각: `{_file_mtime_label(__file__)}`")
    except Exception:
        pass
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

    out_dir = OUT_DIR
    status = _outputs_status(excel_path=excel_path, out_dir=out_dir)
    if not status.get("ok"):
        st.error(f"데이터 로딩 실패: {status.get('reason') or '-'}")
        st.stop()

    if status.get("needs_regen"):
        with st.spinner("엑셀 변경 감지: 데이터 생성 중..."):
            ensure = _ensure_latest_outputs(excel_path=excel_path, out_dir=out_dir)
    else:
        ensure = status

    if not ensure.get("ok"):
        st.error(f"데이터 생성 실패: {ensure.get('reason')}")
        st.stop()
    if ensure.get("regenerated"):
        st.cache_data.clear()

    due_csv = str(ensure["due_csv"])
    detail_csv = str(ensure["detail_csv"])
    equip_code_target_csv = ensure.get("equip_code_target_csv")
    prod_daily_csv = ensure.get("prod_daily_csv")

    detail_for_map: pd.DataFrame | None = None
    try:
        detail_for_map = _load_order_detail_prepared(detail_csv, os.path.getmtime(detail_csv))
    except Exception:
        detail_for_map = None

    equip_code_target_df: pd.DataFrame | None = None
    try:
        if equip_code_target_csv:
            equip_code_target_df = _load_equip_code_min_target_csv(
                str(equip_code_target_csv), os.path.getmtime(str(equip_code_target_csv))
            )
        else:
            equip_code_target_df = None
    except Exception:
        equip_code_target_df = None

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
            "최소목표일",
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

            if (
                detail_for_map is not None
                and equip_code_target_df is not None
                and (not equip_code_target_df.empty)
                and "제품 코드" in detail_for_map.columns
                and "최소목표일" in equip_code_target_df.columns
                and "제품 코드" in equip_code_target_df.columns
                and "공정" in equip_code_target_df.columns
            ):
                d = detail_for_map.copy()
                if "납기일" in d.columns:
                    d["납기일"] = pd.to_datetime(d["납기일"], errors="coerce")
                d["제품 코드"] = d["제품 코드"].astype("string").fillna("").str.strip()
                if allowed_prefixes:
                    prefixes = [str(p).strip().upper() for p in allowed_prefixes if str(p).strip()]
                    if prefixes:
                        d = d.loc[d["제품 코드"].map(lambda x: any(str(x).upper().startswith(p) for p in prefixes))]
                t = equip_code_target_df.loc[equip_code_target_df["공정"].astype("string") == process_only, ["제품 코드", "최소목표일"]]
                d = d.merge(t, on="제품 코드", how="left")
                key_candidates = ["신규분류 요약코드", "제품군", "납기일"]
                key_cols = [c for c in key_candidates if c in view.columns and c in d.columns]
                if key_cols:
                    tgt = d.groupby(key_cols, dropna=False, as_index=False).agg(최소목표일=("최소목표일", "min"))
                    view = view.merge(tgt, on=key_cols, how="left")
            if "최소목표일" not in view.columns:
                view["최소목표일"] = pd.NaT

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
        cols += ["POWER"]
        cols += ["납기일"]
        if process_only:
            cols.append("최소목표일")
        cols += stage_cols_raw
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
            "최소목표일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
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

    view_options = ["납기별 상세", "공정별 보기", "수주별 현황", "리스크", "사출 계획"]
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
        if view_mode == "사출 계획":
            st.session_state["inj_due_quick"] = "해제"
            st.session_state["inj_due_end"] = _today_kst()
            st.session_state["_prev_inj_due_quick"] = "해제"
        if view_mode == "수주별 현황":
            # Always reset due-date filter when entering order view.
            st.session_state["order_due_quick"] = "해제"
            st.session_state["order_due_end"] = _today_kst()
            st.session_state["_prev_order_due_quick"] = "해제"
        if view_mode == "리스크":
            st.session_state["risk_due_quick"] = "해제"
            st.session_state["risk_due_end"] = _today_kst()
            st.session_state["_prev_risk_due_quick"] = "해제"
            st.session_state["risk_grade_pill"] = ["RED", "YELLOW"]
        st.session_state["_prev_view_mode"] = view_mode

    process_only = None
    if view_mode == "납기별 상세":
        # Due date end quick-picks for due view.
        due_quick_options = ["해제", "직접", "당월", "+7일", "+14일"]
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
        if due_quick == "당월":
            due_default_end = _end_of_month(_today_kst())
        elif due_quick == "+7일":
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
        proc_quick_options = ["해제", "직접", "당월", "+7일", "+14일"]
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
        if proc_quick == "당월":
            proc_default_end = _end_of_month(_today_kst())
        elif proc_quick == "+7일":
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

    if view_mode == "사출 계획":
        inj_quick_options = ["해제", "직접", "당월", "+7일", "+14일"]
        _pre_widget_single_select_fix(key="inj_due_quick", default="해제", options=inj_quick_options)
        inj_quick_raw = st.pills(
            "납기일 종료 (빠른 선택)",
            options=inj_quick_options,
            default="해제",
            key="inj_due_quick",
            selection_mode="single",
            on_change=_on_change_single_select,
            args=("inj_due_quick", "해제", inj_quick_options),
            label_visibility="collapsed",
        )
        inj_quick = _coerce_single_value(inj_quick_raw, default="해제", options=inj_quick_options)
        if inj_quick == "당월":
            inj_default_end = _end_of_month(_today_kst())
        elif inj_quick == "+7일":
            inj_default_end = _today_kst() + timedelta(days=7)
        elif inj_quick == "+14일":
            inj_default_end = _today_kst() + timedelta(days=14)
        else:
            inj_default_end = _today_kst()

        prev_inj_quick = st.session_state.get("_prev_inj_due_quick")
        if prev_inj_quick != inj_quick:
            st.session_state["inj_due_end"] = inj_default_end
            st.session_state["_prev_inj_due_quick"] = inj_quick

        inj_end_date = st.date_input(
            "납기일 종료",
            value=st.session_state.get("inj_due_end", inj_default_end),
            key="inj_due_end",
            disabled=(inj_quick == "해제"),
        )

    # 분류 pills (view-mode별로 totals 계산 데이터가 다름)
    order_df_all: pd.DataFrame | None = None
    if view_mode in ("수주별 현황", "리스크"):
        if detail_for_map is None:
            st.error("수주별/리스크 데이터가 없습니다. 엑셀 시트/컬럼을 확인하세요.")
            st.stop()
        order_df = _load_order_detail_grouped(detail_csv, os.path.getmtime(detail_csv))
        if view_mode == "리스크":
            # 리스크의 완료예정일(스케줄)은 전체 backlog 기준으로 계산해야 하므로,
            # 표시 필터(납기일 종료) 적용 전 원본을 별도로 보관한다.
            order_df_all = order_df.copy()

        # Due date end quick-picks
        quick_key = "order_due_quick" if view_mode == "수주별 현황" else "risk_due_quick"
        prev_quick_key = "_prev_order_due_quick" if view_mode == "수주별 현황" else "_prev_risk_due_quick"
        end_key = "order_due_end" if view_mode == "수주별 현황" else "risk_due_end"

        quick_options = ["해제", "직접", "당월", "+7일", "+14일"]
        default_quick = "해제"
        _pre_widget_single_select_fix(key=quick_key, default=default_quick, options=quick_options)
        quick_raw = st.pills(
            "납기일 종료 (빠른 선택)",
            options=quick_options,
            default=default_quick,
            key=quick_key,
            selection_mode="single",
            on_change=_on_change_single_select,
            args=(quick_key, default_quick, quick_options),
            label_visibility="collapsed",
        )
        quick = _coerce_single_value(quick_raw, default=default_quick, options=quick_options)
        if quick == "당월":
            default_end = _end_of_month(_today_kst())
        elif quick == "+7일":
            default_end = _today_kst() + timedelta(days=7)
        elif quick == "+14일":
            default_end = _today_kst() + timedelta(days=14)
        else:
            default_end = _today_kst()

        # Ensure quick pick actually updates the date_input (Streamlit keeps widget state by key).
        prev_quick = st.session_state.get(prev_quick_key)
        if prev_quick != quick:
            st.session_state[end_key] = default_end
            st.session_state[prev_quick_key] = quick

        end_date = st.date_input(
            "납기일 종료",
            value=st.session_state.get(end_key, default_end),
            key=end_key,
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
        if view_mode == "사출 계획":
            inj_quick_state = st.session_state.get("inj_due_quick", "해제")
            if inj_quick_state != "해제":
                codes_src = _apply_due_date_end_filter(codes_src, st.session_state.get("inj_due_end", _today_kst()))
            value_col = "사출" if "사출" in codes_src.columns else "누수규격"
        else:
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

    if view_mode == "리스크":
        st.subheader("리스크 (수주 기준)")

        if not prod_daily_csv:
            st.error("`생산실적` 시트 기반 CAPA 데이터가 없습니다. 엑셀에 `생산실적` 시트가 있는지 확인하세요.")
            st.stop()

        st.caption(
            f"기준: CAPA=최근 {RISK_CAPA_RUN_DAYS} 가동일(전일까지, 양품>0) 평균 · "
            f"YELLOW 버퍼={RISK_YELLOW_BUFFER_DAYS:.0f}일 · 24/7 연속운영 가정"
        )
        st.caption(
            "등급: RED=납기내 불가(필요일수>남은일수) · YELLOW=여유부족(버퍼 1일 이내) · GREEN=가능"
        )

        # 완료예정일(스케줄)은 전체 backlog 기준으로 계산하고, 화면 표시만 필터 적용한다.
        schedule_src = order_df_all if isinstance(order_df_all, pd.DataFrame) and not order_df_all.empty else order_df
        view_src = order_df

        def _to_order_level(df0: pd.DataFrame) -> pd.DataFrame:
            stage_cols = [c for c in [*DEFAULT_STAGE_COLS, "필요수량"] if c in df0.columns]
            if not stage_cols:
                return df0
            work = df0.copy()
            for c in stage_cols:
                work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0)
            if "납기일" in work.columns:
                work["납기일"] = pd.to_datetime(work["납기일"], errors="coerce")
            group_key = [c for c in ["이니셜", "수주번호", "신규분류 요약코드"] if c in work.columns]
            if not group_key:
                group_key = ["수주번호"] if "수주번호" in work.columns else []
            if not group_key:
                return work
            agg: dict[str, str] = {c: "sum" for c in stage_cols}
            if "납기일" in work.columns:
                agg["납기일"] = "min"
            if "품명" in work.columns:
                agg["품명"] = "first"
            return work.groupby(group_key, dropna=False, as_index=False).agg(agg)

        schedule_orders = _to_order_level(schedule_src)
        view_orders = _to_order_level(view_src)

        prod_daily_df = _load_prod_daily_csv(str(prod_daily_csv), os.path.getmtime(str(prod_daily_csv)))
        as_of = _today_kst() - timedelta(days=1)
        capa_table = _compute_capa_table_from_prod_daily(prod_daily_df, n_run_days=int(RISK_CAPA_RUN_DAYS), as_of=as_of)

        # Compute risk/schedule on ALL orders (schedule_orders), then filter down for display (subset).
        risk_all = _build_order_risk_table(
            schedule_orders,
            capa_table,
            today=_today_kst(),
            buffer_days=float(RISK_YELLOW_BUFFER_DAYS),
            start_offset_days=int(RISK_SCHED_START_OFFSET_DAYS),
        )
        if risk_all.empty:
            st.caption("표시할 리스크 대상이 없습니다.")
            st.stop()

        # Base filters (code + due) for counts/pills.
        risk_base = risk_all.copy()
        if code != "전체" and new_code_col in risk_base.columns:
            risk_base = risk_base.loc[risk_base[new_code_col].astype("string") == code].copy()
        if quick != "해제" and "납기일" in risk_base.columns:
            risk_base = _apply_due_date_end_filter(risk_base, st.session_state.get("risk_due_end", _today_kst()))

        grade_options = ["RED", "YELLOW", "GREEN"]
        counts = risk_base["리스크등급"].value_counts().to_dict() if "리스크등급" in risk_base.columns else {}
        filter_options = grade_options

        def _grade_label(opt: str) -> str:
            s = str(opt)
            if s in grade_options:
                return f"{s} ({int(counts.get(s, 0))})"
            return s

        grade_raw = st.pills(
            "등급 필터",
            options=filter_options,
            default=st.session_state.get("risk_grade_pill", ["RED", "YELLOW"]),
            key="risk_grade_pill",
            format_func=_grade_label,
            selection_mode="multi",
            on_change=_on_change_risk_grade_pills,
            kwargs={"key": "risk_grade_pill", "grade_options": grade_options},
            label_visibility="collapsed",
        )
        selected_grades = grade_raw if isinstance(grade_raw, list) else ([grade_raw] if grade_raw else [])
        selected_grades = [g for g in selected_grades if g in grade_options]
        if not selected_grades:
            selected_grades = ["RED", "YELLOW"]

        search_raw = st.text_input(
            "검색 (이니셜/수주번호/품명)",
            placeholder="예: 해외, 202601, O2O2",
            key=f"risk_{code}_search",
        )

        risk_df = _filter_by_any_contains(risk_base, ["품명", "이니셜", "수주번호"], search_raw)
        risk_df = risk_df.loc[risk_df["리스크등급"].isin(selected_grades)].copy()
        if risk_df.empty:
            st.caption("필터 조건에 해당하는 항목이 없습니다.")
            st.stop()
        st.caption(f"표시 건수: {len(risk_df):,}")

        show_cols = [
            "우선순위",
            "이니셜",
            "수주번호",
            "신규분류 요약코드",
            "품명",
            "납기일",
            "리스크등급",
            "리스크사유",
            "누수규격",
            "완료예정일",
        ]
        show_cols = [c for c in show_cols if c in risk_df.columns]

        xlsx_bytes_risk = _to_excel_bytes(risk_df[show_cols], sheet_name="리스크")
        st.download_button(
            "엑셀 다운로드",
            data=xlsx_bytes_risk,
            file_name=f"리스크_{code}_{_today_kst().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"risk_{code}_download",
        )

        column_config = {
            "우선순위": st.column_config.NumberColumn(format="%d", width="small"),
            "이니셜": st.column_config.TextColumn(width="small"),
            "수주번호": st.column_config.TextColumn(width="medium"),
            "신규분류 요약코드": st.column_config.TextColumn(width="medium"),
            "품명": st.column_config.TextColumn(width="large"),
            "납기일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
            "리스크등급": st.column_config.TextColumn(width="small"),
            "리스크사유": st.column_config.TextColumn(width="large"),
            "누수규격": st.column_config.NumberColumn(format="localized", width="small"),
            "완료예정일": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
        }
        column_config = {k: v for k, v in column_config.items() if k in show_cols}

        table_h = _table_height_for_rows(len(risk_df), min_height=320, max_height=780)
        st.dataframe(
            _style_dataframe_like_dashboard(risk_df[show_cols]),
            use_container_width=True,
            height=table_h,
            hide_index=True,
            column_config=column_config,
        )
        return

    if view_mode == "사출 계획":
        st.subheader("사출 4~5일 단기 스케줄 (자동 생성)")
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            horizon_days = int(st.selectbox("기간(일)", options=[4, 5], index=1, key="inj_horizon_days"))
        with col2:
            start_date = st.date_input("시작일", value=_today_kst(), key="inj_start_date")
        with col3:
            st.caption("설비 1대/일=2블록, 블록당 POWER 최대 8칸(중복 허용). E 공란+F 기입 설비는 배정하지 않습니다.")

        excel_mtime = float(os.path.getmtime(excel_path))
        base_df = df
        inj_quick_state = st.session_state.get("inj_due_quick", "해제")
        if inj_quick_state != "해제":
            base_df = _apply_due_date_end_filter(base_df, st.session_state.get("inj_due_end", _today_kst()))

        subset = base_df if code == "전체" else base_df[base_df[new_code_col].astype("string") == code].copy()
        subset2 = subset.copy()
        subset2 = _attach_item_codes(subset2, detail_for_map, allowed_prefixes=["R"])
        if "제품코드" not in subset2.columns:
            subset2["제품코드"] = ""

        inj_info = _load_injection_sheet_cached(excel_path, excel_mtime)
        inj_equip = inj_info.get("equip", pd.DataFrame())
        inj_arrange = inj_info.get("arrange", pd.DataFrame())
        if inj_equip is None or inj_equip.empty:
            st.error("엑셀에 `사출` 시트(설비 현황)가 없습니다.")
            st.stop()

        sched, remaining, warns = _build_injection_schedule(
            demand=subset2,
            inj_equip=inj_equip,
            arrange=inj_arrange,
            excel_path=excel_path,
            excel_mtime=excel_mtime,
            start_date=start_date,
            horizon_days=horizon_days,
        )

        if warns:
            for w in warns[:8]:
                st.warning(w)

        if sched is None or sched.empty:
            st.caption("생성된 스케줄이 없습니다.")
        else:
            xlsx_bytes = _to_excel_bytes(sched, sheet_name="사출스케줄")
            st.download_button(
                "엑셀 다운로드 (사출 스케줄)",
                data=xlsx_bytes,
                file_name=f"사출스케줄_{code}_{_today_kst().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"inj_{code}_download",
            )
            col_cfg = {
                "날짜": st.column_config.DatetimeColumn(format="YYYY-MM-DD", width="small"),
                "설비명": st.column_config.TextColumn(width="small"),
                "제품명코드": st.column_config.TextColumn(width="small"),
                "제품명": st.column_config.TextColumn(width="large"),
                "Block": st.column_config.NumberColumn(format="%d", width="small"),
                "POWER 리스트": st.column_config.TextColumn(width="large"),
                "POWER 개수": st.column_config.NumberColumn(format="%d", width="small"),
                "배정수량": st.column_config.NumberColumn(format="localized", width="small"),
                "잔여수량": st.column_config.NumberColumn(format="localized", width="small"),
                "세팅구분": st.column_config.TextColumn(width="small"),
            }
            show_cols = [c for c in col_cfg.keys() if c in sched.columns]
            table_h = _table_height_for_rows(len(sched), min_height=360, max_height=860)
            st.dataframe(
                _style_dataframe_like_dashboard(sched[show_cols]),
                use_container_width=True,
                height=table_h,
                hide_index=True,
                column_config={k: v for k, v in col_cfg.items() if k in show_cols},
            )

        if remaining is not None and (not remaining.empty):
            st.divider()
            st.subheader("미배정 잔여수량")
            rem_show = remaining.copy()
            if "잔여수량" in rem_show.columns:
                rem_show["잔여수량"] = pd.to_numeric(rem_show["잔여수량"], errors="coerce").fillna(0).astype(int)
            if "납기일" in rem_show.columns:
                rem_show["납기일"] = pd.to_datetime(rem_show["납기일"], errors="coerce")
            sort_cols = [c for c in ["납기일", "잔여수량", "제품명코드"] if c in rem_show.columns]
            if sort_cols:
                asc_map = {"납기일": True, "잔여수량": False, "제품명코드": True}
                rem_show = rem_show.sort_values(
                    sort_cols,
                    ascending=[asc_map.get(c, True) for c in sort_cols],
                    na_position="last",
                )
            xlsx_bytes_rem = _to_excel_bytes(rem_show, sheet_name="미배정")
            st.download_button(
                "엑셀 다운로드 (미배정 잔여)",
                data=xlsx_bytes_rem,
                file_name=f"사출_미배정_{code}_{_today_kst().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"inj_{code}_download_rem",
            )
            st.dataframe(
                _style_dataframe_like_dashboard(rem_show),
                use_container_width=True,
                height=_table_height_for_rows(len(rem_show), min_height=220, max_height=520),
                hide_index=True,
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
