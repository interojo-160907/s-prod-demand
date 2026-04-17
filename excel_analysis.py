import argparse
import glob
import json
import math
import os
import re
from datetime import datetime

import pandas as pd


def _find_default_excel_path() -> str:
    candidates = sorted(glob.glob("*.xlsx"))
    if not candidates:
        raise SystemExit("No .xlsx found in current directory. Pass --file <path>.")
    return candidates[0]


def _safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _coerce_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def _to_datetime(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c in df.columns:
            s = df[c]
            dt = pd.to_datetime(s, errors="coerce")
            # Handle Excel serial dates (e.g., 46076 -> 2026-02-23) which pandas
            # would otherwise interpret as nanoseconds-from-epoch (1970-01-01...).
            s_num = pd.to_numeric(s, errors="coerce")
            mask_excel = s_num.notna() & (s_num >= 20000) & (s_num <= 80000)
            if mask_excel.any():
                dt_excel = pd.to_datetime(s_num, unit="D", origin="1899-12-30", errors="coerce")
                dt = dt.where(~mask_excel, dt_excel)
            df[c] = dt
    return df


def _format_power(power) -> str:
    if pd.isna(power):
        return ""
    # Display with explicit sign and zero-padded integer part (e.g., -02.75, +00.00).
    # Also preserve the sign of negative zero (-0.0 -> -00.00).
    try:
        power_f = float(str(power).strip())
        if math.isnan(power_f):
            return ""
        # Business rule: any zero (0, 0.00, +0.0, -0.0) must display as "-00.00".
        if power_f == 0.0:
            sign = "-"
        else:
            sign = "-" if power_f < 0 else "+"
        mag = abs(power_f)
        mag_s = f"{mag:05.2f}" if mag < 100 else f"{mag:.2f}"
        return f"{sign}{mag_s}"
    except Exception:
        return str(power).strip()


_RE_TORIC = re.compile(r"([+-]\d+\.\d{2})([+-]\d+\.\d{2})(\d{3})$")
_RE_TWO_FLOATS = re.compile(r"([+-]\d+\.\d{2})([+-]\d+\.\d{2})$")


def _format_spec(value: float, *, zero_sign: str = "+") -> str:
    if value is None:
        return ""
    try:
        v = float(value)
    except Exception:
        return ""
    if math.isnan(v):
        return ""
    if v == 0.0:
        sign = zero_sign
    else:
        sign = "-" if v < 0 else "+"
    mag = abs(v)
    # For CP/ADD display: prefer +0.00, -0.75, +2.50 (no 2-digit padding).
    mag_s = f"{mag:.2f}"
    return f"{sign}{mag_s}"


def _parse_lens_spec_from_code(code) -> tuple[str, str, str]:
    """
    Returns (ADD, CP, AXIS) as formatted strings.
    - Multifocal: ...<POWER><ADD> (e.g., R1025-02.50+1.50) -> ADD populated.
    - Toric: ...<POWER><CP><AXIS> (e.g., R1052-01.50-0.75090) -> CP/AXIS populated.
    """
    if code is None or (isinstance(code, float) and math.isnan(code)):
        return ("", "", "")
    s = str(code).strip()
    if not s:
        return ("", "", "")

    m = _RE_TORIC.search(s)
    if m:
        cp = _format_spec(float(m.group(2)), zero_sign="+")
        axis = m.group(3).zfill(3)
        return ("", cp, axis)

    m = _RE_TWO_FLOATS.search(s)
    if m:
        add = _format_spec(float(m.group(2)), zero_sign="+")
        return (add, "", "")

    return ("", "", "")


def _build_product_family_map(xl: pd.ExcelFile) -> dict[str, str]:
    if "설비별" not in xl.sheet_names:
        return {}

    equip = _clean_columns(xl.parse("설비별"))

    name_col = None
    for c in ["제품 이름", "제품명"]:
        if c in equip.columns:
            name_col = c
            break
    if name_col is None or "POWER" not in equip.columns:
        return {}

    key_cols: list[str] = []
    for c in ["제품 그룹 코드", "제품그룹코드"]:
        if c in equip.columns:
            key_cols.append(c)
    for c in ["제품코드(Full)", "제품코드"]:
        if c in equip.columns:
            key_cols.append(c)
    if not key_cols:
        return {}

    equip["POWER_fmt"] = equip["POWER"].map(_format_power)
    equip["제품군"] = equip[name_col].astype(str).str.strip() + " + " + equip["POWER_fmt"]

    m: dict[str, str] = {}
    for code_col in key_cols:
        codes = equip[code_col].astype(str).str.strip()
        fams = equip["제품군"].astype(str)
        for code, fam in zip(codes, fams, strict=False):
            if not code or code.lower() == "nan":
                continue
            # Prefer first-seen mapping to keep deterministic behavior.
            if code not in m:
                m[code] = fam
    return m


def _drop_total_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    total_tokens = {"총합계", "종합계"}
    for col in ["설비 사이트 코드", "수주번호", "제품 코드", "제품코드(Full)", "제품코드"]:
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            df = df[~s.isin(total_tokens)]
    return df


def export_due_shortage_with_wip(file_path: str, out_dir: str) -> dict:
    _safe_mkdir(out_dir)
    xl = pd.ExcelFile(file_path)

    # Demand (납기 기준 필요수량): from 이니셜별 (수주/납기/제품코드 기반)
    if "이니셜별" not in xl.sheet_names:
        return {"enabled": False, "reason": "Missing sheet: 이니셜별"}

    demand = _clean_columns(xl.parse("이니셜별"))
    demand = _drop_total_rows(demand)
    demand = _to_datetime(demand, ["납기일"])
    if "생산 수량" not in demand.columns:
        return {"enabled": False, "reason": "Missing column: 생산 수량"}
    if "제품 코드" not in demand.columns:
        return {"enabled": False, "reason": "Missing column: 제품 코드"}

    demand = _coerce_numeric(demand, ["생산 수량"])
    demand["제품 코드"] = demand["제품 코드"].astype(str).str.strip()
    if "신규분류 요약코드" in demand.columns:
        demand["신규분류 요약코드"] = demand["신규분류 요약코드"].astype(str).str.strip()

    product_family_map = _build_product_family_map(xl)
    demand["제품군"] = demand["제품 코드"].map(product_family_map)
    miss = demand.loc[demand["제품군"].isna()].copy() if product_family_map else demand.copy()

    # Fallback: keep stable grouping even if mapping fails.
    if "수요 제품 이름" in demand.columns:
        fallback = (
            demand["수요 제품 이름"].astype(str).str.strip()
            + " | "
            + demand["제품 코드"].astype(str).str.strip()
        )
    else:
        fallback = demand["제품 코드"].astype(str).str.strip()
    demand["제품군"] = demand["제품군"].fillna(fallback)

    # Mapping misses for review.
    miss_out = os.path.join(out_dir, "제품코드_매핑누락.csv")
    miss_key_col = "제품 코드"
    (
        miss.groupby([miss_key_col], as_index=False)
        .agg(건수=(miss_key_col, "size"), 예시_제품군=("제품군", "first"))
        .sort_values("건수", ascending=False)
        .head(5000)
        .to_csv(miss_out, index=False, encoding="utf-8-sig")
    )
    if "_map_key" in demand.columns:
        demand = demand.drop(columns=["_map_key"])

    demand_g = (
        demand.groupby(["제품군", "납기일"], dropna=False, as_index=False)[["생산 수량"]]
        .sum(numeric_only=True)
        .rename(columns={"생산 수량": "필요수량"})
    )

    # Supply/WIP snapshot: from 재고 (제품명 + 파워 기준)
    if "재고" not in xl.sheet_names:
        return {"enabled": False, "reason": "Missing sheet: 재고"}

    inv = _clean_columns(xl.parse("재고"))
    inv = _drop_total_rows(inv)
    required_inv_cols = [
        "제품명",
        "파워",
        "사출재고",
        "분리재고",
        "검사접착재고",
        "누수규격검사재고",
        "생산완제품재고",
        "포장완제품재고",
    ]
    for col in required_inv_cols:
        if col not in inv.columns:
            return {"enabled": False, "reason": f"Missing column in 재고: {col}"}

    inv["파워_fmt"] = inv["파워"].map(_format_power)
    inv["제품군"] = inv["제품명"].astype(str).str.strip() + " + " + inv["파워_fmt"]
    inv = _coerce_numeric(
        inv,
        [
            "사출재고",
            "분리재고",
            "검사접착재고",
            "누수규격검사재고",
            "생산완제품재고",
            "포장완제품재고",
        ],
    )

    inv_g = (
        inv.groupby(["제품군"], dropna=False, as_index=False)[
            [
                "사출재고",
                "분리재고",
                "검사접착재고",
                "누수규격검사재고",
                "생산완제품재고",
                "포장완제품재고",
            ]
        ]
        .sum(numeric_only=True)
        .fillna(0)
    )

    # Derive "공정별 완료수량" as cumulative passed counts (downstream sums).
    inv_g["완료_검사"] = inv_g["생산완제품재고"] + inv_g["포장완제품재고"]
    inv_g["완료_접착"] = inv_g["누수규격검사재고"] + inv_g["완료_검사"]
    inv_g["완료_하이드"] = inv_g["검사접착재고"] + inv_g["완료_접착"]
    inv_g["완료_분리"] = inv_g["분리재고"] + inv_g["완료_하이드"]
    inv_g["완료_사출"] = inv_g["사출재고"] + inv_g["완료_분리"]

    # WIP by transition (matches the example differences).
    inv_g["재공_분리"] = (inv_g["완료_사출"] - inv_g["완료_분리"]).clip(lower=0)  # ~= 사출재고
    inv_g["재공_하이드"] = (inv_g["완료_분리"] - inv_g["완료_하이드"]).clip(lower=0)  # ~= 분리재고
    inv_g["재공_접착"] = (inv_g["완료_하이드"] - inv_g["완료_접착"]).clip(lower=0)  # ~= 검사접착재고
    inv_g["재공_검사"] = (inv_g["완료_접착"] - inv_g["완료_검사"]).clip(lower=0)  # ~= 누수규격검사재고
    inv_g["재공_합계"] = (inv_g["재공_분리"] + inv_g["재공_하이드"] + inv_g["재공_접착"] + inv_g["재공_검사"]).clip(lower=0)
    inv_g["가용총량"] = inv_g["완료_검사"] + inv_g["재공_합계"]

    merged = demand_g.merge(inv_g, on="제품군", how="left")
    # Track demand families that do not match inventory families (naming mismatch).
    supply_miss = merged.loc[merged["가용총량"].isna(), ["제품군"]].drop_duplicates()
    supply_miss_out = os.path.join(out_dir, "제품군_재고매칭누락.csv")
    supply_miss.to_csv(supply_miss_out, index=False, encoding="utf-8-sig")

    # Fill missing inventory with 0 (means no snapshot supply found for that 제품군).
    fill_cols = [
        "사출재고",
        "분리재고",
        "검사접착재고",
        "누수규격검사재고",
        "생산완제품재고",
        "포장완제품재고",
        "완료_사출",
        "완료_분리",
        "완료_하이드",
        "완료_접착",
        "완료_검사",
        "재공_분리",
        "재공_하이드",
        "재공_접착",
        "재공_검사",
        "재공_합계",
        "가용총량",
    ]
    for c in fill_cols:
        if c in merged.columns:
            merged[c] = merged[c].fillna(0)

    # Allocation by due date (earliest first) per 제품군.
    merged = merged.sort_values(["제품군", "납기일"], ascending=[True, True])

    def _allocate(group: pd.DataFrame) -> pd.DataFrame:
        group = group.copy()
        avail_total = float(group["가용총량"].iloc[0] if len(group) else 0)
        avail_finished = float(group["완료_검사"].iloc[0] if len(group) else 0)
        avail_wip = max(0.0, avail_total - avail_finished)

        alloc_finished_list = []
        alloc_wip_list = []
        shortage_list = []
        remain_total_list = []

        for need in group["필요수량"].astype(float).fillna(0).tolist():
            use_finished = min(avail_finished, need)
            avail_finished -= use_finished
            remaining_need = need - use_finished

            use_wip = min(avail_wip, remaining_need)
            avail_wip -= use_wip
            remaining_need2 = remaining_need - use_wip

            shortage = max(0.0, remaining_need2)
            alloc_finished_list.append(use_finished)
            alloc_wip_list.append(use_wip)
            shortage_list.append(shortage)
            remain_total_list.append(avail_finished + avail_wip)

        group["할당_검사완료"] = alloc_finished_list
        group["할당_WIP"] = alloc_wip_list
        group["부족수량"] = shortage_list
        group["잔여가용"] = remain_total_list
        return group

    # Avoid groupby.apply deprecation warnings by allocating per group explicitly.
    allocated_parts: list[pd.DataFrame] = []
    for _, group in merged.groupby("제품군", sort=False):
        allocated_parts.append(_allocate(group))
    merged = pd.concat(allocated_parts, ignore_index=True) if allocated_parts else merged

    # Bottleneck from WIP transition stocks (largest queue).
    wip_cols = ["재공_분리", "재공_하이드", "재공_접착", "재공_검사"]
    bottleneck_stage_for_col = {
        "재공_분리": "분리",
        "재공_하이드": "하이드",
        "재공_접착": "접착",
        "재공_검사": "검사",
    }

    def _bottleneck(row) -> str:
        max_col = max(wip_cols, key=lambda c: float(row.get(c, 0) or 0))
        max_val = float(row.get(max_col, 0) or 0)
        if max_val > 0:
            return bottleneck_stage_for_col.get(max_col, "")
        if float(row.get("부족수량", 0) or 0) > 0:
            return "사출(미투입)"
        return ""

    merged["병목공정"] = merged.apply(_bottleneck, axis=1)

    # Remaining processes (for scanability) + risk flag.
    wip_to_step = {
        "재공_분리": "분리",
        "재공_하이드": "하이드",
        "재공_접착": "접착",
        "재공_검사": "검사",
    }

    def _remaining_steps(row) -> str:
        steps: list[str] = []
        if float(row.get("부족수량", 0) or 0) > 0:
            steps.append("사출")
        # If we must rely on WIP for this due date, show which queues exist for that 제품군.
        if float(row.get("할당_WIP", 0) or 0) > 0:
            for col, step in wip_to_step.items():
                if float(row.get(col, 0) or 0) > 0:
                    steps.append(step)
        # De-dup while preserving order.
        seen = set()
        uniq = []
        for s in steps:
            if s not in seen:
                uniq.append(s)
                seen.add(s)
        return " > ".join(uniq)

    merged["남은공정"] = merged.apply(_remaining_steps, axis=1)
    merged["납기리스크"] = merged.apply(
        lambda r: "Y"
        if (float(r.get("부족수량", 0) or 0) > 0 and str(r.get("남은공정", "")).strip() != "")
        else "",
        axis=1,
    )

    # Add D-day (Seoul local date).
    today = pd.Timestamp.now(tz="Asia/Seoul").normalize().tz_localize(None)
    merged["D_day"] = (pd.to_datetime(merged["납기일"], errors="coerce") - today).dt.days

    out_path = os.path.join(out_dir, "납기_제품군_부족분석.csv")
    cols = [
        "제품군",
        "납기일",
        "D_day",
        "필요수량",
        "완료_사출",
        "완료_분리",
        "완료_하이드",
        "완료_접착",
        "완료_검사",
        "재공_분리",
        "재공_하이드",
        "재공_접착",
        "재공_검사",
        "재공_합계",
        "가용총량",
        "할당_검사완료",
        "할당_WIP",
        "부족수량",
        "잔여가용",
        "병목공정",
        "남은공정",
        "납기리스크",
    ]
    merged[cols].to_csv(out_path, index=False, encoding="utf-8-sig")

    return {
        "enabled": True,
        "rows": int(merged.shape[0]),
        "outputs": [out_path, miss_out, supply_miss_out],
        "mapping_size": int(len(product_family_map)),
        "mapping_miss_unique_codes": int(miss["제품 코드"].nunique() if "제품 코드" in miss.columns else 0),
        "supply_miss_families": int(supply_miss.shape[0]),
    }


def export_due_process_shortage(file_path: str, out_dir: str) -> dict:
    """
    Export 납기 기준 제품군별 공정별 부족(=이니셜별 시트의 공정 컬럼 합계) + 필요수량.

    This matches the user's expectation that 공정 컬럼([10]/[20]/[45]/[55]/[80]) are "부족/필요" values
    rather than computed from inventory snapshots.
    """
    _safe_mkdir(out_dir)
    xl = pd.ExcelFile(file_path)

    if "이니셜별" not in xl.sheet_names:
        return {"enabled": False, "reason": "Missing sheet: 이니셜별"}

    demand = _clean_columns(xl.parse("이니셜별"))
    demand = _drop_total_rows(demand)
    demand = _to_datetime(demand, ["납기일"])

    required_cols = ["제품 코드", "납기일", "생산 수량"]
    for c in required_cols:
        if c not in demand.columns:
            return {"enabled": False, "reason": f"Missing column: {c}"}

    process_cols = [
        "[10]사출조립",
        "[20]분리",
        "[45]하이드레이션/전면검사",
        "[55]접착/멸균",
        "[80]누수/규격검사",
    ]
    present_process_cols = [c for c in process_cols if c in demand.columns]

    demand = _coerce_numeric(demand, ["생산 수량"] + present_process_cols)
    demand["제품 코드"] = demand["제품 코드"].astype(str).str.strip()
    if "신규분류 요약코드" in demand.columns:
        demand["신규분류 요약코드"] = demand["신규분류 요약코드"].astype(str).str.strip()

    add_cp_axis = demand["제품 코드"].map(_parse_lens_spec_from_code)
    demand["ADD"], demand["CP"], demand["AXIS"] = zip(*add_cp_axis, strict=False)

    product_family_map = _build_product_family_map(xl)
    demand["제품군"] = demand["제품 코드"].map(product_family_map)
    miss = demand.loc[demand["제품군"].isna()].copy() if product_family_map else demand.copy()
    if "수요 제품 이름" in demand.columns:
        fallback = (
            demand["수요 제품 이름"].astype(str).str.strip()
            + " | "
            + demand["제품 코드"].astype(str).str.strip()
        )
    else:
        fallback = demand["제품 코드"].astype(str).str.strip()
    demand["제품군"] = demand["제품군"].fillna(fallback)
    miss_out = os.path.join(out_dir, "제품코드_매핑누락.csv")
    miss_key_col = "제품 코드"
    (
        miss.groupby([miss_key_col], as_index=False)
        .agg(건수=(miss_key_col, "size"), 예시_제품군=("제품군", "first"))
        .sort_values("건수", ascending=False)
        .head(5000)
        .to_csv(miss_out, index=False, encoding="utf-8-sig")
    )

    agg = ["생산 수량"] + present_process_cols
    group_cols = []
    if "신규분류 요약코드" in demand.columns:
        group_cols.append("신규분류 요약코드")
    group_cols += ["제품군", "ADD", "CP", "AXIS", "납기일"]
    g = demand.groupby(group_cols, dropna=False, as_index=False)[agg].sum(numeric_only=True)
    g = g.rename(
        columns={
            "생산 수량": "필요수량",
            "[10]사출조립": "사출",
            "[20]분리": "분리",
            "[45]하이드레이션/전면검사": "하이드레이션",
            "[55]접착/멸균": "접착",
            "[80]누수/규격검사": "누수규격",
        }
    )

    out_path = os.path.join(out_dir, "납기_제품군_공정별부족.csv")
    cols = []
    if "신규분류 요약코드" in g.columns:
        cols.append("신규분류 요약코드")
    cols += ["제품군", "ADD", "CP", "AXIS", "납기일", "필요수량", "사출", "분리", "하이드레이션", "접착", "누수규격"]
    cols = [c for c in cols if c in g.columns]
    sort_cols = ["납기일"]
    if "신규분류 요약코드" in g.columns:
        sort_cols.append("신규분류 요약코드")
    sort_cols.append("제품군")
    g[cols].sort_values(sort_cols, ascending=[True] * len(sort_cols)).to_csv(
        out_path, index=False, encoding="utf-8-sig"
    )

    # Order-level detail export (for per-order priority view).
    detail = demand.copy()
    detail = detail.rename(
        columns={
            "생산 수량": "필요수량",
            "[10]사출조립": "사출",
            "[20]분리": "분리",
            "[45]하이드레이션/전면검사": "하이드레이션",
            "[55]접착/멸균": "접착",
            "[80]누수/규격검사": "누수규격",
        }
    )
    detail_cols = [
        "이니셜",
        "수주번호",
        "신규분류 요약코드",
        "수요 제품 이름",
        "제품군",
        "제품 코드",
        "ADD",
        "CP",
        "AXIS",
        "납기일",
        "사출",
        "분리",
        "하이드레이션",
        "접착",
        "누수규격",
        "필요수량",
    ]
    detail_cols = [c for c in detail_cols if c in detail.columns]
    detail_path = os.path.join(out_dir, "이니셜별_수주상세.csv")
    detail[detail_cols].sort_values(["납기일"], ascending=[True]).to_csv(
        detail_path, index=False, encoding="utf-8-sig"
    )

    if "_map_key" in demand.columns:
        demand = demand.drop(columns=["_map_key"])

    return {
        "enabled": True,
        "rows": int(g.shape[0]),
        "outputs": [out_path, detail_path, miss_out],
        "mapping_size": int(len(product_family_map)),
        "mapping_miss_unique_codes": int(miss[miss_key_col].nunique() if miss_key_col in miss.columns else 0),
    }


def analyze(file_path: str, out_dir: str) -> dict:
    _safe_mkdir(out_dir)

    xl = pd.ExcelFile(file_path)
    sheets = xl.sheet_names

    report: dict = {
        "file": file_path,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "sheets": {},
    }

    # Sheet: 재고
    if "재고" in sheets:
        inv = _clean_columns(xl.parse("재고"))
        inv = _coerce_numeric(
            inv,
            [
                "오더수량",
                "포장단위",
                "포장완제품재고대비부족수량",
                "사출재고",
                "분리재고",
                "검사접착재고",
                "누수규격검사재고",
                "생산완제품재고",
                "포장완제품재고",
            ],
        )

        shortage_col = "포장완제품재고대비부족수량"
        inv_short = inv.loc[inv[shortage_col].fillna(0) > 0].copy()

        top_short = (
            inv_short.groupby(["이니셜", "제품명", "파워"], dropna=False, as_index=False)[
                [shortage_col, "오더수량", "포장완제품재고"]
            ]
            .sum(numeric_only=True)
            .sort_values(shortage_col, ascending=False)
        )
        top_short_path = os.path.join(out_dir, "재고_부족_top200.csv")
        top_short.head(200).to_csv(top_short_path, index=False, encoding="utf-8-sig")

        report["sheets"]["재고"] = {
            "rows": int(inv.shape[0]),
            "cols": int(inv.shape[1]),
            "shortage_positive_rows": int(inv_short.shape[0]),
            "shortage_sum": float(inv_short[shortage_col].sum(skipna=True)),
            "outputs": [top_short_path],
        }

    # Sheet: 설비별
    if "설비별" in sheets:
        equip = _clean_columns(xl.parse("설비별"))
        equip = _coerce_numeric(equip, ["계획 수량"])
        equip = _to_datetime(equip, ["최소 납기일", "최소 목표일"])

        equip_by_process = (
            equip.groupby(["설비 사이트 코드", "공정 코드"], dropna=False, as_index=False)[
                ["계획 수량"]
            ]
            .sum(numeric_only=True)
            .sort_values("계획 수량", ascending=False)
        )
        equip_proc_path = os.path.join(out_dir, "설비별_공정_계획수량.csv")
        equip_by_process.to_csv(equip_proc_path, index=False, encoding="utf-8-sig")

        report["sheets"]["설비별"] = {
            "rows": int(equip.shape[0]),
            "cols": int(equip.shape[1]),
            "outputs": [equip_proc_path],
        }

    # Sheet: 이니셜별
    if "이니셜별" in sheets:
        initial = _clean_columns(xl.parse("이니셜별"))
        # Process columns look like "[10]사출조립", etc.
        process_cols = [c for c in initial.columns if str(c).startswith("[") and "]" in str(c)]
        initial = _coerce_numeric(initial, process_cols + ["생산 수량"])
        initial = _to_datetime(initial, ["납기일"])

        # Aggregate totals per process column (interpreted as "remaining/need" style quantities).
        proc_totals = (
            initial[process_cols]
            .sum(numeric_only=True, skipna=True)
            .sort_values(ascending=False)
            .rename_axis("공정")
            .reset_index(name="합계")
        )
        proc_totals_path = os.path.join(out_dir, "이니셜별_공정별_합계.csv")
        proc_totals.to_csv(proc_totals_path, index=False, encoding="utf-8-sig")

        report["sheets"]["이니셜별"] = {
            "rows": int(initial.shape[0]),
            "cols": int(initial.shape[1]),
            "process_cols": process_cols,
            "outputs": [proc_totals_path],
        }

    report_path = os.path.join(out_dir, "analysis_summary.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    report["outputs"] = [report_path]
    return report


def validate_workbook(file_path: str, template_path: str | None = None) -> dict:
    """
    Validate that required sheets/columns exist.

    If template_path is provided and exists, validate each matching sheet's columns
    against the template (after column normalization via _clean_columns).
    """
    def _sheet_columns(path: str, sheet: str) -> list[str]:
        try:
            df0 = pd.read_excel(path, sheet_name=sheet, nrows=0)
        except Exception:
            return []
        df0 = _clean_columns(df0)
        return [str(c).strip() for c in df0.columns.tolist()]

    result: dict = {"ok": True, "file": file_path, "errors": []}

    try:
        xl = pd.ExcelFile(file_path)
        sheets = set(xl.sheet_names)
    except Exception as e:
        return {"ok": False, "file": file_path, "errors": [f"엑셀 열기 실패: {e}"]}

    # Minimal requirements for this dashboard.
    required_sheets = ["이니셜별"]
    for s in required_sheets:
        if s not in sheets:
            result["ok"] = False
            result["errors"].append(f"필수 시트 누락: {s}")

    if not result["ok"]:
        return result

    # Required columns (normalized names) for key logic.
    required_cols_by_sheet: dict[str, list[str]] = {
        "이니셜별": [
            "제품 코드",
            "납기일",
            "생산 수량",
            "이니셜",
            "수주번호",
            "수요 제품 이름",
            "신규분류 요약코드",
        ],
    }

    for sheet, req_cols in required_cols_by_sheet.items():
        if sheet not in sheets:
            continue
        cols = set(_sheet_columns(file_path, sheet))
        missing = [c for c in req_cols if c not in cols]
        if missing:
            result["ok"] = False
            result["errors"].append(f"{sheet} 시트 컬럼 누락: {', '.join(missing)}")

    # Template-based validation (optional).
    if template_path and os.path.exists(template_path):
        try:
            t_xl = pd.ExcelFile(template_path)
            t_sheets = set(t_xl.sheet_names)
        except Exception as e:
            result["ok"] = False
            result["errors"].append(f"양식 파일 열기 실패: {e}")
            return result

        for sheet in sorted(sheets & t_sheets):
            exp = set(_sheet_columns(template_path, sheet))
            got = set(_sheet_columns(file_path, sheet))
            missing = sorted(exp - got)
            if missing:
                result["ok"] = False
                result["errors"].append(f"[양식대비] {sheet} 컬럼 누락: {', '.join(missing)}")

    return result


def main() -> None:
    ap = argparse.ArgumentParser(description="Analyze S관 부족수량 Excel and export CSV summaries.")
    ap.add_argument("--file", default=None, help="Excel file path (.xlsx). Defaults to first *.xlsx in cwd.")
    ap.add_argument("--out", default="out", help="Output directory (default: out).")
    ap.add_argument(
        "--due-wip",
        action="store_true",
        help="Export 납기 기준 제품군 부족수량 + 재공(WIP) 분석 (out/납기_제품군_부족분석.csv).",
    )
    ap.add_argument(
        "--due-process",
        action="store_true",
        help="Export 납기 기준 제품군 공정별 부족 (out/납기_제품군_공정별부족.csv).",
    )
    args = ap.parse_args()

    file_path = args.file or _find_default_excel_path()
    report = analyze(file_path=file_path, out_dir=args.out)

    due_wip_info = None
    if args.due_wip:
        due_wip_info = export_due_shortage_with_wip(file_path=file_path, out_dir=args.out)
        report["due_wip"] = due_wip_info
        # Update summary JSON to include due_wip results as well.
        report_path = os.path.join(args.out, "analysis_summary.json")
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    due_process_info = None
    if args.due_process:
        due_process_info = export_due_process_shortage(file_path=file_path, out_dir=args.out)
        report["due_process"] = due_process_info
        report_path = os.path.join(args.out, "analysis_summary.json")
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    print("OK")
    print(f"- file: {report['file']}")
    print(f"- out:  {args.out}")
    for sheet_name, info in report.get("sheets", {}).items():
        print(f"- {sheet_name}: rows={info.get('rows')} cols={info.get('cols')} outputs={len(info.get('outputs', []))}")
    if due_wip_info and due_wip_info.get("enabled"):
        print(
            f"- due_wip: rows={due_wip_info.get('rows')} mapping_size={due_wip_info.get('mapping_size')} "
            f"missing_codes={due_wip_info.get('mapping_miss_unique_codes')}"
        )
    if due_process_info and due_process_info.get("enabled"):
        print(
            f"- due_process: rows={due_process_info.get('rows')} mapping_size={due_process_info.get('mapping_size')} "
            f"missing_codes={due_process_info.get('mapping_miss_unique_codes')}"
        )


if __name__ == "__main__":
    main()
