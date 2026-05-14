import os
import re
import shutil
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd

# 📌 경로
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
SHORTAGE_PATH = os.path.join(BASE_PATH, "부족수량")
PROD_PATH = BASE_PATH
OUTPUT_DIR = Path(BASE_PATH) / "outputs"
ARCHIVE_DIR = OUTPUT_DIR / "archive"

# 최근 N개월만 재계산하여 누적본(유효생산량_결과.xlsx)에 교체 반영
REFRESH_MONTHS = 2

# 📌 날짜 추출 함수
DATE_PATTERNS = [
    re.compile(r"(\d{4})[.-]?(\d{2})[.-]?(\d{2})"),  # YYYYMMDD or YYYY-MM-DD
    re.compile(r"(\d{2})(\d{2})(\d{2})")             # YYMMDD
]

def parse_date_from_filename(filename):
    for pattern in DATE_PATTERNS:
        match = pattern.search(filename)
        if match:
            parts = match.groups()
            if len(parts[0]) == 2:
                year = int(parts[0])
                year += 2000 if year < 70 else 1900
                parts = (str(year), parts[1], parts[2])
            return pd.to_datetime("-".join(parts), errors='raise')
    raise ValueError(f"파일명에서 날짜를 찾을 수 없습니다: {filename}")

def _norm_col(s: object) -> str:
    v = "" if s is None else str(s)
    return "".join(v.split()).strip()


def _find_col(df: pd.DataFrame, *, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    norm_to_actual: dict[str, str] = {_norm_col(c): str(c) for c in df.columns}
    for cand in candidates:
        key = _norm_col(cand)
        if key in norm_to_actual:
            return norm_to_actual[key]
    return None


def _find_col_contains(df: pd.DataFrame, *, required_substrings: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    req = [_norm_col(s) for s in required_substrings if _norm_col(s)]
    for c in df.columns:
        nc = _norm_col(c)
        if all(r in nc for r in req):
            return str(c)
    return None

def _norm_process_code(v: object) -> str:
    return "".join(str(v or "").split())


_PROCESS_CODE_TO_NAME: dict[str, str] = {
    "[10]사출조립": "사출",
    "[20]분리": "분리",
    "[45]하이드레이션/전면검사": "하드레이션",
    "[55]접착/멸균": "접착",
    "[80]누수/규격검사": "누수규격",
}


def _find_process_need_col(df: pd.DataFrame, *, code: str, keywords: list[str]) -> str | None:
    # Prefer specific names, then fall back to code-only contains.
    cand = [f"{code}{k}" for k in keywords] + [code]
    found = _find_col(df, candidates=cand)
    if found is not None:
        return found
    for k in keywords:
        f2 = _find_col_contains(df, required_substrings=[code, k])
        if f2 is not None:
            return f2
    return _find_col_contains(df, required_substrings=[code])


# 🔥 부족수량 전체 파일 읽기
if not os.path.isdir(SHORTAGE_PATH):
    os.makedirs(SHORTAGE_PATH, exist_ok=True)
    raise FileNotFoundError(
        f"부족수량 폴더가 없습니다. 폴더를 생성했습니다: {SHORTAGE_PATH}\n"
        "부족수량(필요수량) 원본 엑셀을 해당 폴더에 넣고 다시 실행해주세요."
    )
shortage_files = [f for f in os.listdir(SHORTAGE_PATH) if f.endswith(('.xlsx', '.xls'))]
df_short_list = []
df_short_proc_list = []

for filename in shortage_files:
    file_path = os.path.join(SHORTAGE_PATH, filename)
    df = pd.read_excel(file_path)
    df.columns = df.columns.map(lambda c: str(c).strip())

    col_site = _find_col(df, candidates=["설비 사이트 코드", "설비사이트코드"])
    col_summary = _find_col(df, candidates=["신규분류 요약코드", "신규분류요약코드"])
    col_pname = _find_col(df, candidates=["수요 제품 이름", "수요제품이름", "제품명"])
    col_pcode = _find_col(df, candidates=["제품 코드", "제품코드"])
    col_short = (
        _find_col(df, candidates=["[80]누수/규격검사", "[80] 누수/규격검사"])
        or _find_col_contains(df, required_substrings=["[80]", "누수"])
        or _find_col_contains(df, required_substrings=["[80]", "규격"])
    )

    missing = [n for n, c in [("설비사이트코드", col_site), ("신규분류요약코드", col_summary), ("제품명", col_pname), ("제품코드", col_pcode), ("[80]누수/규격검사", col_short)] if c is None]
    if missing:
        print(f"[WARN] 부족수량 파일 스킵(필수 컬럼 누락: {', '.join(missing)}): {filename}")
        continue

    df['날짜'] = parse_date_from_filename(filename)
    df['설비사이트코드'] = df[col_site].astype(str).str.strip()
    df['신규분류요약코드'] = df[col_summary].astype(str).str.strip()
    df['제품명'] = df[col_pname].astype(str).str.strip()
    df['제품코드'] = df[col_pcode].astype(str).str.strip()
    df['부족수량'] = pd.to_numeric(df[col_short], errors='coerce').fillna(0).astype('int32')

    df = df[~df['제품코드'].str.contains(r"합계|총합|총계", na=False)]
    df = df[df['제품코드'].notna() & (df['제품코드'].str.strip() != "")]

    df_short_list.append(df[['날짜', '설비사이트코드', '신규분류요약코드', '제품명', '제품코드', '부족수량']])

    # 전공정(공정별 필요수량) long format
    # - 공정별 컬럼이 없는 과거 파일은 [80]만 존재할 수 있어, 존재하는 공정만 추가합니다.
    for code, pname in [
        ("[10]", ["사출", "사출조립"]),
        ("[20]", ["분리"]),
        ("[45]", ["하이드", "하이드레이션"]),
        ("[55]", ["접착", "멸균"]),
        ("[80]", ["누수", "규격"]),
    ]:
        col_need = _find_process_need_col(df, code=code, keywords=pname)
        if col_need is None:
            continue
        tmp = df[['날짜', '설비사이트코드', '신규분류요약코드', '제품코드']].copy()
        tmp['공정코드'] = _norm_process_code(code + (pname[0] if pname else "")).split(pname[0])[0]  # just code prefix
        # Normalize to full code label by matching PROCESS map keys later (we keep code prefix here).
        tmp['공정코드'] = code
        tmp['필요수량'] = pd.to_numeric(df[col_need], errors='coerce').fillna(0).astype('int32')
        df_short_proc_list.append(tmp)

if not df_short_list:
    raise FileNotFoundError(f"부족수량 폴더({SHORTAGE_PATH})에 읽을 수 있는 파일이 없습니다. (형식/컬럼 확인 필요)")

# 🔥 하나로 합치기
df_short = pd.concat(df_short_list, ignore_index=True)
del df_short_list

df_short_proc = pd.concat(df_short_proc_list, ignore_index=True) if df_short_proc_list else pd.DataFrame()
del df_short_proc_list

# 📌 설비사이트코드 -> 공장(관) 정규화
_site = df_short["설비사이트코드"].astype("string").str.strip()
_site_nospace = _site.str.replace(r"\s+", "", regex=True)
df_short["공장"] = np.select(
    [_site_nospace.str.contains("A관", na=False), _site_nospace.str.contains("C관", na=False), _site_nospace.str.contains("S관", na=False)],
    ["A관(1공장)", "C관(2공장)", "S관(3공장)"],
    default=_site,
)

if len(df_short_proc) > 0:
    _sitep = df_short_proc["설비사이트코드"].astype("string").str.strip()
    _sitep_nospace = _sitep.str.replace(r"\s+", "", regex=True)
    df_short_proc["공장"] = np.select(
        [_sitep_nospace.str.contains("A관", na=False), _sitep_nospace.str.contains("C관", na=False), _sitep_nospace.str.contains("S관", na=False)],
        ["A관(1공장)", "C관(2공장)", "S관(3공장)"],
        default=_sitep,
    )

# 📌 생산실적 전체 파일 읽기
output_filename = "유효생산량_결과.xlsx"
prod_files = [
    f for f in os.listdir(PROD_PATH)
    if f.endswith(('.xlsx', '.xls'))
    and not f.startswith("~$")
    and f != output_filename
    and not f.startswith("유효생산량_결과")
]
df_prod_list = []

for filename in prod_files:
    file_path = os.path.join(PROD_PATH, filename)
    df = pd.read_excel(file_path, usecols=[1, 2, 3, 4, 6, 7, 8, 16])
    df.columns = df.columns.str.strip()

    df['생산일자'] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
    df['공장'] = df.iloc[:, 1].astype(str).str.strip()
    df['공정코드'] = df.iloc[:, 2].astype(str).str.strip()
    df['신규분류요약'] = df.iloc[:, 3].astype(str).str.strip()
    df['품목코드'] = df.iloc[:, 4].astype(str).str.strip()
    df['품명'] = df.iloc[:, 5].astype(str).str.strip()
    df['양품수량'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0).astype('int32')
    df['상태'] = df.iloc[:, 7].astype(str).str.strip()

    df_prod_list.append(df[['생산일자', '공장', '공정코드', '신규분류요약', '품목코드', '품명', '양품수량', '상태']])

if not df_prod_list:
    raise FileNotFoundError(f"생산실적 폴더({PROD_PATH})에 읽을 파일이 없습니다.")

# 🔥 하나로 합치기
df_prod = pd.concat(df_prod_list, ignore_index=True)
del df_prod_list

# 🔥 공정 및 상태 필터
valid_prod = df_prod[
    (df_prod['공정코드'] == '[80] 누수/규격검사') &
    (df_prod['상태'] == '확인')
]

# 🔥 전공정(공정별) 실적: 공정코드가 확장되어도 여기서 필요한 5개 공정만 사용
df_prod["_공정코드_norm"] = df_prod["공정코드"].map(_norm_process_code)
process_norms = set(_PROCESS_CODE_TO_NAME.keys())
valid_prod_proc = df_prod[
    (df_prod["_공정코드_norm"].isin(process_norms)) &
    (df_prod["상태"].astype(str).str.strip() == "확인")
].copy()

# 🔥 집계
prod_factory_item = (
    valid_prod
    .groupby(['생산일자', '공장', '신규분류요약', '품목코드'])['양품수량']
    .sum()
    .reset_index(name='공장양품수량')
)

prod_factory_item_proc = (
    valid_prod_proc
    .groupby(["생산일자", "공장", "_공정코드_norm", "신규분류요약", "품목코드"], dropna=False)["양품수량"]
    .sum()
    .reset_index(name="실적수량")
)

# 🔥 부족수량 집계
# - 관별(공장별): 설비사이트코드(=관) 기반 공장별 집계
# - 전사(일별요약/규격대응)는 공장별 매칭 합산으로 생성합니다.
shortage_agg_factory = (
    df_short
    .groupby(['날짜', '공장', '신규분류요약코드', '제품코드'], dropna=False)['부족수량']
    .sum()
    .reset_index()
)

shortage_agg_factory_proc = pd.DataFrame()
if len(df_short_proc) > 0:
    df_short_proc["_공정코드_norm"] = df_short_proc["공정코드"].map(_norm_process_code)
    # Expand "[10]" -> full keys if possible by matching prefix.
    # Input shortage columns are keyed by code prefix like "[10]" so convert to best matching full key.
    def _expand_code_prefix(x: str) -> str:
        x = _norm_process_code(x)
        if x in _PROCESS_CODE_TO_NAME:
            return x
        for k in _PROCESS_CODE_TO_NAME.keys():
            if k.startswith(x):
                return k
        return x
    df_short_proc["_공정코드_norm"] = df_short_proc["_공정코드_norm"].map(_expand_code_prefix)

    shortage_agg_factory_proc = (
        df_short_proc
        .groupby(["날짜", "공장", "_공정코드_norm", "신규분류요약코드", "제품코드"], dropna=False)["필요수량"]
        .sum()
        .reset_index()
    )

# 🔥 공장/관별 매칭(생산 + 부족) 생성
factory_result = pd.merge(
    prod_factory_item,
    shortage_agg_factory,
    left_on=['생산일자', '공장', '신규분류요약', '품목코드'],
    right_on=['날짜', '공장', '신규분류요약코드', '제품코드'],
    how='outer',
)

factory_result["생산일자"] = factory_result["생산일자"].combine_first(factory_result["날짜"])
factory_result["날짜"] = factory_result["생산일자"]
factory_result["신규분류요약"] = factory_result["신규분류요약"].combine_first(factory_result["신규분류요약코드"])
factory_result["제품코드"] = factory_result["품목코드"].combine_first(factory_result["제품코드"])

factory_result["양품수량"] = pd.to_numeric(factory_result["공장양품수량"], errors="coerce").fillna(0).astype("int32")
factory_result["부족수량"] = pd.to_numeric(factory_result["부족수량"], errors="coerce").fillna(0).astype("int32")

factory_result["유효생산량"] = np.minimum(factory_result["양품수량"], factory_result["부족수량"]).astype("int32")
factory_result["과생산량"] = np.where(
    factory_result["부족수량"] > 0,
    (factory_result["양품수량"] - factory_result["부족수량"]).clip(lower=0),
    0,
).astype("int32")
factory_result["불필요생산량"] = np.where(factory_result["부족수량"] == 0, factory_result["양품수량"], 0).astype("int32")

# ---------------------------
# 전공정 결과(유효생산량_결과2.xlsx): 공정별_일별실적
# ---------------------------
process_daily_perf = pd.DataFrame(columns=["날짜", "공장", "공정", "실적수량", "부족수량", "과생산수량"])
matching_result_proc = pd.DataFrame(
    columns=[
        "날짜",
        "생산일자",
        "공장",
        "공정",
        "신규분류요약",
        "제품코드",
        "실적수량",
        "필요수량",
        "부족수량",
        "유효생산량",
        "과생산량",
        "불필요생산량",
    ]
)
daily_summary_proc = pd.DataFrame(
    columns=[
        "날짜",
        "공정",
        "총실적",
        "총부족수량",
        "유효생산량",
        "과생산량",
        "불필요생산량",
        "유효비율(%)",
        "과생산비율(%)",
        "불필요비율(%)",
    ]
)
factory_summary_proc = pd.DataFrame(
    columns=[
        "생산일자",
        "공장",
        "공정",
        "신규분류요약",
        "총실적",
        "총부족수량",
        "유효생산량",
        "과생산량",
        "불필요생산량",
        "유효비율(%)",
        "과생산비율(%)",
        "불필요비율(%)",
    ]
)
if len(prod_factory_item_proc) > 0 and len(shortage_agg_factory_proc) > 0:
    prod_factory_item_proc = prod_factory_item_proc.copy()
    prod_factory_item_proc["날짜"] = prod_factory_item_proc["생산일자"]
    prod_factory_item_proc["제품코드"] = prod_factory_item_proc["품목코드"].astype(str)
    prod_factory_item_proc["신규분류요약코드"] = prod_factory_item_proc["신규분류요약"].astype(str)

    need = shortage_agg_factory_proc.copy()
    need = need.rename(columns={"필요수량": "필요수량"})

    merged_proc = pd.merge(
        prod_factory_item_proc[["날짜", "공장", "_공정코드_norm", "신규분류요약코드", "제품코드", "실적수량"]],
        need[["날짜", "공장", "_공정코드_norm", "신규분류요약코드", "제품코드", "필요수량"]],
        on=["날짜", "공장", "_공정코드_norm", "신규분류요약코드", "제품코드"],
        how="outer",
    )
    merged_proc["실적수량"] = pd.to_numeric(merged_proc["실적수량"], errors="coerce").fillna(0).astype("int32")
    merged_proc["필요수량"] = pd.to_numeric(merged_proc["필요수량"], errors="coerce").fillna(0).astype("int32")
    merged_proc["부족수량"] = (merged_proc["필요수량"] - merged_proc["실적수량"]).clip(lower=0).astype("int32")
    merged_proc["과생산수량"] = (merged_proc["실적수량"] - merged_proc["필요수량"]).clip(lower=0).astype("int32")
    merged_proc["공정"] = merged_proc["_공정코드_norm"].map(lambda x: _PROCESS_CODE_TO_NAME.get(str(x), str(x)))

    merged_proc["유효생산량"] = np.minimum(merged_proc["실적수량"], merged_proc["필요수량"]).astype("int32")
    merged_proc["과생산량"] = merged_proc["과생산수량"].astype("int32")
    merged_proc["불필요생산량"] = np.where(merged_proc["필요수량"] == 0, merged_proc["실적수량"], 0).astype("int32")
    merged_proc["신규분류요약"] = merged_proc["신규분류요약코드"]
    merged_proc["생산일자"] = merged_proc["날짜"]

    matching_result_proc = merged_proc[
        [
            "날짜",
            "생산일자",
            "공장",
            "공정",
            "신규분류요약",
            "제품코드",
            "실적수량",
            "필요수량",
            "부족수량",
            "유효생산량",
            "과생산량",
            "불필요생산량",
        ]
    ].copy()

    process_daily_perf = (
        merged_proc
        .groupby(["날짜", "공장", "공정"], dropna=False)
        .agg(
            실적수량=("실적수량", "sum"),
            부족수량=("부족수량", "sum"),
            과생산수량=("과생산수량", "sum"),
        )
        .reset_index()
        .sort_values(["날짜", "공장", "공정"])
        .reset_index(drop=True)
    )

    daily_summary_proc = (
        merged_proc
        .groupby(["날짜", "공정"], dropna=False)
        .agg(
            총실적=("실적수량", "sum"),
            총부족수량=("부족수량", "sum"),
            유효생산량=("유효생산량", "sum"),
            과생산량=("과생산량", "sum"),
            불필요생산량=("불필요생산량", "sum"),
        )
        .reset_index()
        .sort_values(["날짜", "공정"])
        .reset_index(drop=True)
    )
    daily_summary_proc["유효비율(%)"] = (daily_summary_proc["유효생산량"] / daily_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    daily_summary_proc["과생산비율(%)"] = (daily_summary_proc["과생산량"] / daily_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    daily_summary_proc["불필요비율(%)"] = (daily_summary_proc["불필요생산량"] / daily_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

    factory_summary_proc = (
        merged_proc
        .groupby(["생산일자", "공장", "공정", "신규분류요약"], dropna=False)
        .agg(
            총실적=("실적수량", "sum"),
            총부족수량=("부족수량", "sum"),
            유효생산량=("유효생산량", "sum"),
            과생산량=("과생산량", "sum"),
            불필요생산량=("불필요생산량", "sum"),
        )
        .reset_index()
        .sort_values(["생산일자", "공장", "공정", "신규분류요약"])
        .reset_index(drop=True)
    )
    factory_summary_proc["유효비율(%)"] = (factory_summary_proc["유효생산량"] / factory_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    factory_summary_proc["과생산비율(%)"] = (factory_summary_proc["과생산량"] / factory_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    factory_summary_proc["불필요비율(%)"] = (factory_summary_proc["불필요생산량"] / factory_summary_proc["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)


def _write_excel2(
    out_path: Path,
    process_daily: pd.DataFrame,
    *,
    matching: pd.DataFrame,
    daily: pd.DataFrame,
    factory: pd.DataFrame,
    archive_mode: str = "copy",
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if archive_mode == "move":
        _archive_if_exists(out_path)
    elif archive_mode == "copy":
        _backup_if_exists(out_path)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        matching.to_excel(writer, sheet_name="매칭결과", index=False)
        daily.to_excel(writer, sheet_name="일별요약", index=False)
        factory.to_excel(writer, sheet_name="공장_신규분류별", index=False)
        process_daily.to_excel(writer, sheet_name="공정별_일별실적", index=False)


def _read_existing_process_daily(base_path: Path) -> pd.DataFrame | None:
    if not base_path.exists():
        return None
    try:
        df = pd.read_excel(base_path, sheet_name="공정별_일별실적")
    except Exception:
        return None
    return df if df is not None and len(df) > 0 else None


def _read_existing_result2_sources(base_path: Path) -> dict[str, pd.DataFrame] | None:
    if not base_path.exists():
        return None
    sheet_names = ["매칭결과", "일별요약", "공장_신규분류별", "공정별_일별실적"]
    try:
        sheets = pd.read_excel(base_path, sheet_name=sheet_names)
    except Exception:
        return None
    out: dict[str, pd.DataFrame] = {}
    for k in sheet_names:
        df = sheets.get(k)
        if df is None or len(df) == 0:
            continue
        df = df.copy()
        if "공정" in df.columns:
            # Backward-compat: migrate old label
            df["공정"] = df["공정"].replace({"최종공정": "누수규격"})
        out[k] = df
    return out or None


# 📌 공장_신규분류별(관별 요약)
factory_summary = factory_result.groupby(["생산일자", "공장", "신규분류요약"], dropna=False).agg(
    총실적=("양품수량", "sum"),
    총부족수량=("부족수량", "sum"),
    유효생산량=("유효생산량", "sum"),
    과생산량=("과생산량", "sum"),
    불필요생산량=("불필요생산량", "sum"),
).reset_index()
factory_summary["유효비율(%)"] = (factory_summary["유효생산량"] / factory_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
factory_summary["과생산비율(%)"] = (factory_summary["과생산량"] / factory_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
factory_summary["불필요비율(%)"] = (factory_summary["불필요생산량"] / factory_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

# 📌 매칭결과(공장 포함)
matching_result_factory = factory_result[
    [
        "날짜",
        "생산일자",
        "공장",
        "신규분류요약",
        "제품코드",
        "양품수량",
        "부족수량",
        "유효생산량",
        "과생산량",
        "불필요생산량",
    ]
].copy()

# 🔥 전사 매칭 결과는 공장별 매칭 합산으로 재구성(일별요약/규격대응과 일관성 유지)
result = (
    matching_result_factory.groupby(["날짜", "신규분류요약", "제품코드"], dropna=False)[
        ["양품수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"]
    ]
    .sum()
    .reset_index()
)

for col in ["양품수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"]:
    result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0).astype("int32")

# 🔥 일별 요약 계산 (관별 매칭 합산 기준)
daily_summary = result.groupby("날짜", dropna=False).agg(
    총실적=("양품수량", "sum"),
    총부족수량=("부족수량", "sum"),
    유효생산량=("유효생산량", "sum"),
    과생산량=("과생산량", "sum"),
    불필요생산량=("불필요생산량", "sum"),
).reset_index()
daily_summary["유효비율(%)"] = (daily_summary["유효생산량"] / daily_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
daily_summary["과생산비율(%)"] = (daily_summary["과생산량"] / daily_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
daily_summary["불필요비율(%)"] = (daily_summary["불필요생산량"] / daily_summary["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

# 📌 일별 규격(SKU) 대응률 계산
sku_daily = (
    result.groupby(["날짜", "제품코드"], dropna=False)
    .agg(양품수량=("양품수량", "sum"), 부족수량=("부족수량", "sum"))
    .reset_index()
)
sku_daily["생산여부"] = sku_daily["양품수량"] > 0
sku_daily["부족여부"] = sku_daily["부족수량"] > 0

produced_sku = sku_daily[sku_daily["생산여부"]].groupby("날짜", dropna=False)["제품코드"].nunique().rename("생산SKU수")
shortage_sku = sku_daily[sku_daily["부족여부"]].groupby("날짜", dropna=False)["제품코드"].nunique().rename("부족SKU수")
responded_sku = (
    sku_daily[sku_daily["생산여부"] & sku_daily["부족여부"]]
    .groupby("날짜", dropna=False)["제품코드"]
    .nunique()
    .rename("부족대응SKU수")
)

spec_coverage_daily = pd.concat([produced_sku, shortage_sku, responded_sku], axis=1, sort=False).fillna(0).reset_index()
spec_coverage_daily["생산SKU수"] = spec_coverage_daily["생산SKU수"].astype("int32")
spec_coverage_daily["부족SKU수"] = spec_coverage_daily["부족SKU수"].astype("int32")
spec_coverage_daily["부족대응SKU수"] = spec_coverage_daily["부족대응SKU수"].astype("int32")
spec_coverage_daily["규격대응률(%)"] = np.where(
    spec_coverage_daily["생산SKU수"] > 0,
    spec_coverage_daily["부족대응SKU수"] / spec_coverage_daily["생산SKU수"] * 100,
    0,
)

# 📌 공장별/일별 규격(SKU) 대응률 계산
# - "생산한 SKU 중 필요가 있었던 SKU 비중" = (부족대응SKU수 / 생산SKU수) * 100
sku_factory_daily = (
    matching_result_factory.groupby(["날짜", "공장", "제품코드"], dropna=False)
    .agg(양품수량=("양품수량", "sum"), 부족수량=("부족수량", "sum"))
    .reset_index()
)
sku_factory_daily["생산여부"] = sku_factory_daily["양품수량"] > 0
sku_factory_daily["부족여부"] = sku_factory_daily["부족수량"] > 0

produced_sku_f = (
    sku_factory_daily[sku_factory_daily["생산여부"]]
    .groupby(["날짜", "공장"], dropna=False)["제품코드"]
    .nunique()
    .rename("생산SKU수")
)
shortage_sku_f = (
    sku_factory_daily[sku_factory_daily["부족여부"]]
    .groupby(["날짜", "공장"], dropna=False)["제품코드"]
    .nunique()
    .rename("부족SKU수")
)
responded_sku_f = (
    sku_factory_daily[sku_factory_daily["생산여부"] & sku_factory_daily["부족여부"]]
    .groupby(["날짜", "공장"], dropna=False)["제품코드"]
    .nunique()
    .rename("부족대응SKU수")
)

spec_coverage_factory_daily = (
    pd.concat([produced_sku_f, shortage_sku_f, responded_sku_f], axis=1, sort=False)
    .fillna(0)
    .reset_index()
)
spec_coverage_factory_daily["생산SKU수"] = spec_coverage_factory_daily["생산SKU수"].astype("int32")
spec_coverage_factory_daily["부족SKU수"] = spec_coverage_factory_daily["부족SKU수"].astype("int32")
spec_coverage_factory_daily["부족대응SKU수"] = spec_coverage_factory_daily["부족대응SKU수"].astype("int32")
spec_coverage_factory_daily["규격대응률(%)"] = np.where(
    spec_coverage_factory_daily["생산SKU수"] > 0,
    spec_coverage_factory_daily["부족대응SKU수"] / spec_coverage_factory_daily["생산SKU수"] * 100,
    0,
)

# 📌 공장별 기간누적 규격(SKU) 대응률(가중)
# - 기간 전체에서 '고유 SKU' 기준으로 계산(일별 합산으로 중복 카운트 방지)
sku_factory_period = (
    matching_result_factory.groupby(["공장", "제품코드"], dropna=False)
    .agg(양품수량=("양품수량", "sum"), 부족수량=("부족수량", "sum"))
    .reset_index()
)
sku_factory_period["생산여부"] = sku_factory_period["양품수량"] > 0
sku_factory_period["부족여부"] = sku_factory_period["부족수량"] > 0

produced_sku_p = (
    sku_factory_period[sku_factory_period["생산여부"]]
    .groupby("공장", dropna=False)["제품코드"]
    .nunique()
    .rename("생산SKU수")
)
shortage_sku_p = (
    sku_factory_period[sku_factory_period["부족여부"]]
    .groupby("공장", dropna=False)["제품코드"]
    .nunique()
    .rename("부족SKU수")
)
responded_sku_p = (
    sku_factory_period[sku_factory_period["생산여부"] & sku_factory_period["부족여부"]]
    .groupby("공장", dropna=False)["제품코드"]
    .nunique()
    .rename("부족대응SKU수")
)

spec_coverage_factory_total = (
    pd.concat([produced_sku_p, shortage_sku_p, responded_sku_p], axis=1, sort=False)
    .fillna(0)
    .reset_index()
)
spec_coverage_factory_total["생산SKU수"] = spec_coverage_factory_total["생산SKU수"].astype("int32")
spec_coverage_factory_total["부족SKU수"] = spec_coverage_factory_total["부족SKU수"].astype("int32")
spec_coverage_factory_total["부족대응SKU수"] = spec_coverage_factory_total["부족대응SKU수"].astype("int32")
spec_coverage_factory_total["규격대응률(%)"] = np.where(
    spec_coverage_factory_total["생산SKU수"] > 0,
    spec_coverage_factory_total["부족대응SKU수"] / spec_coverage_factory_total["생산SKU수"] * 100,
    0,
)

# 📌 간단 진단 로그(극단값/데이터 누락 탐지)
_diag_missing_prod = spec_coverage_factory_total.loc[
    (spec_coverage_factory_total["생산SKU수"] == 0) & (spec_coverage_factory_total["부족SKU수"] > 0),
    ["공장", "생산SKU수", "부족SKU수", "부족대응SKU수"],
]
if len(_diag_missing_prod) > 0:
    print("[경고] 부족은 있는데 생산 SKU가 0인 공장이 있습니다(공장명/코드 매칭 또는 입력 누락 가능):")
    print(_diag_missing_prod.to_string(index=False))

# 📌 저장
def _archive_if_exists(path: Path) -> None:
    if not path.exists():
        return
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = ARCHIVE_DIR / f"{path.stem}__{ts}{path.suffix}"
    shutil.move(str(path), str(archived))

def _backup_if_exists(path: Path) -> None:
    if not path.exists():
        return
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = ARCHIVE_DIR / f"{path.stem}__{ts}{path.suffix}"
    shutil.copyfile(str(path), str(archived))


def _merge_keep_latest(
    old: pd.DataFrame,
    new: pd.DataFrame,
    subset_cols: list[str],
    sort_cols: list[str] | None = None,
) -> pd.DataFrame:
    if old is None or len(old) == 0:
        out = new.copy()
    elif new is None or len(new) == 0:
        out = old.copy()
    else:
        out = pd.concat([old, new], ignore_index=True)

    if subset_cols and all(c in out.columns for c in subset_cols):
        out = out.drop_duplicates(subset=subset_cols, keep="last")

    if sort_cols and all(c in out.columns for c in sort_cols):
        out = out.sort_values(sort_cols, kind="stable")

    return out.reset_index(drop=True)


def _months_to_refresh_from_periods(periods: pd.Series, n_months: int) -> set[pd.Period]:
    periods = pd.Series(periods).dropna()
    if len(periods) == 0:
        return set()
    uniq = sorted(set(periods.tolist()))
    n = max(int(n_months), 1)
    return set(uniq[-n:])


def _filter_by_months(df: pd.DataFrame, date_col: str, months: set[pd.Period]) -> pd.DataFrame:
    if df is None or len(df) == 0 or not months or date_col not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()
    period = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M")
    return df.loc[period.isin(months)].copy()


def _drop_months(df: pd.DataFrame, date_col: str, months: set[pd.Period]) -> pd.DataFrame:
    if df is None or len(df) == 0 or not months or date_col not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()
    period = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M")
    return df.loc[~period.isin(months)].copy()


def _read_existing_result_sources(base_path: Path) -> dict[str, pd.DataFrame] | None:
    sheet_names = ["매칭결과", "일별요약", "공장_신규분류별", "일별_규격대응"]

    candidates: list[Path] = []
    if base_path.exists():
        candidates.append(base_path)

    if not candidates:
        return None

    buckets: dict[str, list[pd.DataFrame]] = {k: [] for k in sheet_names}
    for p in candidates:
        try:
            sheets = pd.read_excel(p, sheet_name=sheet_names)
        except Exception:
            continue
        for k in sheet_names:
            df = sheets.get(k)
            if df is None or len(df) == 0:
                continue
            buckets[k].append(df.copy())

    out: dict[str, pd.DataFrame] = {}
    for k, frames in buckets.items():
        if frames:
            out[k] = pd.concat(frames, ignore_index=True)

    return out or None


def _write_excel(
    out_path: Path,
    res: pd.DataFrame,
    daily: pd.DataFrame,
    factory: pd.DataFrame,
    spec_daily: pd.DataFrame,
    spec_factory_daily: pd.DataFrame,
    spec_factory_total: pd.DataFrame,
    *,
    archive_mode: str = "move",
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if archive_mode == "move":
        _archive_if_exists(out_path)
    elif archive_mode == "copy":
        _backup_if_exists(out_path)
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        res.to_excel(writer, sheet_name='매칭결과', index=False)
        daily.to_excel(writer, sheet_name='일별요약', index=False)
        factory.to_excel(writer, sheet_name='공장_신규분류별', index=False)
        spec_daily.to_excel(writer, sheet_name='일별_규격대응', index=False)
        spec_factory_daily.to_excel(writer, sheet_name="공장별_일별_규격대응", index=False)
        spec_factory_total.to_excel(writer, sheet_name="공장별_기간누적_규격대응", index=False)

# 1) 월별 파일로 분리 저장 (누적 시 파일 비대화 방지)
# NOTE: 누적본 교체 대상 월(months_to_refresh)은 "생산실적(valid_prod)" 기준으로 산정합니다.
#       부족수량만 있는 월까지 포함하면, 생산실적을 안 넣은 월의 기존 누적 데이터가 0으로 덮일 수 있습니다.
result_month_period = pd.to_datetime(matching_result_factory["날짜"], errors="coerce").dt.to_period("M")
daily_month_period = pd.to_datetime(daily_summary["날짜"], errors="coerce").dt.to_period("M")
spec_month_period = pd.to_datetime(spec_coverage_daily["날짜"], errors="coerce").dt.to_period("M")
factory_month_period = pd.to_datetime(factory_summary["생산일자"], errors="coerce").dt.to_period("M")

prod_month_period = pd.to_datetime(valid_prod["생산일자"], errors="coerce").dt.to_period("M")
months_all = pd.Series(prod_month_period).dropna()
months_to_refresh = _months_to_refresh_from_periods(months_all, REFRESH_MONTHS)
months = [m for m in sorted(months_to_refresh)]
for m in months:
    ym = str(m)  # YYYY-MM
    res_m = matching_result_factory.loc[result_month_period == m].copy()
    daily_m = daily_summary.loc[daily_month_period == m].copy()
    spec_m = spec_coverage_daily.loc[spec_month_period == m].copy()
    factory_m = factory_summary.loc[factory_month_period == m].copy()
    spec_f_daily_m = spec_coverage_factory_daily.loc[
        pd.to_datetime(spec_coverage_factory_daily["날짜"], errors="coerce").dt.to_period("M") == m
    ].copy()
    sku_f_period_m = (
        res_m.groupby(["공장", "제품코드"], dropna=False)
        .agg(양품수량=("양품수량", "sum"), 부족수량=("부족수량", "sum"))
        .reset_index()
    )
    sku_f_period_m["생산여부"] = sku_f_period_m["양품수량"] > 0
    sku_f_period_m["부족여부"] = sku_f_period_m["부족수량"] > 0
    produced_m = (
        sku_f_period_m[sku_f_period_m["생산여부"]]
        .groupby("공장", dropna=False)["제품코드"]
        .nunique()
        .rename("생산SKU수")
    )
    shortage_m = (
        sku_f_period_m[sku_f_period_m["부족여부"]]
        .groupby("공장", dropna=False)["제품코드"]
        .nunique()
        .rename("부족SKU수")
    )
    responded_m = (
        sku_f_period_m[sku_f_period_m["생산여부"] & sku_f_period_m["부족여부"]]
        .groupby("공장", dropna=False)["제품코드"]
        .nunique()
        .rename("부족대응SKU수")
    )
    spec_f_total_m = pd.concat([produced_m, shortage_m, responded_m], axis=1, sort=False).fillna(0).reset_index()
    spec_f_total_m["생산SKU수"] = spec_f_total_m["생산SKU수"].astype("int32")
    spec_f_total_m["부족SKU수"] = spec_f_total_m["부족SKU수"].astype("int32")
    spec_f_total_m["부족대응SKU수"] = spec_f_total_m["부족대응SKU수"].astype("int32")
    spec_f_total_m["규격대응률(%)"] = np.where(
        spec_f_total_m["생산SKU수"] > 0,
        spec_f_total_m["부족대응SKU수"] / spec_f_total_m["생산SKU수"] * 100,
        0,
    )

    monthly_path = OUTPUT_DIR / f"유효생산량_결과_{ym}.xlsx"
    _write_excel(monthly_path, res_m, daily_m, factory_m, spec_m, spec_f_daily_m, spec_f_total_m)

# 2) 대시보드 기본 파일은 "전체기간" 누적본으로 유지
#    (월별 파일은 outputs/에 따로 생성되므로, 기본 파일에서 4월 데이터가 사라지지 않게 함)
base_path = Path(BASE_PATH) / "유효생산량_결과.xlsx"
existing = _read_existing_result_sources(base_path)
if existing is not None:
    # 최근 N개월 구간은 '교체' 반영(기존 월 데이터 제거 후, 새 계산분을 추가)
    matching_old = _drop_months(existing.get("매칭결과", pd.DataFrame()), "날짜", months_to_refresh)
    daily_old = _drop_months(existing.get("일별요약", pd.DataFrame()), "날짜", months_to_refresh)
    factory_old = _drop_months(existing.get("공장_신규분류별", pd.DataFrame()), "생산일자", months_to_refresh)
    spec_old = _drop_months(existing.get("일별_규격대응", pd.DataFrame()), "날짜", months_to_refresh)

    matching_new = _filter_by_months(matching_result_factory, "날짜", months_to_refresh)
    daily_new = _filter_by_months(daily_summary, "날짜", months_to_refresh)
    factory_new = _filter_by_months(factory_summary, "생산일자", months_to_refresh)
    spec_new = _filter_by_months(spec_coverage_daily, "날짜", months_to_refresh)

    matching_merged = _merge_keep_latest(
        matching_old,
        matching_new,
        subset_cols=[c for c in ["날짜", "생산일자", "공장", "신규분류요약", "제품코드"] if c in matching_result_factory.columns],
        sort_cols=["날짜", "생산일자"] if ("날짜" in matching_result_factory.columns and "생산일자" in matching_result_factory.columns) else None,
    )
    daily_merged = _merge_keep_latest(daily_old, daily_new, subset_cols=["날짜"], sort_cols=["날짜"])
    factory_merged = _merge_keep_latest(
        factory_old,
        factory_new,
        subset_cols=[c for c in ["생산일자", "공장", "신규분류요약"] if c in factory_summary.columns],
        sort_cols=["생산일자", "공장"] if ("생산일자" in factory_summary.columns and "공장" in factory_summary.columns) else None,
    )
    spec_merged = _merge_keep_latest(spec_old, spec_new, subset_cols=["날짜"], sort_cols=["날짜"])
    _write_excel(
        base_path,
        matching_merged,
        daily_merged,
        factory_merged,
        spec_merged,
        spec_coverage_factory_daily,
        spec_coverage_factory_total,
        archive_mode="copy",
    )
else:
    _write_excel(
        base_path,
        matching_result_factory,
        daily_summary,
        factory_summary,
        spec_coverage_daily,
        spec_coverage_factory_daily,
        spec_coverage_factory_total,
        archive_mode="copy",
    )

# 3) 전공정 결과 파일: 유효생산량_결과2.xlsx (공정별_일별실적)
base_path2 = Path(BASE_PATH) / "유효생산량_결과2.xlsx"
process_existing = _read_existing_result2_sources(base_path2)

matching_out = matching_result_proc.copy()
daily_out = daily_summary_proc.copy()
factory_out = factory_summary_proc.copy()
process_daily_out = process_daily_perf.copy()

for df, date_col in [
    (matching_out, "날짜"),
    (daily_out, "날짜"),
    (factory_out, "생산일자"),
    (process_daily_out, "날짜"),
]:
    if len(df) > 0 and date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

if process_existing is not None:
    matching_old = _drop_months(process_existing.get("매칭결과", pd.DataFrame()), "날짜", months_to_refresh)
    daily_old = _drop_months(process_existing.get("일별요약", pd.DataFrame()), "날짜", months_to_refresh)
    factory_old = _drop_months(process_existing.get("공장_신규분류별", pd.DataFrame()), "생산일자", months_to_refresh)
    process_old = _drop_months(process_existing.get("공정별_일별실적", pd.DataFrame()), "날짜", months_to_refresh)

    matching_new = _filter_by_months(matching_out, "날짜", months_to_refresh)
    daily_new = _filter_by_months(daily_out, "날짜", months_to_refresh)
    factory_new = _filter_by_months(factory_out, "생산일자", months_to_refresh)
    process_new = _filter_by_months(process_daily_out, "날짜", months_to_refresh)

    matching_merged2 = _merge_keep_latest(
        matching_old,
        matching_new,
        subset_cols=[c for c in ["날짜", "공장", "공정", "신규분류요약", "제품코드"] if c in matching_new.columns],
        sort_cols=["날짜", "공장", "공정"] if ("날짜" in matching_new.columns and "공장" in matching_new.columns and "공정" in matching_new.columns) else None,
    )
    daily_merged2 = _merge_keep_latest(daily_old, daily_new, subset_cols=[c for c in ["날짜", "공정"] if c in daily_new.columns], sort_cols=["날짜", "공정"] if ("날짜" in daily_new.columns and "공정" in daily_new.columns) else None)
    factory_merged2 = _merge_keep_latest(
        factory_old,
        factory_new,
        subset_cols=[c for c in ["생산일자", "공장", "공정", "신규분류요약"] if c in factory_new.columns],
        sort_cols=["생산일자", "공장", "공정"] if ("생산일자" in factory_new.columns and "공장" in factory_new.columns and "공정" in factory_new.columns) else None,
    )
    process_merged2 = _merge_keep_latest(process_old, process_new, subset_cols=[c for c in ["날짜", "공장", "공정"] if c in process_new.columns], sort_cols=["날짜", "공장", "공정"] if ("날짜" in process_new.columns and "공장" in process_new.columns and "공정" in process_new.columns) else None)

    _write_excel2(
        base_path2,
        process_merged2,
        matching=matching_merged2,
        daily=daily_merged2,
        factory=factory_merged2,
        archive_mode="copy",
    )
else:
    _write_excel2(
        base_path2,
        process_daily_out,
        matching=matching_out,
        daily=daily_out,
        factory=factory_out,
        archive_mode="copy",
    )

# 월별 파일도 outputs/에 저장 (최근 N개월만)
if len(process_daily_out) > 0 and months_to_refresh:
    p = process_daily_out.copy()
    period = pd.to_datetime(p["날짜"], errors="coerce").dt.to_period("M")
    for ym in sorted(months_to_refresh):
        pm = p.loc[period == ym].copy()
        if len(pm) == 0:
            continue
        monthly_path2 = OUTPUT_DIR / f"유효생산량_결과2_{ym}.xlsx"
        mm = _filter_by_months(matching_out, "날짜", {ym})
        dd = _filter_by_months(daily_out, "날짜", {ym})
        ff = _filter_by_months(factory_out, "생산일자", {ym})
        _write_excel2(monthly_path2, pm, matching=mm, daily=dd, factory=ff, archive_mode="copy")

print("완료")
