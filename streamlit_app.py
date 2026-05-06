import os
from pathlib import Path
import calendar
from datetime import datetime
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px


def _month_end(d: datetime.date) -> datetime.date:
    last_day = calendar.monthrange(d.year, d.month)[1]
    return datetime(d.year, d.month, last_day).date()


def _period_start(ts: pd.Series, bucket: str) -> pd.Series:
    if bucket == "D":
        return ts.dt.normalize()
    if bucket == "W":
        return ts.dt.normalize() - pd.to_timedelta(ts.dt.weekday, unit="D")
    return ts.dt.to_period("M").dt.to_timestamp()


def _build_axis(start_d: datetime.date, end_d: datetime.date, bucket: str) -> pd.DatetimeIndex:
    start_ts = pd.Timestamp(start_d)
    end_ts = pd.Timestamp(end_d)
    if bucket == "D":
        return pd.date_range(start_ts, end_ts, freq="D")
    if bucket == "W":
        start_monday = start_ts.normalize() - pd.to_timedelta(start_ts.weekday(), unit="D")
        end_monday = end_ts.normalize() - pd.to_timedelta(end_ts.weekday(), unit="D")
        return pd.date_range(start_monday, end_monday, freq="W-MON")
    start_ms = start_ts.to_period("M").to_timestamp()
    end_ms = end_ts.to_period("M").to_timestamp()
    return pd.date_range(start_ms, end_ms, freq="MS")


def _build_tick_labels(axis: pd.DatetimeIndex, bucket: str) -> tuple[list[pd.Timestamp], list[str]]:
    tickvals = [pd.Timestamp(x) for x in axis.to_list()]
    if bucket == "D":
        ticktext = [x.strftime("%m-%d") for x in tickvals]
        return tickvals, ticktext
    if bucket == "W":
        years = {x.year for x in tickvals}
        iso_weeks = pd.Series(tickvals).dt.isocalendar().week.astype(int).tolist()
        if len(years) > 1:
            ticktext = [f"{x.year % 100:02d}W{w}" for x, w in zip(tickvals, iso_weeks, strict=False)]
        else:
            ticktext = [f"W{w}" for w in iso_weeks]
        return tickvals, ticktext
    # bucket == "M"
    years = {x.year for x in tickvals}
    if len(years) > 1:
        ticktext = [f"{x.year}-{x.month}월" for x in tickvals]
    else:
        ticktext = [f"{x.month}월" for x in tickvals]
    return tickvals, ticktext

# 페이지 설정
DASHBOARD_TITLE = "생산 운영 현황 대시보드"
KPI_LABEL_MAP = {
    "총실적": "총 생산량",
    "총부족수량": "필요 수량",
    "유효생산량": "정확 대응 생산량",
    "과생산량": "초과 생산량",
    "불필요생산량": "비정형 생산량",
}
RATE_LABEL_MAP = {
    "유효비율(%)": "정확 대응 비중(%)",
    "과생산비율(%)": "초과 생산 비중(%)",
    "불필요비율(%)": "비정형 생산 비중(%)",
}

st.set_page_config(page_title=DASHBOARD_TITLE, layout="wide", initial_sidebar_state="collapsed")

# CSS 스타일링
st.markdown("""
<style>
    [data-testid="metric.container"] {
        background-color: #f0f4f8;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .kpi-card {
        background-color: #f0f4f8;
        border-radius: 12px;
        padding: 16px 16px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        height: 100%;
    }
    .kpi-head {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        gap: 10px;
        margin-bottom: 6px;
    }
    .kpi-title {
        font-size: 14px;
        font-weight: 700;
        color: #374151;
        line-height: 1.2;
    }
    .kpi-right {
        font-size: 13px;
        font-weight: 800;
        white-space: nowrap;
    }
    .kpi-value {
        font-size: clamp(22px, 2.2vw, 34px);
        font-weight: 900;
        color: #111827;
        letter-spacing: 0.3px;
        line-height: 1.0;
        margin: 0;
    }
    .kpi-sub {
        margin-top: 8px;
        font-size: 12px;
        color: #6b7280;
        line-height: 1.2;
    }
    .kpi-split {
        background-color: #f0f4f8;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        display: flex;
        overflow: hidden;
        height: 100%;
    }
    .kpi-cell {
        padding: 16px 16px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-width: 0;
    }
    .kpi-cell.left {
        flex: 1.6;
    }
    .kpi-cell.right {
        flex: 1.0;
    }
    .kpi-divider {
        width: 1px;
        background: rgba(17, 24, 39, 0.12);
        margin: 14px 0;
        flex: 0 0 1px;
    }
    .kpi-cell-title {
        font-size: 14px;
        font-weight: 700;
        color: #374151;
        line-height: 1.2;
        margin-bottom: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-cell-value {
        font-size: clamp(22px, 2.2vw, 34px);
        font-weight: 900;
        color: #111827;
        letter-spacing: 0.3px;
        line-height: 1.0;
        margin: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-cell-sub {
        margin-top: 8px;
        font-size: 12px;
        color: #6b7280;
        line-height: 1.2;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    @media (max-width: 900px) {
        .kpi-split {
            flex-direction: column;
        }
        .kpi-divider {
            width: 100%;
            height: 1px;
            margin: 0 14px;
        }
        .kpi-cell.left, .kpi-cell.right {
            flex: unset;
        }
    }
    h1 {
        text-align: center;
        color: #1f3a93;
        margin-bottom: 90px;
    }
    h2 {
        color: #2c5aa0;
        border-bottom: 3px solid #2c5aa0;
        padding-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


def render_kpi_card(title: str, value: str, right_label: str | None = None, right_value: float | None = None, right_color: str = "#111827", sub: str | None = None) -> None:
    right_html = ""
    if right_label is not None and right_value is not None:
        right_html = f"<span class='kpi-right' style='color:{right_color};'>{right_label} {right_value:.1f}%</span>"
    sub_html = f"<div class='kpi-sub'>{sub}</div>" if sub else ""
    st.markdown(
        f"""
<div class="kpi-card">
  <div class="kpi-head">
    <div class="kpi-title">{title}</div>
    {right_html}
  </div>
  <div class="kpi-value">{value}</div>
  {sub_html}
</div>
""",
        unsafe_allow_html=True,
    )


def render_kpi_split_card(
    left_title: str,
    left_value: str,
    right_title: str,
    right_value: str,
    right_sub: str | None = None,
) -> None:
    right_sub_html = f"<div class='kpi-cell-sub'>{right_sub}</div>" if right_sub else ""
    st.markdown(
        f"""
<div class="kpi-split">
  <div class="kpi-cell left">
    <div class="kpi-cell-title">{left_title}</div>
    <div class="kpi-cell-value">{left_value}</div>
  </div>
  <div class="kpi-divider"></div>
  <div class="kpi-cell right">
    <div class="kpi-cell-title">{right_title}</div>
    <div class="kpi-cell-value">{right_value}</div>
    {right_sub_html}
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_result_excels(
    result_paths: tuple[str, ...],
    mtime_nss: tuple[int, ...],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """결과 엑셀 파일(여러 개 가능) 로드 + 전처리 (Streamlit 캐시 적용)

    - 월별로 파일이 분리되어 저장되는 경우(예: 유효생산량_결과_2026-05.xlsx)에도
      전월/기간조회가 동작하도록 여러 결과 파일을 합쳐서 사용합니다.
    - mtime_nss는 파일 변경 시 캐시 무효화를 위해 사용됩니다.
    """
    _ = mtime_nss  # cache key only

    required_sheets = ["매칭결과", "일별요약", "공장_신규분류별"]
    matching_frames: list[pd.DataFrame] = []
    daily_frames: list[pd.DataFrame] = []
    factory_frames: list[pd.DataFrame] = []

    for path_str, mtime_ns in zip(result_paths, mtime_nss, strict=False):
        path = Path(path_str)
        sheets = pd.read_excel(path, sheet_name=required_sheets)
        required = set(required_sheets)
        missing = required - set(sheets.keys())
        if missing:
            raise ValueError(f"결과 엑셀에 필요한 시트가 없습니다({path.name}): {', '.join(sorted(missing))}")

        mr = sheets["매칭결과"].copy()
        ds = sheets["일별요약"].copy()
        fs = sheets["공장_신규분류별"].copy()

        mr["_source_mtime_ns"] = mtime_ns
        ds["_source_mtime_ns"] = mtime_ns
        fs["_source_mtime_ns"] = mtime_ns

        matching_frames.append(mr)
        daily_frames.append(ds)
        factory_frames.append(fs)

    matching_result = pd.concat(matching_frames, ignore_index=True) if matching_frames else pd.DataFrame()
    daily_summary = pd.concat(daily_frames, ignore_index=True) if daily_frames else pd.DataFrame()
    factory_summary = pd.concat(factory_frames, ignore_index=True) if factory_frames else pd.DataFrame()

    if len(matching_result) > 0:
        matching_result["날짜"] = pd.to_datetime(matching_result["날짜"], errors="coerce")
        matching_result["생산일자"] = pd.to_datetime(matching_result["생산일자"], errors="coerce")
        matching_result = matching_result.sort_values("_source_mtime_ns", kind="stable")
        dedup_cols = [c for c in ["날짜", "생산일자", "공장", "신규분류요약", "제품코드"] if c in matching_result.columns]
        if dedup_cols:
            matching_result = matching_result.drop_duplicates(subset=dedup_cols, keep="last")
        matching_result["날짜_date"] = matching_result["날짜"].dt.date
        matching_result["생산일자_date"] = matching_result["생산일자"].dt.date
        matching_result = matching_result.drop(columns=["_source_mtime_ns"], errors="ignore")

    if len(daily_summary) > 0:
        daily_summary["날짜"] = pd.to_datetime(daily_summary["날짜"], errors="coerce")
        daily_summary = daily_summary[daily_summary["날짜"].notna()].copy()
        daily_summary = daily_summary.sort_values("_source_mtime_ns", kind="stable")
        if "날짜" in daily_summary.columns:
            daily_summary = daily_summary.drop_duplicates(subset=["날짜"], keep="last")
        daily_summary["날짜_date"] = daily_summary["날짜"].dt.date
        daily_summary = daily_summary.drop(columns=["_source_mtime_ns"], errors="ignore")

    if len(factory_summary) > 0:
        factory_summary["생산일자"] = pd.to_datetime(factory_summary["생산일자"], errors="coerce")
        factory_summary = factory_summary[factory_summary["생산일자"].notna()].copy()
        factory_summary = factory_summary.sort_values("_source_mtime_ns", kind="stable")
        dedup_cols = [c for c in ["생산일자", "공장", "신규분류요약"] if c in factory_summary.columns]
        if dedup_cols:
            factory_summary = factory_summary.drop_duplicates(subset=dedup_cols, keep="last")
        factory_summary["생산일자_date"] = factory_summary["생산일자"].dt.date
        factory_summary = factory_summary.drop(columns=["_source_mtime_ns"], errors="ignore")

    # SKU 기반 일자 규격 대응률(전사/공장별) 프리컴퓨트 (대용량 groupby는 1회만 수행)
    sku_daily_all = pd.DataFrame()
    sku_daily_factory = pd.DataFrame()
    required_cols = {"날짜_date", "제품코드", "양품수량", "부족수량", "유효생산량"}
    if len(matching_result) > 0 and required_cols.issubset(set(matching_result.columns)):
        base = matching_result[matching_result["제품코드"].notna()].copy()
        for col in ["양품수량", "부족수량", "유효생산량"]:
            base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)
        base["_need_qty"] = (base["유효생산량"] + base["부족수량"]).fillna(0)

        by_day_sku = base.groupby(["날짜_date", "제품코드"], dropna=False).agg(
            prod_qty=("양품수량", "sum"),
            need_qty=("_need_qty", "sum"),
        ).reset_index()
        by_day_sku["produced_flag"] = by_day_sku["prod_qty"] > 0
        by_day_sku["need_flag"] = by_day_sku["need_qty"] > 0

        produced_skus = (
            by_day_sku[by_day_sku["produced_flag"]]
            .groupby("날짜_date", dropna=False)["제품코드"]
            .nunique()
            .rename("생산SKU수")
        )
        needed_skus = (
            by_day_sku[by_day_sku["produced_flag"] & by_day_sku["need_flag"]]
            .groupby("날짜_date", dropna=False)["제품코드"]
            .nunique()
            .rename("필요대응SKU수")
        )
        sku_daily_all = pd.concat([produced_skus, needed_skus], axis=1).fillna(0).reset_index()
        sku_daily_all["규격대응률(%)"] = np.where(
            sku_daily_all["생산SKU수"] > 0,
            sku_daily_all["필요대응SKU수"] / sku_daily_all["생산SKU수"] * 100,
            0,
        )
        sku_daily_all["규격대응률(%)"] = sku_daily_all["규격대응률(%)"].clip(0, 100)
        sku_daily_all = sku_daily_all.sort_values("날짜_date").reset_index(drop=True)

        if "공장" in matching_result.columns:
            base_f = matching_result[(matching_result["공장"].notna()) & (matching_result["제품코드"].notna())].copy()
            for col in ["양품수량", "부족수량", "유효생산량"]:
                base_f[col] = pd.to_numeric(base_f[col], errors="coerce").fillna(0)
            base_f["_need_qty"] = (base_f["유효생산량"] + base_f["부족수량"]).fillna(0)

            by_day_factory_sku = base_f.groupby(["날짜_date", "공장", "제품코드"], dropna=False).agg(
                prod_qty=("양품수량", "sum"),
                need_qty=("_need_qty", "sum"),
            ).reset_index()
            by_day_factory_sku["produced_flag"] = by_day_factory_sku["prod_qty"] > 0
            by_day_factory_sku["need_flag"] = by_day_factory_sku["need_qty"] > 0

            produced_f = (
                by_day_factory_sku[by_day_factory_sku["produced_flag"]]
                .groupby(["날짜_date", "공장"], dropna=False)["제품코드"]
                .nunique()
                .rename("생산SKU수")
            )
            needed_f = (
                by_day_factory_sku[by_day_factory_sku["produced_flag"] & by_day_factory_sku["need_flag"]]
                .groupby(["날짜_date", "공장"], dropna=False)["제품코드"]
                .nunique()
                .rename("필요대응SKU수")
            )
            sku_daily_factory = pd.concat([produced_f, needed_f], axis=1).fillna(0).reset_index()
            sku_daily_factory["규격대응률(%)"] = np.where(
                sku_daily_factory["생산SKU수"] > 0,
                sku_daily_factory["필요대응SKU수"] / sku_daily_factory["생산SKU수"] * 100,
                0,
            )
            sku_daily_factory["규격대응률(%)"] = sku_daily_factory["규격대응률(%)"].clip(0, 100)
            sku_daily_factory = sku_daily_factory.sort_values(["날짜_date", "공장"]).reset_index(drop=True)

    return matching_result, daily_summary, factory_summary, sku_daily_all, sku_daily_factory


# 결과 파일 선택(월별 분리 저장 지원)
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
base_dir = Path(BASE_PATH)

# 결과 파일이 repo 루트뿐 아니라 `outputs/` 아래에 저장되는 경우도 있어 함께 검색합니다.
search_dirs = [base_dir, base_dir / "outputs", base_dir / "outputs" / "archive"]
_cands: list[Path] = []
for d in search_dirs:
    if not d.exists():
        continue
    _cands.extend([p for p in d.glob("유효생산량_결과*.xlsx") if not p.name.startswith("~$")])

_seen: set[str] = set()
result_candidates: list[Path] = []
for p in _cands:
    rp = str(p.resolve())
    if rp in _seen:
        continue
    _seen.add(rp)
    result_candidates.append(p)

result_candidates = sorted(
    result_candidates,
    key=lambda p: p.stat().st_mtime_ns if p.exists() else 0,
    reverse=True,
)
if not result_candidates:
    st.error(
        "⚠️ 결과 파일을 찾을 수 없습니다. 검색 경로: "
        + ", ".join(str(d) for d in search_dirs)
    )
    st.info("전처리 완료된 결과 파일(`유효생산량_결과*.xlsx`)을 repo 루트 또는 `outputs/`에 넣어주세요.")
    st.stop()

try:
    # 최신 파일이 월별로 분리되어 저장될 수 있어, 후보 파일들을 합쳐서 사용
    result_paths = tuple(str(p) for p in result_candidates)
    mtime_nss = tuple(int(p.stat().st_mtime_ns) for p in result_candidates)
    matching_result, daily_summary, factory_summary, sku_daily_all, sku_daily_factory = load_result_excels(result_paths, mtime_nss)

    # 금일 데이터 제외 (아직 생산 중이므로) - KST 기준
    now_kst = datetime.now(ZoneInfo("Asia/Seoul"))
    today = now_kst.date()
    st.caption(f"기준 시각(KST): {now_kst.strftime('%Y-%m-%d %H:%M:%S')}")

    # 필수 데이터 검증/정규화
    if daily_summary is None or len(daily_summary) == 0 or "날짜_date" not in daily_summary.columns:
        st.error("⚠️ `일별요약` 시트에 날짜 데이터가 없어서 대시보드를 표시할 수 없습니다. (컬럼: `날짜`)")
        st.info("`유효생산량_결과*.xlsx`를 최신 버전으로 다시 생성한 뒤, repo 루트 또는 `outputs/`에 넣어주세요.")
        st.stop()

    for col in ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]:
        if col in daily_summary.columns:
            daily_summary[col] = pd.to_numeric(daily_summary[col], errors="coerce").fillna(0)

    factory_has_dates = factory_summary is not None and len(factory_summary) > 0 and "생산일자_date" in factory_summary.columns
    if factory_has_dates:
        for col in ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]:
            if col in factory_summary.columns:
                factory_summary[col] = pd.to_numeric(factory_summary[col], errors="coerce").fillna(0)
    else:
        factory_summary = pd.DataFrame()

    # 제목
    st.markdown(f"<h1 style='text-align:center; color:#1f3a93; margin:0;'>🏭 {DASHBOARD_TITLE}</h1>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # 기간 필터 (기본: 당월)
    filter_option = st.radio("조회 기간", ["당월", "전월", "기간조회"], horizontal=True, label_visibility="collapsed")

    # 날짜 범위 계산
    current_month_start = datetime(today.year, today.month, 1).date()
    data_max_date = daily_summary["날짜_date"].max()
    current_month_end = (today - pd.Timedelta(days=1))  # 어제까지 (date)
    if pd.notna(data_max_date):
        current_month_end = min(current_month_end, data_max_date)

    # 전월 계산
    first_day_current = current_month_start
    last_day_prev = first_day_current - pd.Timedelta(days=1)
    prev_month_start = datetime(last_day_prev.year, last_day_prev.month, 1).date()

    # 전체 기간(데이터 기준) 계산 (기간조회 범위 제한용)
    full_min_date = daily_summary[daily_summary["날짜_date"] != today]["날짜_date"].min()
    full_max_date = daily_summary[daily_summary["날짜_date"] != today]["날짜_date"].max()

    # 날짜 범위 결정
    if filter_option == "당월":
        start_date = current_month_start
        end_date = current_month_end
    elif filter_option == "전월":
        start_date = prev_month_start
        end_date = last_day_prev
    else:  # 기간조회
        min_date = full_min_date
        max_date = full_max_date
        if pd.isna(min_date) or pd.isna(max_date):
            st.warning("선택 가능한 날짜 범위를 계산할 수 없습니다. (데이터 없음)")
            min_date = today
            max_date = today

        col_filter1, col_space, col_filter2 = st.columns([1.5, 0.2, 1.5])

        with col_filter1:
            start_date = st.date_input("시작 날짜", value=min_date, min_value=min_date, max_value=max_date)

        with col_filter2:
            end_date = st.date_input("종료 날짜", value=max_date, min_value=min_date, max_value=max_date)

    if start_date > end_date:
        st.warning("시작 날짜가 종료 날짜보다 커서 자동으로 교체했습니다.")
        start_date, end_date = end_date, start_date

    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)

    # 필터 적용 (기간 범위)
    daily_summary_filtered = daily_summary[
        (daily_summary["날짜_date"] >= start_date) &
        (daily_summary["날짜_date"] <= end_date) &
        (daily_summary["날짜_date"] != today)
    ]

    if factory_has_dates and len(factory_summary) > 0:
        factory_summary_filtered = factory_summary[
            (factory_summary["생산일자_date"] >= start_date) &
            (factory_summary["생산일자_date"] <= end_date) &
            (factory_summary["생산일자_date"] != today)
        ]
    else:
        factory_summary_filtered = pd.DataFrame()

    # 메트릭 계산
    total_prod = int(daily_summary_filtered["총실적"].sum()) if len(daily_summary_filtered) > 0 else 0
    valid_prod = int(daily_summary_filtered["유효생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    over_prod = int(daily_summary_filtered["과생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    waste_prod = int(daily_summary_filtered["불필요생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    prod_days = int(daily_summary_filtered["날짜_date"].nunique()) if len(daily_summary_filtered) > 0 else 0

    valid_rate = (valid_prod / total_prod * 100) if total_prod > 0 else 0
    over_rate = (over_prod / total_prod * 100) if total_prod > 0 else 0
    waste_rate = (waste_prod / total_prod * 100) if total_prod > 0 else 0

    # 규격 대응률(SKU 기준): "그날 생산한 SKU" 중 "그날 필요(수요)가 있던 SKU" 비율
    # - 사용자 정의: (일자별 필요 SKU ∩ 일자별 생산 SKU) / 일자별 생산 SKU
    # - 공장별 규격 대응률도 동일 기준으로 계산하려면 `매칭결과` 시트에 `공장` 컬럼이 필요합니다.
    shortage_prod_daily = None
    shortage_prod_rate = None
    if sku_daily_all is not None and len(sku_daily_all) > 0 and {"날짜_date", "생산SKU수", "필요대응SKU수", "규격대응률(%)"}.issubset(set(sku_daily_all.columns)):
        shortage_prod_daily = sku_daily_all[
            (sku_daily_all["날짜_date"] >= start_date) &
            (sku_daily_all["날짜_date"] <= end_date) &
            (sku_daily_all["날짜_date"] != today)
        ].copy()
        if len(shortage_prod_daily) > 0:
            produced_skus_total = float(pd.to_numeric(shortage_prod_daily["생산SKU수"], errors="coerce").fillna(0).sum())
            need_responded_skus_total = float(pd.to_numeric(shortage_prod_daily["필요대응SKU수"], errors="coerce").fillna(0).sum())
            shortage_prod_rate = (need_responded_skus_total / produced_skus_total * 100) if produced_skus_total > 0 else None

    colA, col3, col4, col5 = st.columns([2.6, 1.1, 1.1, 1.1])
    with colA:
        spec_value = f"{shortage_prod_rate:.1f}%" if shortage_prod_rate is not None else "-"
        spec_sub = "일자별 (필요SKU∩생산SKU) / 생산SKU"
        if shortage_prod_rate is None:
            spec_sub = "계산 불가: 매칭결과에 제품코드/수량 필요"
        render_kpi_split_card(
            f"{KPI_LABEL_MAP['총실적']} (pcs)",
            f"{total_prod:,}",
            "규격 대응률 (%)",
            f"<span style='color:#1d4ed8'>{spec_value}</span>",
            right_sub=spec_sub,
        )

    with col3:
        render_kpi_card(
            "정확 대응 비중",
            f"<span style='color:#047857'>{valid_rate:.1f}%</span>",
            sub=f"수량: {valid_prod:,} pcs",
        )
    with col4:
        render_kpi_card(
            "초과 생산 비중",
            f"<span style='color:#b91c1c'>{over_rate:.1f}%</span>",
            sub=f"수량: {over_prod:,} pcs",
        )
    with col5:
        render_kpi_card(
            "비정형 생산 비중",
            f"<span style='color:#b45309'>{waste_rate:.1f}%</span>",
            sub=f"수량: {waste_prod:,} pcs",
        )

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    with st.expander("지표 정의/상세 보기", expanded=False):
        st.markdown(
            "- `규격 대응률` : 일자별 `(필요 SKU ∩ 생산 SKU) ÷ 생산 SKU` 의 비율\n"
            "- `정확 대응 생산량` : SKU별 `min(생산, 필요)`의 합\n"
            "- `정확 대응 비중` : `정확 대응 생산량` ÷ `총 생산량`\n"
            "- `초과 생산량` : SKU별 `max(생산-필요, 0)`의 합\n"
            "- `초과 생산 비중` : `초과 생산량` ÷ `총 생산량`\n"
            "- `비정형 생산량` : 필요 SKU 외 생산(필요=0인데 생산>0)\n"
            "- `비정형 생산 비중` : `비정형 생산량` ÷ `총 생산량`"
        )
        st.caption("참고: 공장별 `규격 대응률(SKU 기준)`은 `매칭결과` 시트에 `공장`/`제품코드`가 있어야 계산 가능합니다.")

    st.markdown("<div style='margin-top:50px'></div>", unsafe_allow_html=True)

    st.markdown("### 📈 공장별 운영 현황")

    if len(factory_summary_filtered) == 0:
        st.info("선택한 기간에 공장별 데이터가 없습니다.")
    else:
        # 공장별 기간 집계 (정확/초과/비정형 분해)
        factory_data = factory_summary_filtered.groupby("공장", dropna=False).agg(
            {
                "총실적": "sum",
                "유효생산량": "sum",
                "과생산량": "sum",
                "불필요생산량": "sum",
            }
        ).reset_index()

        factory_data["유효비율(%)"] = (factory_data["유효생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["과생산비율(%)"] = (factory_data["과생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["불필요비율(%)"] = (factory_data["불필요생산량"] / factory_data["총실적"] * 100).fillna(0)

        # 공장별 KPI (정확도 기반)
        factory_data["유효 대응률(수량)(%)"] = factory_data["유효비율(%)"]

        # 공장별 규격 대응률(SKU 기준): 일자별 생산 SKU 중 필요 SKU 비중의 기간 합산
        # - 정의: (Σ 일자별 필요대응SKU수) / (Σ 일자별 생산SKU수)
        sku_coverage_available = False
        sku_coverage_unavailable_reason: str | None = None
        required_cols = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
        if sku_daily_factory is None or len(sku_daily_factory) == 0 or not required_cols.issubset(set(sku_daily_factory.columns)):
            sku_coverage_unavailable_reason = "필수 컬럼 누락(`공장/제품코드/수량/날짜`)"
        else:
            day_counts = sku_daily_factory[
                (sku_daily_factory["날짜_date"] >= start_date) &
                (sku_daily_factory["날짜_date"] <= end_date) &
                (sku_daily_factory["날짜_date"] != today)
            ].copy()
            if len(day_counts) == 0:
                sku_coverage_unavailable_reason = "선택 기간에 매칭결과 데이터 없음"
            else:
                sku_counts = day_counts.groupby("공장", dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                sku_counts["규격대응률(%)"] = np.where(
                    sku_counts["생산SKU수"] > 0,
                    sku_counts["필요대응SKU수"] / sku_counts["생산SKU수"] * 100,
                    0,
                )
                sku_counts["규격대응률(%)"] = sku_counts["규격대응률(%)"].clip(0, 100)
                factory_data = factory_data.merge(
                    sku_counts[["공장", "생산SKU수", "필요대응SKU수", "규격대응률(%)"]],
                    on="공장",
                    how="left",
                )
                if "생산SKU수" in factory_data.columns:
                    factory_data["생산SKU수"] = pd.to_numeric(factory_data["생산SKU수"], errors="coerce").fillna(0)
                else:
                    factory_data["생산SKU수"] = 0
                if "필요대응SKU수" in factory_data.columns:
                    factory_data["필요대응SKU수"] = pd.to_numeric(factory_data["필요대응SKU수"], errors="coerce").fillna(0)
                else:
                    factory_data["필요대응SKU수"] = 0
                if "규격대응률(%)" in factory_data.columns:
                    factory_data["규격대응률(%)"] = (
                        pd.to_numeric(factory_data["규격대응률(%)"], errors="coerce")
                        .replace([np.inf, -np.inf], 0)
                        .fillna(0)
                    )
                else:
                    factory_data["규격대응률(%)"] = 0.0
                sku_coverage_available = True

        # NOTE: 규격 대응률은 메인 지표이므로 항상 노출합니다.
        # 계산이 불가한 경우(예: 매칭결과에 공장 없음)에는 "전사 규격 대응률"을 동일 적용해 표시합니다.
        if "규격대응률(%)" not in factory_data.columns:
            factory_data["규격대응률(%)"] = np.nan
        if not sku_coverage_available:
            factory_data["규격대응률(%)"] = float(shortage_prod_rate) if shortage_prod_rate is not None else np.nan
        metric_choices = ["규격 대응률", "정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"]
        radio_key = "factory_metric_option"
        if radio_key not in st.session_state or st.session_state[radio_key] not in metric_choices:
            st.session_state[radio_key] = metric_choices[0]
        metric_option = st.radio("공장 비교 지표", metric_choices, horizontal=True, key=radio_key)
        metric_desc = {
            "규격 대응률": "생산한 SKU(제품코드) 중 필요가 있었던 SKU 비중",
            "정확 대응 비중": "총 생산량 중 정확 대응 생산량이 차지하는 비중",
            "초과 생산 비중": "총 생산량 중 초과 생산량이 차지하는 비중",
            "비정형 생산 비중": "총 생산량 중 비정형 생산량이 차지하는 비중",
        }
        st.caption(f"설명: {metric_desc[metric_option]}")

        metric_map = {
            "규격 대응률": ("규격대응률(%)", "유효생산량"),
            "정확 대응 비중": ("유효비율(%)", "유효생산량"),
            "초과 생산 비중": ("과생산비율(%)", "과생산량"),
            "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
        }
        metric_col, pcs_col = metric_map[metric_option]
        factory_data["선택지표"] = factory_data[metric_col].replace([np.inf, -np.inf], 0).fillna(0)
        if metric_option == "규격 대응률" and not sku_coverage_available:
            reason = sku_coverage_unavailable_reason or "원인 미상"
            st.warning(f"공장별 `규격 대응률(SKU 기준)` 계산 불가: {reason}. (전사 규격 대응률을 동일 적용해 표시)")

        hover_data = {
            "총실적": ":,",
            "유효생산량": ":,",
            "과생산량": ":,",
            "불필요생산량": ":,",
            "생산SKU수": ":,",
            "필요대응SKU수": ":,",
            "규격대응률(%)": ":.1f",
            "유효비율(%)": ":.1f",
            "과생산비율(%)": ":.1f",
            "불필요비율(%)": ":.1f",
            "선택지표": ":.1f",
        }
        hover_data = {k: v for k, v in hover_data.items() if k in factory_data.columns}

        fig = px.bar(
            factory_data,
            x="공장",
            y="선택지표",
            color="공장",
            title=f"공장별 {metric_option} (%)",
            text="선택지표",
            hover_data=hover_data,
        )
        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
            textfont=dict(size=24, family="Arial", color="#222222"),
            marker=dict(cornerradius="15"),
            cliponaxis=False,
        )
        fig.update_layout(
            height=520,
            showlegend=False,
            margin=dict(l=0, r=0, t=60, b=0),
            yaxis=dict(range=[0, 105], title=dict(text=f"{metric_option} (%)", font=dict(size=16, family="Arial", color="#222222"))),
            xaxis=dict(
                title=dict(text="공장", font=dict(size=16, family="Arial", color="#222222")),
                tickfont=dict(size=18, family="Arial", color="#222222")
            ),
            title=dict(font=dict(size=22, family="Arial", color="#111111"))
        )
        st.plotly_chart(fig, use_container_width=True)

        # 공장별 날짜 추이 (라인 차트)
        display_start_date = start_date
        display_end_date = end_date
        if filter_option == "당월":
            display_end_date = _month_end(display_start_date)

        if filter_option in {"당월", "전월"}:
            bucket = "D"
        else:
            span_days = (display_end_date - display_start_date).days + 1
            if span_days <= 30:
                bucket = "D"
            elif span_days <= 210:
                bucket = "W"
            else:
                bucket = "M"
        axis = _build_axis(display_start_date, display_end_date, bucket)
        tickvals, ticktext = _build_tick_labels(axis, bucket)

        factories = [f for f in factory_data["공장"].dropna().astype(str).unique().tolist()]
        ts_rows: list[dict] = []

        if metric_option != "규격 대응률":
            base_ts = factory_summary_filtered[
                ["생산일자_date", "공장", "총실적", "유효생산량", "과생산량", "불필요생산량"]
            ].copy()
            base_ts["date"] = pd.to_datetime(base_ts["생산일자_date"], errors="coerce")
            base_ts = base_ts.dropna(subset=["date"])
            base_ts["period"] = _period_start(base_ts["date"], bucket)
            agg = base_ts.groupby(["period", "공장"], dropna=False).agg(
                total=("총실적", "sum"),
                valid=("유효생산량", "sum"),
                over=("과생산량", "sum"),
                waste=("불필요생산량", "sum"),
            ).reset_index()

            num_col = {
                "정확 대응 비중": "valid",
                "초과 생산 비중": "over",
                "비정형 생산 비중": "waste",
            }[metric_option]
            agg["value"] = np.where(agg["total"] > 0, agg[num_col] / agg["total"] * 100, np.nan)
            agg["value"] = pd.to_numeric(agg["value"], errors="coerce").clip(0, 100)

            for _, r in agg.iterrows():
                ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
        else:
            spec_done = False
            required_cols_ts = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
            if sku_daily_factory is not None and len(sku_daily_factory) > 0 and required_cols_ts.issubset(set(sku_daily_factory.columns)):
                day_counts_ts = sku_daily_factory[
                    (sku_daily_factory["날짜_date"] >= start_date) &
                    (sku_daily_factory["날짜_date"] <= end_date) &
                    (sku_daily_factory["날짜_date"] != today)
                ].copy()
                if len(day_counts_ts) > 0:
                    day_counts_ts["date"] = pd.to_datetime(day_counts_ts["날짜_date"], errors="coerce")
                    day_counts_ts = day_counts_ts.dropna(subset=["date"])
                    day_counts_ts["period"] = _period_start(day_counts_ts["date"], bucket)
                    agg_ts = day_counts_ts.groupby(["period", "공장"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                    agg_ts["value"] = np.where(
                        agg_ts["생산SKU수"] > 0,
                        agg_ts["필요대응SKU수"] / agg_ts["생산SKU수"] * 100,
                        np.nan,
                    )
                    agg_ts["value"] = pd.to_numeric(agg_ts["value"], errors="coerce").clip(0, 100)
                    for _, r in agg_ts.iterrows():
                        ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
                    spec_done = True

            if (not spec_done) and (sku_daily_all is not None) and (len(sku_daily_all) > 0) and {"날짜_date", "생산SKU수", "필요대응SKU수"}.issubset(set(sku_daily_all.columns)):
                daily_spec = sku_daily_all[
                    (sku_daily_all["날짜_date"] >= start_date) &
                    (sku_daily_all["날짜_date"] <= end_date) &
                    (sku_daily_all["날짜_date"] != today)
                ].copy()
                if len(daily_spec) > 0:
                    daily_spec["date"] = pd.to_datetime(daily_spec["날짜_date"], errors="coerce")
                    daily_spec = daily_spec.dropna(subset=["date"])
                    daily_spec["period"] = _period_start(daily_spec["date"], bucket)
                    agg_spec = daily_spec.groupby(["period"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                    agg_spec["value"] = np.where(
                        agg_spec["생산SKU수"] > 0,
                        agg_spec["필요대응SKU수"] / agg_spec["생산SKU수"] * 100,
                        np.nan,
                    )
                    agg_spec["value"] = pd.to_numeric(agg_spec["value"], errors="coerce").clip(0, 100)
                    for _, r in agg_spec.iterrows():
                        for f in factories:
                            ts_rows.append({"기간": r["period"], "공장": f, "값": r["value"]})

        ts_df = pd.DataFrame(ts_rows)
        if len(ts_df) > 0:
            ts_df["기간"] = pd.to_datetime(ts_df["기간"], errors="coerce")
            full_grid = pd.MultiIndex.from_product([axis, factories], names=["기간", "공장"]).to_frame(index=False)
            ts_df = full_grid.merge(ts_df, on=["기간", "공장"], how="left")
            label_map = {pd.Timestamp(v): t for v, t in zip(tickvals, ticktext, strict=False)}
            ts_df["x_label"] = ts_df["기간"].map(label_map)

            line_fig = px.line(
                ts_df,
                x="기간",
                y="값",
                color="공장",
                title=f"공장별 {metric_option} 추이",
                markers=False,
                custom_data=["x_label"],
            )
            line_fig.update_traces(
                line=dict(width=3.5),
                hovertemplate="공장=%{legendgroup}<br>기간=%{customdata[0]}<br>값=%{y:.1f}%<extra></extra>",
            )
            line_fig.update_layout(
                height=360,
                margin=dict(l=0, r=0, t=60, b=0),
                yaxis=dict(range=[0, 105], title=f"{metric_option} (%)", tickformat=".1f"),
                xaxis=dict(
                    tickmode="array",
                    tickvals=tickvals,
                    ticktext=ticktext,
                    tickangle=-45,
                    tickfont=dict(size=10),
                ),
                legend_title_text="공장",
            )
            st.plotly_chart(line_fig, use_container_width=True)

        st.markdown(f"**선택 지표: {metric_option} (%)**")
        if not sku_coverage_available:
            st.caption("Tip: 공장별 `규격 대응률(SKU 기준)`은 `매칭결과` 시트에 `공장` 컬럼이 있어야 계산 가능합니다.")

        if metric_option == "규격 대응률":
            if not sku_coverage_available:
                st.info("공장별 SKU 집계가 불가합니다: `매칭결과` 시트에 `공장` 컬럼이 필요합니다.")
            else:
                if matching_result is None or len(matching_result) == 0 or not {"공장", "신규분류요약", "제품코드", "양품수량", "부족수량", "유효생산량", "날짜_date"}.issubset(set(matching_result.columns)):
                    st.info("신규분류 기준 SKU 상세 집계를 위해 `매칭결과`에 `공장/신규분류요약/제품코드/수량/날짜` 컬럼이 필요합니다.")
                else:
                    mf2 = matching_result[
                        (matching_result["날짜_date"] >= start_date) &
                        (matching_result["날짜_date"] <= end_date) &
                        (matching_result["날짜_date"] != today) &
                        (matching_result["공장"].notna()) &
                        (matching_result["제품코드"].notna())
                    ].copy()

                    if len(mf2) == 0:
                        st.info("선택한 기간에 신규분류 기준 SKU 집계 데이터가 없습니다.")
                    else:
                        for col in ["양품수량", "부족수량", "유효생산량"]:
                            mf2[col] = pd.to_numeric(mf2[col], errors="coerce").fillna(0)
                        mf2["_need_qty"] = (mf2["유효생산량"] + mf2["부족수량"]).fillna(0)

                        by_day_factory_class_sku = mf2.groupby(["날짜_date", "공장", "신규분류요약", "제품코드"], dropna=False).agg(
                            prod_qty=("양품수량", "sum"),
                            need_qty=("_need_qty", "sum"),
                        ).reset_index()
                        by_day_factory_class_sku["produced_flag"] = by_day_factory_class_sku["prod_qty"] > 0
                        by_day_factory_class_sku["need_flag"] = by_day_factory_class_sku["need_qty"] > 0

                        produced = (
                            by_day_factory_class_sku[by_day_factory_class_sku["produced_flag"]]
                            .groupby(["날짜_date", "공장", "신규분류요약"], dropna=False)["제품코드"]
                            .nunique()
                            .rename("생산SKU수")
                        )
                        needed = (
                            by_day_factory_class_sku[by_day_factory_class_sku["produced_flag"] & by_day_factory_class_sku["need_flag"]]
                            .groupby(["날짜_date", "공장", "신규분류요약"], dropna=False)["제품코드"]
                            .nunique()
                            .rename("필요대응SKU수")
                        )
                        day_counts = pd.concat([produced, needed], axis=1).fillna(0).reset_index()
                        sku_counts = day_counts.groupby(["공장", "신규분류요약"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                        sku_counts["규격대응률(%)"] = np.where(
                            sku_counts["생산SKU수"] > 0,
                            sku_counts["필요대응SKU수"] / sku_counts["생산SKU수"] * 100,
                            0,
                        )
                        sku_counts["규격대응률(%)"] = sku_counts["규격대응률(%)"].clip(0, 100)

                        factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
                        sku_counts["_factory_sort"] = sku_counts["공장"].map(factory_order)
                        sku_counts = sku_counts.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True).drop("_factory_sort", axis=1)

                        sku_counts_fmt = sku_counts.copy()
                        sku_counts_fmt["생산SKU수"] = sku_counts_fmt["생산SKU수"].map("{:,.0f}".format)
                        sku_counts_fmt["필요대응SKU수"] = sku_counts_fmt["필요대응SKU수"].map("{:,.0f}".format)
                        sku_counts_fmt["규격대응률(%)"] = sku_counts_fmt["규격대응률(%)"].map("{:.1f}%".format)
                        sku_counts_fmt["신규분류요약"] = sku_counts_fmt["신규분류요약"].fillna("미분류")

                        html_parts = []
                        header_lines = [
                            "<style>",
                            ".custom-table { width: 100%; border-collapse: collapse; font-size: 14px; }",
                            ".custom-table th, .custom-table td { padding: 10px 12px; border: 1px solid #e2e8f0; }",
                            ".custom-table th { background: #f8fafc; color: #111827; text-align: left; }",
                            ".custom-table td { vertical-align: middle; }",
                            ".custom-table td.number { text-align: right; }",
                            ".custom-table tbody tr:nth-child(even) { background: #f8fafc22; }",
                            "</style>",
                            "<table class=\"custom-table\">",
                            "<thead>",
                            "<tr>",
                            "<th>공장</th>",
                            "<th>신규분류요약</th>",
                            "<th>총 생산 SKU</th>",
                            "<th>규격 대응 SKU</th>",
                            "<th>규격 대응률(%)</th>",
                            "</tr>",
                            "</thead>",
                            "<tbody>",
                        ]
                        html_parts.append("\n".join(header_lines) + "\n")

                        grouped = sku_counts_fmt.groupby("공장", sort=False)
                        for factory_name, group in grouped:
                            rowspan = len(group)
                            for idx, row in group.iterrows():
                                html_parts.append("<tr>")
                                if idx == group.index[0]:
                                    html_parts.append(f"<td rowspan='{rowspan}' style='vertical-align: middle; font-weight: 600;'>{factory_name}</td>")
                                html_parts.append(f"<td>{row['신규분류요약']}</td>")
                                html_parts.append(f"<td class='number'>{row['생산SKU수']}</td>")
                                html_parts.append(f"<td class='number'>{row['필요대응SKU수']}</td>")
                                html_parts.append(f"<td class='number'>{row['규격대응률(%)']}</td>")
                                html_parts.append("</tr>")

                        html_parts.append("</tbody></table>")
                        st.markdown("".join(html_parts), unsafe_allow_html=True)
        else:
            # 공장_신규분류별 통합 현황
            combined_metric_option = metric_option if metric_option in {"정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"} else "정확 대응 비중"
            combined_summary = factory_summary_filtered.groupby(["공장", "신규분류요약"], dropna=False).agg({
                "총실적": "sum",
                "유효생산량": "sum",
                "과생산량": "sum",
                "불필요생산량": "sum"
            }).reset_index()

            # 비율 계산
            combined_summary["유효비율(%)"] = (combined_summary["유효생산량"] / combined_summary["총실적"] * 100).fillna(0)
            combined_summary["과생산비율(%)"] = (combined_summary["과생산량"] / combined_summary["총실적"] * 100).fillna(0)
            combined_summary["불필요비율(%)"] = (combined_summary["불필요생산량"] / combined_summary["총실적"] * 100).fillna(0)

            combined_summary["유효 대응률(수량)(%)"] = combined_summary["유효비율(%)"]

            # 선택지표 추가 (공장 비교 지표와 동일 3종)
            metric_map = {
                "정확 대응 비중": ("유효비율(%)", "유효생산량"),
                "초과 생산 비중": ("과생산비율(%)", "과생산량"),
                "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
            }
            metric_col, pcs_col = metric_map[combined_metric_option]
            combined_summary["선택지표"] = combined_summary[metric_col].fillna(0)

            # 테이블 표시
            base_cols = ["공장", "신규분류요약", "총실적"]
            display_combined = combined_summary[base_cols + [pcs_col, "선택지표"]].copy()
            total_hdr = f"{KPI_LABEL_MAP['총실적']} (pcs)"
            pcs_hdr = f"{KPI_LABEL_MAP[pcs_col]} (pcs)"
            rate_hdr = f"{combined_metric_option} (%)"
            display_combined.columns = ["공장", "신규분류요약", total_hdr, pcs_hdr, rate_hdr]

            # 공장 순서 지정 (A관 > C관 > S관)
            factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
            display_combined["_factory_sort"] = display_combined["공장"].map(factory_order)
            display_combined = display_combined.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True)
            display_combined = display_combined.drop("_factory_sort", axis=1)

            display_combined[total_hdr] = display_combined[total_hdr].map("{:,.0f}".format)
            display_combined[pcs_hdr] = display_combined[pcs_hdr].map("{:,.0f}".format)
            display_combined[rate_hdr] = display_combined[rate_hdr].map("{:.1f}%".format)

            html_parts = []
            # NOTE: Markdown에서는 4칸 이상 들여쓰기된 HTML이 코드블록으로 취급될 수 있어,
            # 모든 라인을 "맨 앞 공백 없이" 생성합니다.
            header_lines = [
                "<style>",
                ".custom-table { width: 100%; border-collapse: collapse; font-size: 14px; }",
                ".custom-table th, .custom-table td { padding: 10px 12px; border: 1px solid #e2e8f0; }",
                ".custom-table th { background: #f8fafc; color: #111827; text-align: left; }",
                ".custom-table td { vertical-align: middle; }",
                ".custom-table td.number { text-align: right; }",
                ".custom-table tbody tr:nth-child(even) { background: #f8fafc22; }",
                "</style>",
                "<table class=\"custom-table\">",
                "<thead>",
                "<tr>",
                "<th>공장</th>",
                "<th>신규분류요약</th>",
                f"<th>{total_hdr}</th>",
            ]
            header_lines.extend(
                [
                    f"<th>{pcs_hdr}</th>",
                    f"<th>{rate_hdr}</th>",
                    "</tr>",
                    "</thead>",
                    "<tbody>",
                ]
            )
            html_parts.append("\n".join(header_lines) + "\n")

            grouped = display_combined.groupby("공장", sort=False)
            for factory_name, group in grouped:
                rowspan = len(group)
                for idx, row in group.iterrows():
                    html_parts.append("<tr>")
                    if idx == group.index[0]:
                        html_parts.append(f"<td rowspan='{rowspan}' style='vertical-align: middle; font-weight: 600;'>{factory_name}</td>")
                    html_parts.append(f"<td>{row['신규분류요약']}</td>")
                    html_parts.append(f"<td class='number'>{row[total_hdr]}</td>")
                    html_parts.append(f"<td class='number'>{row[pcs_hdr]}</td>")
                    html_parts.append(f"<td class='number'>{row[rate_hdr]}</td>")
                    html_parts.append("</tr>")
            html_parts.append("</tbody></table>")
            st.markdown("".join(html_parts), unsafe_allow_html=True)

    # ============== 일별 요약 ==============
    st.markdown("### 📊 일별 요약")

    daily_display = daily_summary_filtered[
        [
            "날짜",
            "날짜_date",
            "총실적",
            "총부족수량",
            "유효생산량",
            "과생산량",
            "불필요생산량",
            "유효비율(%)",
            "과생산비율(%)",
            "불필요비율(%)",
        ]
    ].copy()

    # 일자별 규격 대응률(가능한 경우) 병합
    if shortage_prod_daily is not None and len(shortage_prod_daily) > 0:
        spec_rate = shortage_prod_daily[["날짜_date", "규격대응률(%)"]].copy()
        daily_display = daily_display.merge(spec_rate, on="날짜_date", how="left")
        daily_display["규격대응률(%)"] = daily_display["규격대응률(%)"].fillna(0)

    # 날짜는 일자까지만 표시 (시간 제거)
    daily_display["날짜"] = daily_display["날짜"].dt.strftime("%Y-%m-%d")

    # pcs 컬럼은 콤마 표시 및 컬럼명에 (pcs) 추가
    pcs_cols = ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]
    daily_display.rename(
        columns={c: f"{KPI_LABEL_MAP.get(c, c)} (pcs)" for c in pcs_cols},
        inplace=True,
    )
    if "규격대응률(%)" in daily_display.columns:
        daily_display.rename(columns={"규격대응률(%)": "규격 대응률(%)"}, inplace=True)
    daily_display.rename(columns=RATE_LABEL_MAP, inplace=True)

    # 컬럼 순서 정리 (비중(%) 우선, 수량(pcs)은 뒤쪽)
    daily_cols = [
        "날짜",
        f"{KPI_LABEL_MAP['총실적']} (pcs)",
        f"{KPI_LABEL_MAP['총부족수량']} (pcs)",
    ]
    if "규격 대응률(%)" in daily_display.columns:
        daily_cols.append("규격 대응률(%)")
    daily_cols.extend(
        [
            RATE_LABEL_MAP["유효비율(%)"],
            RATE_LABEL_MAP["과생산비율(%)"],
            RATE_LABEL_MAP["불필요비율(%)"],
            f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
            f"{KPI_LABEL_MAP['과생산량']} (pcs)",
            f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
        ]
    )
    daily_display = daily_display[daily_cols].copy()

    st.dataframe(
        daily_display.style.format(
            {
                **{f"{KPI_LABEL_MAP.get(c, c)} (pcs)": "{:,.0f}" for c in pcs_cols},
                **({"규격 대응률(%)": "{:.1f}%"} if "규격 대응률(%)" in daily_display.columns else {}),
                RATE_LABEL_MAP["유효비율(%)"]: "{:.1f}%",
                RATE_LABEL_MAP["과생산비율(%)"]: "{:.1f}%",
                RATE_LABEL_MAP["불필요비율(%)"]: "{:.1f}%",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("🔎 관별(공장별) 일별 상세 펼치기", expanded=False):
        if len(factory_summary_filtered) == 0:
            st.info("선택한 기간에 공장별 데이터가 없습니다.")
        else:
            factory_daily = factory_summary_filtered.groupby(["생산일자_date", "공장"], dropna=False).agg({
                "총실적": "sum",
                "총부족수량": "sum",
                "유효생산량": "sum",
                "과생산량": "sum",
                "불필요생산량": "sum",
            }).reset_index()

            factory_daily[RATE_LABEL_MAP["유효비율(%)"]] = (factory_daily["유효생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            factory_daily[RATE_LABEL_MAP["과생산비율(%)"]] = (factory_daily["과생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            factory_daily[RATE_LABEL_MAP["불필요비율(%)"]] = (factory_daily["불필요생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

            # 관별(공장별) 일자 규격 대응률(SKU 기준)
            factory_daily_spec = None
            factory_daily_spec_by_day = None
            if sku_daily_factory is not None and len(sku_daily_factory) > 0 and {"날짜_date", "공장", "규격대응률(%)"}.issubset(set(sku_daily_factory.columns)):
                factory_daily_spec = sku_daily_factory[
                    (sku_daily_factory["날짜_date"] >= start_date) &
                    (sku_daily_factory["날짜_date"] <= end_date) &
                    (sku_daily_factory["날짜_date"] != today)
                ][["날짜_date", "공장", "규격대응률(%)"]].copy()
                factory_daily_spec.rename(columns={"규격대응률(%)": "규격 대응률(%)"}, inplace=True)
            elif shortage_prod_daily is not None and len(shortage_prod_daily) > 0:
                factory_daily_spec_by_day = shortage_prod_daily[["날짜_date", "규격대응률(%)"]].copy()
                factory_daily_spec_by_day.rename(columns={"날짜_date": "날짜", "규격대응률(%)": "규격 대응률(%)"}, inplace=True)

            factory_need_label = f"{KPI_LABEL_MAP['총부족수량']} (pcs)"
            factory_daily_display = factory_daily.rename(columns={
                "생산일자_date": "날짜",
                "총실적": f"{KPI_LABEL_MAP['총실적']} (pcs)",
                "총부족수량": factory_need_label,
                "유효생산량": f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                "과생산량": f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                "불필요생산량": f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
            }).copy()

            if factory_daily_spec is not None and len(factory_daily_spec) > 0:
                factory_daily_display = factory_daily_display.merge(
                    factory_daily_spec.rename(columns={"날짜_date": "날짜"}),
                    on=["날짜", "공장"],
                    how="left",
                )
            elif factory_daily_spec_by_day is not None and len(factory_daily_spec_by_day) > 0:
                st.warning(
                    "관별(공장별) `규격 대응률(SKU 기준)` 계산 불가: 선택 기간 데이터에 `공장` 값이 비어있거나 SKU 집계에 필요한 컬럼이 누락되었습니다. "
                    "(전사 일자 규격 대응률을 동일 적용해 표시)"
                )
                factory_daily_display = factory_daily_display.merge(
                    factory_daily_spec_by_day,
                    on=["날짜"],
                    how="left",
                )

            if "규격 대응률(%)" in factory_daily_display.columns:
                factory_daily_display["규격 대응률(%)"] = (
                    pd.to_numeric(factory_daily_display["규격 대응률(%)"], errors="coerce")
                    .replace([np.inf, -np.inf], 0)
                )
            else:
                factory_daily_display["규격 대응률(%)"] = np.nan

            factory_daily_display["날짜"] = pd.to_datetime(factory_daily_display["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

            # 공장 순서 지정 (A관 > C관 > S관)
            factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
            factory_daily_display["_factory_sort"] = factory_daily_display["공장"].map(factory_order)
            factory_daily_display = factory_daily_display.sort_values(["날짜", "_factory_sort", "공장"]).drop(columns=["_factory_sort"]).reset_index(drop=True)

            # 컬럼 순서 정리 (비중(%) 우선, 수량(pcs)은 뒤쪽)
            factory_daily_cols = [
                "날짜",
                "공장",
                f"{KPI_LABEL_MAP['총실적']} (pcs)",
                factory_need_label,
                "규격 대응률(%)",
            ]
            factory_daily_cols.extend(
                [
                    RATE_LABEL_MAP["유효비율(%)"],
                    RATE_LABEL_MAP["과생산비율(%)"],
                    RATE_LABEL_MAP["불필요비율(%)"],
                    f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                    f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                    f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
                ]
            )
            factory_daily_display = factory_daily_display[factory_daily_cols].copy()

            st.dataframe(
                factory_daily_display.style.format({
                    f"{KPI_LABEL_MAP['총실적']} (pcs)": "{:,.0f}",
                    factory_need_label: "{:,.0f}",
                    f"{KPI_LABEL_MAP['유효생산량']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['과생산량']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['불필요생산량']} (pcs)": "{:,.0f}",
                    **({"규격 대응률(%)": "{:.1f}%"} if "규격 대응률(%)" in factory_daily_display.columns else {}),
                    RATE_LABEL_MAP["유효비율(%)"]: "{:.1f}%",
                    RATE_LABEL_MAP["과생산비율(%)"]: "{:.1f}%",
                    RATE_LABEL_MAP["불필요비율(%)"]: "{:.1f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )

except Exception as e:
    st.error(f"❌ 오류가 발생했습니다: {str(e)}")
    st.info("결과 파일을 다시 생성해주세요.")
