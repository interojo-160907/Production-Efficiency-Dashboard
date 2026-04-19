import os
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# 페이지 설정
DASHBOARD_TITLE = "수요 대응 생산량 대시보드"
KPI_LABEL_MAP = {
    "총실적": "총 생산량",
    "총부족수량": "총 부족수량",
    "유효생산량": "수요 대응 생산량",
    "과생산량": "선행 확보 생산량",
    "불필요생산량": "비계획 생산량",
}
RATE_LABEL_MAP = {
    "유효비율(%)": "수요대응율(%)",
    "과생산비율(%)": "선행확보율(%)",
    "불필요비율(%)": "비계획율(%)",
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
        padding: 18px 20px;
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
        font-size: 34px;
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

@st.cache_data(show_spinner=False)
def load_result_excel(result_path: Path, mtime_ns: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """결과 엑셀 파일 로드 + 전처리 (Streamlit 캐시 적용)

    - 위젯 변경 시 전체 스크립트가 rerun 되므로, 엑셀 로딩/전처리를 캐시해 속도를 개선합니다.
    - mtime_ns는 파일 변경 시 캐시 무효화를 위해 사용됩니다.
    """
    _ = mtime_ns  # cache key only
    sheets = pd.read_excel(result_path, sheet_name=["일별요약", "공장_신규분류별"])

    required = {"일별요약", "공장_신규분류별"}
    missing = required - set(sheets.keys())
    if missing:
        raise ValueError(f"결과 엑셀에 필요한 시트가 없습니다: {', '.join(missing)}")

    daily_summary = sheets["일별요약"]
    factory_summary = sheets["공장_신규분류별"]

    daily_summary["날짜"] = pd.to_datetime(daily_summary["날짜"], errors="coerce")
    factory_summary["생산일자"] = pd.to_datetime(factory_summary["생산일자"], errors="coerce")

    daily_summary["날짜_date"] = daily_summary["날짜"].dt.date
    factory_summary["생산일자_date"] = factory_summary["생산일자"].dt.date

    return daily_summary, factory_summary


# 결과 파일 경로
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
result_path = Path(BASE_PATH) / "유효생산량_결과.xlsx"

if not result_path.exists():
    st.error(f"⚠️ 결과 파일을 찾을 수 없습니다: {result_path}")
    st.info("먼저 `aps_yield_dashboard.py`를 실행해서 결과 파일을 생성하세요.")
    st.stop()

try:
    daily_summary, factory_summary = load_result_excel(result_path, result_path.stat().st_mtime_ns)

    # 금일 데이터 제외 (아직 생산 중이므로) - KST 기준
    now_kst = datetime.now(ZoneInfo("Asia/Seoul"))
    today = now_kst.date()
    st.caption(f"기준 시각(KST): {now_kst.strftime('%Y-%m-%d %H:%M:%S')}")

    # 제목
    st.markdown(f"<h1 style='text-align:center; color:#1f3a93; margin:0;'>🏭 {DASHBOARD_TITLE}</h1>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # 기간 필터
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

    # 날짜 범위 결정
    if filter_option == "당월":
        start_date = current_month_start
        end_date = current_month_end
    elif filter_option == "전월":
        start_date = prev_month_start
        end_date = last_day_prev
    else:  # 기간조회
        min_date = daily_summary[daily_summary["날짜_date"] != today]["날짜"].min()
        max_date = daily_summary[daily_summary["날짜_date"] != today]["날짜"].max()

        col_filter1, col_space, col_filter2 = st.columns([1.5, 0.2, 1.5])

        with col_filter1:
            start_date = st.date_input("시작 날짜", value=max_date, min_value=min_date, max_value=max_date)

        with col_filter2:
            end_date = st.date_input("종료 날짜", value=max_date, min_value=min_date, max_value=max_date)

    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)

    # 필터 적용 (기간 범위)
    daily_summary_filtered = daily_summary[
        (daily_summary["날짜_date"] >= start_date) &
        (daily_summary["날짜_date"] <= end_date) &
        (daily_summary["날짜_date"] != today)
    ]

    factory_summary_filtered = factory_summary[
        (factory_summary["생산일자_date"] >= start_date) &
        (factory_summary["생산일자_date"] <= end_date) &
        (factory_summary["생산일자_date"] != today)
    ]

    # 메트릭 계산
    total_prod = int(daily_summary_filtered["총실적"].sum()) if len(daily_summary_filtered) > 0 else 0
    valid_prod = int(daily_summary_filtered["유효생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    over_prod = int(daily_summary_filtered["과생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    waste_prod = int(daily_summary_filtered["불필요생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    # NOTE: '총부족수량'이 45일 수주 기준 부족(스냅샷)이라면 기간 합계(sum)는 중복 집계가 될 수 있어,
    # 선택 기간의 "마지막 날짜" 기준 스냅샷을 사용합니다.
    if len(daily_summary_filtered) > 0:
        daily_last = daily_summary_filtered.sort_values("날짜_date").iloc[-1]
        shortage_snapshot = int(daily_last["총부족수량"])
        shortage_snapshot_date = daily_last["날짜_date"]
    else:
        shortage_snapshot = 0
        shortage_snapshot_date = None

    # 부족(백로그) 흐름(추정)
    # - 총부족수량은 "전체 수주(45일 기준) 대비 잔여 부족"을 의미하는 스냅샷 값
    # - 생산(유효생산량)으로 부족이 채워지고, 동시에 신규 수주/계획 변경으로 부족이 늘거나 줄 수 있음
    daily_flow = None
    shortage_start_snapshot = 0
    shortage_start_snapshot_date = None
    shortage_start_snapshot_is_estimated = False
    new_shortage_est = 0
    adjust_est = 0
    shortage_target_period = 0
    if len(daily_summary_filtered) > 0:
        daily_sorted = daily_summary_filtered.sort_values("날짜_date").copy()

        # 기간 시작 백로그(스냅샷): 가능하면 start_date 이전 마지막 스냅샷을 사용
        daily_before = daily_summary[
            (daily_summary["날짜_date"] < start_date) &
            (daily_summary["날짜_date"] != today)
        ]
        if len(daily_before) > 0:
            prev_last = daily_before.sort_values("날짜_date").iloc[-1]
            shortage_start_snapshot = int(prev_last["총부족수량"])
            shortage_start_snapshot_date = prev_last["날짜_date"]
        else:
            # 이전 스냅샷이 없으면 "기간 첫날 잔여부족 + 첫날 유효생산"으로 시작 백로그를 보수적으로 추정
            first = daily_sorted.iloc[0]
            shortage_start_snapshot = int(first["총부족수량"] + first["유효생산량"])
            shortage_start_snapshot_date = first["날짜_date"]
            shortage_start_snapshot_is_estimated = True

        daily_flow = daily_sorted[["날짜_date", "총부족수량", "유효생산량"]].copy()
        daily_flow.rename(columns={"총부족수량": "잔여부족(스냅샷)"}, inplace=True)
        daily_flow["잔여부족_prev"] = daily_flow["잔여부족(스냅샷)"].shift(1)
        daily_flow.loc[daily_flow.index[0], "잔여부족_prev"] = shortage_start_snapshot

        daily_flow["잔여부족증감"] = daily_flow["잔여부족(스냅샷)"] - daily_flow["잔여부족_prev"]
        # 잔여부족(t) = 잔여부족(t-1) + 신규부족(t) - 유효생산량(t)  (유효생산량이 전부 부족 해소에 투입된다고 가정)
        daily_flow["추정신규부족"] = daily_flow["잔여부족증감"] + daily_flow["유효생산량"]
        daily_flow["추정신규부족(+)"] = daily_flow["추정신규부족"].clip(lower=0)
        # 추정신규부족이 음수면 수주취소/납기변경/계획조정 등으로 부족 자체가 줄어든 케이스로 해석
        daily_flow["추정조정(-)"] = (-daily_flow["추정신규부족"].clip(upper=0))

        new_shortage_est = int(daily_flow["추정신규부족(+)"].sum())
        adjust_est = int(daily_flow["추정조정(-)"].sum())
        shortage_target_period = int(shortage_start_snapshot + new_shortage_est)

    # KPI용 "기간 부족 타겟"(초기 백로그 + 기간 중 신규부족 추정)
    demand_total = shortage_target_period if shortage_target_period > 0 else shortage_snapshot
    prod_days = int(daily_summary_filtered["날짜_date"].nunique()) if len(daily_summary_filtered) > 0 else 0
    avg_valid_per_day = (valid_prod / prod_days) if prod_days > 0 else 0
    backlog_days = (shortage_snapshot / avg_valid_per_day) if avg_valid_per_day > 0 else None

    valid_rate = (valid_prod / total_prod * 100) if total_prod > 0 else 0
    over_rate = (over_prod / total_prod * 100) if total_prod > 0 else 0
    waste_rate = (waste_prod / total_prod * 100) if total_prod > 0 else 0
    fulfillment_rate = (valid_prod / demand_total * 100) if demand_total > 0 else 0
    fulfillment_rate = min(fulfillment_rate, 100.0)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card(
            f"{KPI_LABEL_MAP['총실적']} (pcs)",
            f"{total_prod:,}",
            right_label="충족률(기간부족대비)",
            right_value=fulfillment_rate,
            right_color="#1d4ed8",
            sub="*초기 부족(스냅샷)+기간 신규부족(추정) 대비",
        )
    with col2:
        render_kpi_card(
            f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
            f"{valid_prod:,}",
            right_label="대응율",
            right_value=valid_rate,
            right_color="#047857",
            sub="총실적 대비",
        )
    with col3:
        render_kpi_card(
            f"{KPI_LABEL_MAP['과생산량']} (pcs)",
            f"{over_prod:,}",
            right_label="선행확보율",
            right_value=over_rate,
            right_color="#b91c1c",
            sub="총실적 대비",
        )
    with col4:
        render_kpi_card(
            f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
            f"{waste_prod:,}",
            right_label="비계획율",
            right_value=waste_rate,
            right_color="#b45309",
            sub="총실적 대비",
        )

    with st.expander("지표 정의/상세 보기", expanded=False):
        st.markdown(
            "- `총부족수량` = **(45일 수주 기준) 전체 수주 대비 잔여 부족수량(스냅샷)**\n"
            "- `유효생산량` = 부족 해소에 기여한 생산량(수요대응)\n"
            "- `충족률(기간부족대비)` = 기간 유효생산량 ÷ (기간 시작 부족(스냅샷) + 기간 신규부족(추정))\n"
            "- `대응율`/`선행확보율`/`비계획율` = 각 생산량 ÷ 총실적\n"
            "- *`총부족수량`은 스냅샷 값이라 기간 합계(sum)로 더하면 중복될 수 있어, 선택기간 종료일 스냅샷을 사용합니다.*"
        )
        st.write(f"- 선택기간 종료일 잔여부족(스냅샷): `{shortage_snapshot:,}` pcs")
        if shortage_snapshot_date is not None:
            st.write(f"- 부족 스냅샷 기준일: `{shortage_snapshot_date}`")
        if shortage_start_snapshot_date is not None:
            start_label = "기간 시작 부족(스냅샷)"
            if shortage_start_snapshot_is_estimated:
                start_label += " (추정)"
            st.write(f"- {start_label}: `{shortage_start_snapshot:,}` pcs (기준일: `{shortage_start_snapshot_date}`)")
        if daily_flow is not None and len(daily_flow) > 0:
            st.write(f"- 기간 신규부족(추정 +): `{new_shortage_est:,}` pcs")
            if adjust_est > 0:
                st.write(f"- 수주/계획 조정(추정 -): `{adjust_est:,}` pcs")
            st.write(f"- 기간 부족 타겟(초기+신규): `{shortage_target_period:,}` pcs")
        st.write(f"- 부족수량 (생산 타겟) = `{demand_total:,}` pcs")
        if backlog_days is not None:
            st.write(f"- 백로그 해소 추정: `{backlog_days:.1f}` 일 (일평균 수요대응 `{avg_valid_per_day:,.0f}` pcs/일, 선택기간 `{prod_days}`일)")

    st.markdown("<div style='margin-top:50px'></div>", unsafe_allow_html=True)
    # ============== 중간: 차트 ==============
    st.markdown("### 📈 공장별 현황")

    if len(factory_summary_filtered) == 0:
        st.info("선택한 기간에 공장별 데이터가 없습니다.")
    else:
        # 공장별 데이터 준비
        factory_data = factory_summary_filtered.groupby("공장", dropna=False).agg({
            "총실적": "sum",
            "총부족수량": "sum",
            "유효생산량": "sum",
            "과생산량": "sum",
            "불필요생산량": "sum"
        }).reset_index()

        factory_data["유효비율(%)"] = (factory_data["유효생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["과생산비율(%)"] = (factory_data["과생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["불필요비율(%)"] = (factory_data["불필요생산량"] / factory_data["총실적"] * 100).fillna(0)

        metric_option = st.radio(
            "표시할 지표를 선택하세요",
            ["수요충족률(부족대비)", "수요대응율(총실적대비)", "선행확보율", "비계획율"],
            horizontal=True
        )
        metric_desc = {
            "수요대응율(총실적대비)": "총 생산량(실적) 중 수요 대응 생산량이 차지하는 비중",
            "수요충족률(부족대비)": "부족수량 (생산 타겟) 대비 수요 대응 생산량 비율",
            "선행확보율": "향후 대응을 위한 선제 생산 비율",
            "비계획율": "수요 기준에 포함되지 않은 생산 비율",
        }
        st.caption(f"설명: {metric_desc[metric_option]}")

        # 공장별 추가 지표 (필요대비)
        # NOTE: 기간 내 '총부족수량' 합계는 중복 집계가 될 수 있어, 공장별 최신일(스냅샷) 기준으로 사용합니다.
        factory_short_snapshot = factory_summary_filtered.dropna(subset=["생산일자_date"]).copy()
        if len(factory_short_snapshot) > 0:
            idx = factory_short_snapshot.groupby("공장")["생산일자_date"].idxmax()
            short_snap = factory_short_snapshot.loc[idx, ["공장", "총부족수량", "생산일자_date"]].rename(
                columns={"총부족수량": "부족수량(스냅샷)", "생산일자_date": "부족기준일"}
            )
        else:
            short_snap = pd.DataFrame(columns=["공장", "부족수량(스냅샷)", "부족기준일"])
        factory_data = factory_data.merge(short_snap, on="공장", how="left")
        factory_data["부족수량(스냅샷)"] = factory_data["부족수량(스냅샷)"].fillna(0)

        factory_data["부족수량 (생산 타겟)"] = factory_data["부족수량(스냅샷)"].fillna(0)
        
        # 일별 수요충족률의 평균으로 계산 (기간 합계의 왜곡 방지)
        tmp_daily = factory_summary_filtered.groupby(["생산일자_date", "공장"])[["유효생산량", "총부족수량"]].sum().reset_index()
        tmp_daily["일별충족률"] = (tmp_daily["유효생산량"] / tmp_daily["총부족수량"] * 100).replace([np.inf, -np.inf], 0).fillna(0).clip(upper=100.0)
        avg_fulfillment = tmp_daily.groupby("공장")["일별충족률"].mean().reset_index().rename(columns={"일별충족률": "수요충족률(%)"})
        
        factory_data = factory_data.merge(avg_fulfillment, on="공장", how="left")
        factory_data["수요충족률(%)"] = factory_data["수요충족률(%)"].fillna(0)
        # 백로그 해소 추정(일) = 부족(스냅샷) / (선택기간 일평균 수요대응 생산량)
        prod_days_factory = int(factory_summary_filtered["생산일자_date"].nunique()) if len(factory_summary_filtered) > 0 else 0
        factory_data["일평균수요대응(pcs/일)"] = (factory_data["유효생산량"] / prod_days_factory).fillna(0) if prod_days_factory > 0 else 0
        factory_data["백로그해소추정(일)"] = (
            factory_data["부족수량(스냅샷)"] / factory_data["일평균수요대응(pcs/일)"]
        ).where(factory_data["일평균수요대응(pcs/일)"] > 0)

        metric_map = {
            "수요대응율(총실적대비)": ("유효비율(%)", "유효생산량"),
            "수요충족률(부족대비)": ("수요충족률(%)", "유효생산량"),
            "선행확보율": ("과생산비율(%)", "과생산량"),
            "비계획율": ("불필요비율(%)", "불필요생산량"),
        }
        metric_col, pcs_col = metric_map[metric_option]
        factory_data["선택지표"] = factory_data[metric_col]

        fig = px.bar(
            factory_data,
            x="공장",
            y="선택지표",
            color="공장",
            title=f"공장별 {metric_option} (%)",
            text="선택지표",
            hover_data={
                "총실적": ":,",
                "부족수량 (생산 타겟)": ":,",
                "유효생산량": ":,",
                "과생산량": ":,",
                "불필요생산량": ":,",
                "수요충족률(%)": ":.1f",
                "일평균수요대응(pcs/일)": ":,.0f",
                "백로그해소추정(일)": ":.1f",
                "선택지표": ":.1f",
            },
        )
        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
            textfont=dict(size=24, family="Arial", color="#222222"),
            marker=dict(cornerradius="15")
        )
        fig.update_layout(
            height=520,
            showlegend=False,
            margin=dict(l=0, r=0, t=60, b=0),
            yaxis=dict(range=[0, 100], title=dict(text=f"{metric_option} (%)", font=dict(size=16, family="Arial", color="#222222"))),
            xaxis=dict(
                title=dict(text="공장", font=dict(size=16, family="Arial", color="#222222")),
                tickfont=dict(size=18, family="Arial", color="#222222")
            ),
            title=dict(font=dict(size=22, family="Arial", color="#111111"))
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown(f"**선택 지표: {metric_option} (%)**")
        st.caption("Tip: 부족수량(생산 타겟)은 '45일 수주 기준 부족수량(스냅샷)'을 의미하므로, '수요대응율(총실적대비)'만 보지 말고 '수요충족률(부족대비)'도 같이 확인하는 것이 안전합니다.")

        # 공장_신규분류별 통합 현황
        combined_summary = factory_summary_filtered.groupby(["공장", "신규분류요약"], dropna=False).agg({
            "총실적": "sum",
            "총부족수량": "sum",
            "유효생산량": "sum",
            "과생산량": "sum",
            "불필요생산량": "sum"
        }).reset_index()

        # 분류별 부족수량도 최신일(스냅샷) 기준으로 별도 계산
        combined_short_snapshot = factory_summary_filtered.dropna(subset=["생산일자_date"]).copy()
        if len(combined_short_snapshot) > 0:
            idx = combined_short_snapshot.groupby(["공장", "신규분류요약"])["생산일자_date"].idxmax()
            short_snap = combined_short_snapshot.loc[idx, ["공장", "신규분류요약", "총부족수량", "생산일자_date"]].rename(
                columns={"총부족수량": "부족수량(스냅샷)", "생산일자_date": "부족기준일"}
            )
        else:
            short_snap = pd.DataFrame(columns=["공장", "신규분류요약", "부족수량(스냅샷)", "부족기준일"])
        combined_summary = combined_summary.merge(short_snap, on=["공장", "신규분류요약"], how="left")
        combined_summary["부족수량(스냅샷)"] = combined_summary["부족수량(스냅샷)"].fillna(0)

        # 비율 계산
        combined_summary["유효비율(%)"] = (combined_summary["유효생산량"] / combined_summary["총실적"] * 100).fillna(0)
        combined_summary["과생산비율(%)"] = (combined_summary["과생산량"] / combined_summary["총실적"] * 100).fillna(0)
        combined_summary["불필요비율(%)"] = (combined_summary["불필요생산량"] / combined_summary["총실적"] * 100).fillna(0)
        combined_summary["부족수량 (생산 타겟)"] = combined_summary["부족수량(스냅샷)"].fillna(0)
        
        # 일별 수요충족률의 평균으로 계산
        tmp_daily_comb = factory_summary_filtered.groupby(["생산일자_date", "공장", "신규분류요약"])[["유효생산량", "총부족수량"]].sum().reset_index()
        tmp_daily_comb["일별충족률"] = (tmp_daily_comb["유효생산량"] / tmp_daily_comb["총부족수량"] * 100).replace([np.inf, -np.inf], 0).fillna(0).clip(upper=100.0)
        avg_fulfillment_comb = tmp_daily_comb.groupby(["공장", "신규분류요약"])["일별충족률"].mean().reset_index().rename(columns={"일별충족률": "수요충족률(%)"})
        
        combined_summary = combined_summary.merge(avg_fulfillment_comb, on=["공장", "신규분류요약"], how="left")
        combined_summary["수요충족률(%)"] = combined_summary["수요충족률(%)"].fillna(0)

        # 선택지표 추가
        metric_map = {
            "수요대응율(총실적대비)": ("유효비율(%)", "유효생산량"),
            "수요충족률(부족대비)": ("수요충족률(%)", "유효생산량"),
            "선행확보율": ("과생산비율(%)", "과생산량"),
            "비계획율": ("불필요비율(%)", "불필요생산량"),
        }
        metric_col, pcs_col = metric_map[metric_option]
        combined_summary["선택지표"] = combined_summary[metric_col]

        # 테이블 표시
        base_cols = ["공장", "신규분류요약", "총실적"]
        if metric_option == "수요충족률(부족대비)":
            display_combined = combined_summary[base_cols + ["부족수량 (생산 타겟)", pcs_col, "선택지표"]].copy()
        else:
            display_combined = combined_summary[base_cols + [pcs_col, "선택지표"]].copy()
        total_hdr = f"{KPI_LABEL_MAP['총실적']} (pcs)"
        demand_hdr = "부족수량 (생산 타겟) (pcs)"
        pcs_hdr = f"{KPI_LABEL_MAP[pcs_col]} (pcs)"
        rate_hdr = f"{metric_option} (%)"
        if metric_option == "수요충족률(부족대비)":
            display_combined.columns = ["공장", "신규분류요약", total_hdr, demand_hdr, pcs_hdr, rate_hdr]
        else:
            display_combined.columns = ["공장", "신규분류요약", total_hdr, pcs_hdr, rate_hdr]

        # 공장 순서 지정 (A관 > C관 > S관)
        factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
        display_combined["_factory_sort"] = display_combined["공장"].map(factory_order)
        display_combined = display_combined.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True)
        display_combined = display_combined.drop("_factory_sort", axis=1)

        display_combined[total_hdr] = display_combined[total_hdr].map("{:,.0f}".format)
        if metric_option == "수요충족률(부족대비)":
            display_combined[demand_hdr] = display_combined[demand_hdr].map("{:,.0f}".format)
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
        if metric_option == "수요충족률(부족대비)":
            header_lines.append(f"<th>{demand_hdr}</th>")
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
                if metric_option == "수요충족률(부족대비)":
                    html_parts.append(f"<td class='number'>{row[demand_hdr]}</td>")
                html_parts.append(f"<td class='number'>{row[pcs_hdr]}</td>")
                html_parts.append(f"<td class='number'>{row[rate_hdr]}</td>")
                html_parts.append("</tr>")
        html_parts.append("</tbody></table>")
        st.markdown("".join(html_parts), unsafe_allow_html=True)

    # ============== 일별 요약 ==============
    st.markdown("### 📊 일별 요약")

    if daily_flow is not None and len(daily_flow) > 0:
        flow_chart = daily_flow.copy()
        flow_chart["날짜"] = pd.to_datetime(flow_chart["날짜_date"], errors="coerce").dt.strftime("%Y-%m-%d")

        fig_flow = go.Figure()
        fig_flow.add_trace(
            go.Bar(
                x=flow_chart["날짜"],
                y=flow_chart["유효생산량"],
                name="유효생산량(부족해소)",
                marker_color="#047857",
                hovertemplate="%{x}<br>유효생산: %{y:,} pcs<extra></extra>",
            )
        )
        fig_flow.add_trace(
            go.Bar(
                x=flow_chart["날짜"],
                y=flow_chart["추정신규부족(+)"],
                name="추정 신규부족(+)",
                marker_color="#b45309",
                hovertemplate="%{x}<br>신규부족(추정): %{y:,} pcs<extra></extra>",
            )
        )
        fig_flow.add_trace(
            go.Scatter(
                x=flow_chart["날짜"],
                y=flow_chart["잔여부족(스냅샷)"],
                name="잔여부족(스냅샷)",
                mode="lines+markers",
                line=dict(color="#1d4ed8", width=2),
                yaxis="y2",
                hovertemplate="%{x}<br>잔여부족: %{y:,} pcs<extra></extra>",
            )
        )
        fig_flow.update_layout(
            title="부족 흐름(추정): 생산으로 채우고, 신규 수주/조정으로 변동",
            barmode="group",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            margin=dict(l=10, r=10, t=60, b=10),
            yaxis=dict(title="pcs (생산/신규부족)"),
            yaxis2=dict(
                title="pcs (잔여부족 스냅샷)",
                overlaying="y",
                side="right",
                showgrid=False,
            ),
        )
        st.plotly_chart(fig_flow, use_container_width=True)

        with st.expander("부족 흐름 수치(추정) 보기", expanded=False):
            flow_table = flow_chart[
                ["날짜", "잔여부족(스냅샷)", "유효생산량", "추정신규부족(+)", "추정조정(-)", "잔여부족증감"]
            ].copy()
            flow_table.rename(
                columns={
                    "잔여부족(스냅샷)": "잔여부족(스냅샷)(pcs)",
                    "유효생산량": "유효생산량(pcs)",
                    "추정신규부족(+)": "추정신규부족(+)(pcs)",
                    "추정조정(-)": "추정조정(-)(pcs)",
                    "잔여부족증감": "잔여부족증감(pcs)",
                },
                inplace=True,
            )
            st.dataframe(
                flow_table.style.format(
                    {
                        "잔여부족(스냅샷)(pcs)": "{:,.0f}",
                        "유효생산량(pcs)": "{:,.0f}",
                        "추정신규부족(+)(pcs)": "{:,.0f}",
                        "추정조정(-)(pcs)": "{:,.0f}",
                        "잔여부족증감(pcs)": "{:,.0f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    daily_display = daily_summary_filtered[["날짜", "총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량", "유효비율(%)"]].copy()
    # 날짜는 일자까지만 표시 (시간 제거)
    daily_display["날짜"] = daily_display["날짜"].dt.strftime("%Y-%m-%d")
    # pcs 컬럼은 콤마 표시 및 컬럼명에 (pcs) 추가
    pcs_cols = ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]
    daily_display.rename(
        columns={c: f"{KPI_LABEL_MAP.get(c, c)} (pcs)" for c in pcs_cols},
        inplace=True,
    )
    daily_display.rename(columns=RATE_LABEL_MAP, inplace=True)

    st.dataframe(
        daily_display.style.format({
            **{f"{KPI_LABEL_MAP.get(c, c)} (pcs)": "{:,.0f}" for c in pcs_cols},
            RATE_LABEL_MAP["유효비율(%)"]: "{:.1f}%",
        }),
        use_container_width=True,
        hide_index=True
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

            factory_daily["수요대응율(총실적대비)(%)"] = (factory_daily["유효생산량"] / factory_daily["총실적"] * 100).fillna(0)
            factory_daily["선행확보율(%)"] = (factory_daily["과생산량"] / factory_daily["총실적"] * 100).fillna(0)
            factory_daily["비계획율(%)"] = (factory_daily["불필요생산량"] / factory_daily["총실적"] * 100).fillna(0)
            factory_daily["부족수량 (생산 타겟)"] = factory_daily["총부족수량"].fillna(0)
            factory_daily["수요충족률(부족대비)(%)"] = (factory_daily["유효생산량"] / factory_daily["부족수량 (생산 타겟)"] * 100).replace([np.inf, -np.inf], 0).fillna(0).clip(upper=100.0)

            factory_daily_display = factory_daily.rename(columns={
                "생산일자_date": "날짜",
                "총실적": f"{KPI_LABEL_MAP['총실적']} (pcs)",
                "총부족수량": f"{KPI_LABEL_MAP['총부족수량']} (pcs)",
                "유효생산량": f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                "과생산량": f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                "불필요생산량": f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
                "부족수량 (생산 타겟)": "부족수량 (생산 타겟) (pcs)",
            }).copy()

            factory_daily_display["날짜"] = pd.to_datetime(factory_daily_display["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

            # 공장 순서 지정 (A관 > C관 > S관)
            factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
            factory_daily_display["_factory_sort"] = factory_daily_display["공장"].map(factory_order)
            factory_daily_display = factory_daily_display.sort_values(["날짜", "_factory_sort", "공장"]).drop(columns=["_factory_sort"]).reset_index(drop=True)

            st.dataframe(
                factory_daily_display.style.format({
                    f"{KPI_LABEL_MAP['총실적']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['총부족수량']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['유효생산량']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['과생산량']} (pcs)": "{:,.0f}",
                    f"{KPI_LABEL_MAP['불필요생산량']} (pcs)": "{:,.0f}",
                    "부족수량 (생산 타겟) (pcs)": "{:,.0f}",
                    "수요대응율(총실적대비)(%)": "{:.1f}%",
                    "수요충족률(부족대비)(%)": "{:.1f}%",
                    "선행확보율(%)": "{:.1f}%",
                    "비계획율(%)": "{:.1f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )

except Exception as e:
    st.error(f"❌ 오류가 발생했습니다: {str(e)}")
    st.info("결과 파일을 다시 생성해주세요.")
