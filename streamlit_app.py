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
DASHBOARD_TITLE = "생산 운영 현황 대시보드"
KPI_LABEL_MAP = {
    "총실적": "총 생산량",
    "총부족수량": "미충족 수요",
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
    # NOTE: '총부족수량(=미충족 수요)'이 45일 수주 기준 스냅샷이라면 기간 합계(sum)는 중복 집계가 될 수 있어,
    # 선택 기간의 "마지막 날짜" 기준 스냅샷을 사용합니다.
    if len(daily_summary_filtered) > 0:
        daily_last = daily_summary_filtered.sort_values("날짜_date").iloc[-1]
        shortage_snapshot = int(daily_last["총부족수량"])
        shortage_snapshot_date = daily_last["날짜_date"]
    else:
        shortage_snapshot = 0
        shortage_snapshot_date = None

    prod_days = int(daily_summary_filtered["날짜_date"].nunique()) if len(daily_summary_filtered) > 0 else 0

    valid_rate = (valid_prod / total_prod * 100) if total_prod > 0 else 0
    over_rate = (over_prod / total_prod * 100) if total_prod > 0 else 0
    waste_rate = (waste_prod / total_prod * 100) if total_prod > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card(
            f"{KPI_LABEL_MAP['총실적']} (pcs)",
            f"{total_prod:,}",
            right_label="유효 대응률",
            right_value=valid_rate,
            right_color="#1d4ed8",
            sub=f"선택기간 `{prod_days}`일",
        )
    with col2:
        render_kpi_card(
            f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
            f"{valid_prod:,}",
            right_label=RATE_LABEL_MAP["유효비율(%)"].replace("(%)", ""),
            right_value=valid_rate,
            right_color="#047857",
            sub=f"{KPI_LABEL_MAP['총실적']} 대비",
        )
    with col3:
        render_kpi_card(
            f"{KPI_LABEL_MAP['과생산량']} (pcs)",
            f"{over_prod:,}",
            right_label=RATE_LABEL_MAP["과생산비율(%)"].replace("(%)", ""),
            right_value=over_rate,
            right_color="#b91c1c",
            sub=f"{KPI_LABEL_MAP['총실적']} 대비",
        )
    with col4:
        render_kpi_card(
            f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
            f"{waste_prod:,}",
            right_label=RATE_LABEL_MAP["불필요비율(%)"].replace("(%)", ""),
            right_value=waste_rate,
            right_color="#b45309",
            sub=f"{KPI_LABEL_MAP['총실적']} 대비",
        )

    with st.expander("지표 정의/상세 보기", expanded=False):
        st.markdown(
            "- `필요 수량(SKU)` : 해당 기간 실제 필요한 SKU 수량(불필요 SKU 제외)\n"
            "- `정확 대응 생산량` : SKU별 `min(생산, 필요)`의 합\n"
            "- `초과 생산량` : SKU별 `max(생산-필요, 0)`의 합\n"
            "- `비정형 생산량` : 필요 SKU 외 생산(필요=0인데 생산>0)\n"
            "- `유효 대응률` = `정확 대응 생산량` ÷ `총 생산량`\n"
            "- 참고: `미충족 수요`는 45일 수주 기준 스냅샷 값으로 운영 참고용입니다."
        )
        st.write(f"- 선택기간 종료일 `미충족 수요`(스냅샷): `{shortage_snapshot:,}` pcs")
        if shortage_snapshot_date is not None:
            st.write(f"- 미충족 수요 기준일: `{shortage_snapshot_date}`")

    st.markdown("<div style='margin-top:50px'></div>", unsafe_allow_html=True)
    # ============== 중간: 차트 ==============
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
        factory_data["유효 대응률(%)"] = factory_data["유효비율(%)"]

        metric_option = st.radio(
            "공장 비교 지표",
            ["유효 대응률", "정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"],
            horizontal=True,
        )
        metric_desc = {
            "유효 대응률": "총 생산 중 ‘정확 대응 생산량’ 비율(=정확 대응 비중)",
            "정확 대응 비중": "총 생산량 중 정확 대응 생산량이 차지하는 비중",
            "초과 생산 비중": "총 생산량 중 초과 생산량이 차지하는 비중",
            "비정형 생산 비중": "총 생산량 중 비정형 생산량이 차지하는 비중",
        }
        st.caption(f"설명: {metric_desc[metric_option]}")

        metric_map = {
            "유효 대응률": ("유효 대응률(%)", "유효생산량"),
            "정확 대응 비중": ("유효비율(%)", "유효생산량"),
            "초과 생산 비중": ("과생산비율(%)", "과생산량"),
            "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
        }
        metric_col, pcs_col = metric_map[metric_option]
        factory_data["선택지표"] = factory_data[metric_col].replace([np.inf, -np.inf], 0).fillna(0)

        fig = px.bar(
            factory_data,
            x="공장",
            y="선택지표",
            color="공장",
            title=f"공장별 {metric_option} (%)",
            text="선택지표",
            hover_data={
                "총실적": ":,",
                "유효생산량": ":,",
                "과생산량": ":,",
                "불필요생산량": ":,",
                "유효 대응률(%)": ":.1f",
                "유효비율(%)": ":.1f",
                "과생산비율(%)": ":.1f",
                "불필요비율(%)": ":.1f",
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
        st.caption("Tip: ‘유효 대응률’은 필요 수량 대비 생산 정확도를 반영한 지표로, backlog 크기 영향 없이 공장 비교가 가능합니다.")

        # 공장_신규분류별 통합 현황
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

        combined_summary["유효 대응률(%)"] = combined_summary["유효비율(%)"]

        # 선택지표 추가
        metric_map = {
            "유효 대응률": ("유효 대응률(%)", "유효생산량"),
            "정확 대응 비중": ("유효비율(%)", "유효생산량"),
            "초과 생산 비중": ("과생산비율(%)", "과생산량"),
            "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
        }
        metric_col, pcs_col = metric_map[metric_option]
        combined_summary["선택지표"] = combined_summary[metric_col].fillna(0)

        # 테이블 표시
        base_cols = ["공장", "신규분류요약", "총실적"]
        display_combined = combined_summary[base_cols + [pcs_col, "선택지표"]].copy()
        total_hdr = f"{KPI_LABEL_MAP['총실적']} (pcs)"
        pcs_hdr = f"{KPI_LABEL_MAP[pcs_col]} (pcs)"
        rate_hdr = f"{metric_option} (%)"
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
        ["날짜", "총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량", "유효비율(%)", "과생산비율(%)", "불필요비율(%)"]
    ].copy()
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
            RATE_LABEL_MAP["과생산비율(%)"]: "{:.1f}%",
            RATE_LABEL_MAP["불필요비율(%)"]: "{:.1f}%",
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

            factory_daily[RATE_LABEL_MAP["유효비율(%)"]] = (factory_daily["유효생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            factory_daily[RATE_LABEL_MAP["과생산비율(%)"]] = (factory_daily["과생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            factory_daily[RATE_LABEL_MAP["불필요비율(%)"]] = (factory_daily["불필요생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

            factory_daily_display = factory_daily.rename(columns={
                "생산일자_date": "날짜",
                "총실적": f"{KPI_LABEL_MAP['총실적']} (pcs)",
                "총부족수량": f"{KPI_LABEL_MAP['총부족수량']} (pcs)",
                "유효생산량": f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                "과생산량": f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                "불필요생산량": f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
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
