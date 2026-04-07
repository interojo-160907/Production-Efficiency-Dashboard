import os
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# 페이지 설정
st.set_page_config(page_title="APS 유효생산량 대시보드", layout="wide", initial_sidebar_state="collapsed")

# CSS 스타일링
st.markdown("""
<style>
    [data-testid="metric.container"] {
        background-color: #f0f4f8;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
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

def load_result_excel(result_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """결과 엑셀 파일 로드"""
    if not result_path.exists():
        raise FileNotFoundError(f"결과 엑셀 파일을 찾을 수 없습니다: {result_path}")

    sheets = pd.read_excel(result_path, sheet_name=None)
    required = {"매칭결과", "일별요약", "공장_신규분류별"}
    missing = required - set(sheets.keys())
    if missing:
        raise ValueError(f"결과 엑셀에 필요한 시트가 없습니다: {', '.join(missing)}")

    return sheets["매칭결과"], sheets["일별요약"], sheets["공장_신규분류별"]


# 결과 파일 경로
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
result_path = Path(BASE_PATH) / "유효생산량_결과.xlsx"

if not result_path.exists():
    st.error(f"⚠️ 결과 파일을 찾을 수 없습니다: {result_path}")
    st.info("먼저 `aps_yield_dashboard.py`를 실행해서 결과 파일을 생성하세요.")
    st.stop()

try:
    result, daily_summary, factory_summary = load_result_excel(result_path)
    
    # 날짜 변환
    daily_summary["날짜"] = pd.to_datetime(daily_summary["날짜"], errors="coerce")
    factory_summary["생산일자"] = pd.to_datetime(factory_summary["생산일자"], errors="coerce")
    
    # 금일 데이터 제외 (아직 생산 중이므로)
    today = pd.Timestamp.today().date()
    daily_summary["날짜_date"] = daily_summary["날짜"].dt.date
    factory_summary["생산일자_date"] = factory_summary["생산일자"].dt.date
    
    # 제목
    st.markdown("<h1 style='text-align:center; color:#1f3a93; margin:0;'>🏭 APS 유효생산량 대시보드</h1>", unsafe_allow_html=True)
    
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    
    # 기간 필터
    filter_option = st.radio("조회 기간", ["당월", "전월", "기간조회"], horizontal=True, label_visibility="collapsed")
    
    # 날짜 범위 계산
    today_ts = pd.Timestamp.today()
    current_month_start = pd.Timestamp(year=today_ts.year, month=today_ts.month, day=1)
    current_month_end = today_ts - pd.Timedelta(days=1)  # 어제까지
    
    # 전월 계산
    first_day_current = pd.Timestamp(year=today_ts.year, month=today_ts.month, day=1)
    last_day_prev = first_day_current - pd.Timedelta(days=1)
    prev_month_start = pd.Timestamp(year=last_day_prev.year, month=last_day_prev.month, day=1)
    
    # 날짜 범위 결정
    if filter_option == "당월":
        start_date = current_month_start.date()
        end_date = current_month_end.date()
    elif filter_option == "전월":
        start_date = prev_month_start.date()
        end_date = last_day_prev.date()
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
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "총실적\n(pcs)",
            f"{total_prod:,}",
            delta=None,
            delta_color="off"
        )
    
    with col2:
        valid_rate = (valid_prod / total_prod * 100) if total_prod > 0 else 0
        st.metric(
            "유효생산량\n(pcs)",
            f"{valid_prod:,}",
            delta=None,
            delta_color="off"
        )
        st.markdown(
            f"<div style='display:inline-block; width:auto; text-align:center; font-size:32px; font-weight:900; color:#047857; margin:4px auto 0; padding:4px 14px; border-radius:8px;'> {valid_rate:.1f}% </div>",
            unsafe_allow_html=True
        )
    
    with col3:
        over_rate = (over_prod / total_prod * 100) if total_prod > 0 else 0
        st.metric(
            "과생산량\n(pcs)",
            f"{over_prod:,}",
            delta=None,
            delta_color="off"
        )
        st.markdown(
            f"<div style='display:inline-block; width:auto; text-align:center; font-size:32px; font-weight:900; color:#b91c1c; margin:4px auto 0; padding:4px 14px; border-radius:8px;'> {over_rate:.1f}% </div>",
            unsafe_allow_html=True
        )
    
    with col4:
        waste_rate = (waste_prod / total_prod * 100) if total_prod > 0 else 0
        st.metric(
            "불필요생산수량\n(pcs)",
            f"{waste_prod:,}",
            delta=None,
            delta_color="off"
        )
        st.markdown(
            f"<div style='display:inline-block; width:auto; text-align:center; font-size:32px; font-weight:900; color:#b45309; margin:4px auto 0; padding:4px 14px; border-radius:8px;'> {waste_rate:.1f}% </div>",
            unsafe_allow_html=True
        )
    
    st.markdown("<div style='margin-top:50px'></div>", unsafe_allow_html=True)
    # ============== 중간: 차트 ==============
    st.markdown("### 📈 공장별 현황")
    
    if len(factory_summary_filtered) == 0:
        st.info("선택한 기간에 공장별 데이터가 없습니다.")
    else:
        # 공장별 데이터 준비
        factory_data = factory_summary_filtered.groupby("공장", dropna=False).agg({
            "총실적": "sum",
            "유효생산량": "sum",
            "과생산량": "sum",
            "불필요생산량": "sum"
        }).reset_index()
        
        factory_data["유효비율(%)"] = (factory_data["유효생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["과생산비율(%)"] = (factory_data["과생산량"] / factory_data["총실적"] * 100).fillna(0)
        factory_data["불필요비율(%)"] = (factory_data["불필요생산량"] / factory_data["총실적"] * 100).fillna(0)
        
        metric_option = st.radio(
            "표시할 지표를 선택하세요",
            ["유효율", "과생산율", "불필요율"],
            horizontal=True
        )
        metric_desc = {
            "유효율": "해당일 필요수량 대비 유효하게 생산된 비율",
            "과생산율": "해당일 필요수량을 초과해 추가 생산된 비율",
            "불필요율": "해당일 필요수량 대비 불필요 규격이 생산된 비율"
        }
        st.caption(f"설명: {metric_desc[metric_option]}")
        
        metric_map = {
            "유효율": ("유효비율(%)", "유효생산량"),
            "과생산율": ("과생산비율(%)", "과생산량"),
            "불필요율": ("불필요비율(%)", "불필요생산량")
        }
        metric_col, pcs_col = metric_map[metric_option]
        factory_data["선택지표"] = factory_data[metric_col]
        
        fig = px.bar(
            factory_data,
            x="공장",
            y="선택지표",
            color="공장",
            title=f"공장별 {metric_option} (%)",
            text="선택지표"
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
            yaxis=dict(
                range=[0, 100],
                title=dict(text=f"{metric_option} (%)", font=dict(size=16, family="Arial", color="#222222"))
            ),
            xaxis=dict(
                title=dict(text="공장", font=dict(size=16, family="Arial", color="#222222")),
                tickfont=dict(size=18, family="Arial", color="#222222")
            ),
            title=dict(font=dict(size=22, family="Arial", color="#111111"))
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown(f"**선택 지표: {metric_option} (%)**")
        
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
        
        # 선택지표 추가
        metric_map = {
            "유효율": ("유효비율(%)", "유효생산량"),
            "과생산율": ("과생산비율(%)", "과생산량"),
            "불필요율": ("불필요비율(%)", "불필요생산량")
        }
        metric_col, pcs_col = metric_map[metric_option]
        combined_summary["선택지표"] = combined_summary[metric_col]
        
        # 테이블 표시
        display_combined = combined_summary[["공장", "신규분류요약", "총실적", pcs_col, "선택지표"]].copy()
        display_combined.columns = ["공장", "신규분류요약", "총실적 (pcs)", f"{pcs_col} (pcs)", f"{metric_option} (%)"]
        
        # 공장 순서 지정 (A관 > C관 > S관)
        factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
        display_combined["_factory_sort"] = display_combined["공장"].map(factory_order)
        display_combined = display_combined.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True)
        display_combined = display_combined.drop("_factory_sort", axis=1)
        
        display_combined["총실적 (pcs)"] = display_combined["총실적 (pcs)"].map("{:,.0f}".format)
        display_combined[f"{pcs_col} (pcs)"] = display_combined[f"{pcs_col} (pcs)"].map("{:,.0f}".format)
        display_combined[f"{metric_option} (%)"] = display_combined[f"{metric_option} (%)"].map("{:.1f}%".format)
        
        html = f"""
        <style>
            .custom-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
            .custom-table th, .custom-table td {{ padding: 10px 12px; border: 1px solid #e2e8f0; }}
            .custom-table th {{ background: #f8fafc; color: #111827; text-align: left; }}
            .custom-table td {{ vertical-align: middle; }}
            .custom-table td.number {{ text-align: right; }}
            .custom-table tbody tr:nth-child(even) {{ background: #f8fafc22; }}
        </style>
        <table class="custom-table">
          <thead>
            <tr>
              <th>공장</th>
              <th>신규분류요약</th>
              <th>총실적 (pcs)</th>
              <th>{pcs_col} (pcs)</th>
              <th>{metric_option} (%)</th>
            </tr>
          </thead>
          <tbody>
        """
        
        grouped = display_combined.groupby("공장", sort=False)
        for factory_name, group in grouped:
            rowspan = len(group)
            for idx, row in group.iterrows():
                html += "<tr>"
                if idx == group.index[0]:
                    html += f"<td rowspan='{rowspan}' style='vertical-align: middle; font-weight: 600;'>{factory_name}</td>"
                html += f"<td>{row['신규분류요약']}</td>"
                html += f"<td class='number'>{row['총실적 (pcs)']}</td>"
                html += f"<td class='number'>{row[f'{pcs_col} (pcs)']}</td>"
                html += f"<td class='number'>{row[f'{metric_option} (%)']}</td>"
                html += "</tr>"
        html += "</tbody></table>"
        st.markdown(html, unsafe_allow_html=True)
    
    # ============== 일별 요약 ==============
    st.markdown("### 📊 일별 요약")
    
    daily_display = daily_summary_filtered[["날짜", "총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량", "유효비율(%)"]].copy()
    # 날짜는 일자까지만 표시 (시간 제거)
    daily_display["날짜"] = daily_display["날짜"].dt.strftime("%Y-%m-%d")
    # pcs 컬럼은 콤마 표시 및 컬럼명에 (pcs) 추가
    pcs_cols = ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]
    daily_display.rename(columns={c: f"{c} (pcs)" for c in pcs_cols}, inplace=True)
    
    st.dataframe(
        daily_display.style.format({
            **{f"{c} (pcs)": "{:,.0f}" for c in pcs_cols},
            "유효비율(%)": "{:.1f}%"
        }),
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(f"❌ 오류가 발생했습니다: {str(e)}")
    st.info("결과 파일을 다시 생성해주세요.")
