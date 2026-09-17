"""Streamlit views over the same aggregates and calculations as Control Tower."""
import io
import hashlib
import json
import os
import time
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from tower_mirror.database import connect
from tower_mirror.presentation import balanced_percentages
from tower_mirror.dashboard_snapshot import (
    available_analysis_range, available_analysis_periods,
    dashboard_payload_for_range, response_payload_for_period,
)

DB = Path(__file__).parent / 'outputs/mirror/dashboard.sqlite'
LABELS = {
    'analysis_date': '날짜', 'factory': '공장', 'process': '공정',
    'process_code': '공정코드', 'classification': '신규분류요약',
    'actual_qty': '총 생산량', 'need_qty': '필요 수량', 'effective_qty': '정확 대응 수량',
    'excess_qty': '초과 생산 수량', 'nonstandard_qty': '비정형 생산 수량',
    'production_sku_count': '생산 SKU 수', 'responded_sku_count': '대응 SKU 수',
    'spec_rate': '규격 대응률(%)', 'exact_rate': '정확 대응 비중(%)',
    'excess_rate': '초과 생산 비중(%)', 'nonstandard_rate': '비정형 생산 비중(%)',
    'process_score': '공정점수', 'score': '점수', 'grade': '등급',
    'over_qty': '공정 초과 수량', 'over_rate': '공정 초과 비중(%)',
    'spec_source': '규격 집계 기준',
}
METRICS = {LABELS[k]: k for k in ('spec_rate', 'exact_rate', 'excess_rate', 'nonstandard_rate')}
COLORS = {'A관(1공장)': '#5B63F5', 'C관(2공장)': '#8B52FF', 'S관(3공장)': '#EB3994'}


@st.cache_data(show_spinner=False, max_entries=4)
def load_view(version, start, end):
    return (dashboard_payload_for_range(DB, start, end),
            response_payload_for_period(DB, start, end))


def frame(rows):
    result = pd.DataFrame(rows)
    result = result.drop(columns=['aps_snapshot_id', 'score_profile_id'], errors='ignore')
    return result.rename(columns=LABELS)


@st.cache_data(show_spinner=False, max_entries=4)
def excel(tables):
    stream = io.BytesIO()
    with pd.ExcelWriter(stream, engine='xlsxwriter') as writer:
        workbook = writer.book
        header = workbook.add_format({'bold': True, 'bg_color': '#EAF1FF'})
        number = workbook.add_format({'num_format': '#,##0.0'})
        for name, rows in tables.items():
            df = frame(rows)
            df.to_excel(writer, sheet_name=name[:31], index=False)
            sheet = writer.sheets[name[:31]]
            sheet.freeze_panes(1, 0)
            sheet.set_column(0, max(0, len(df.columns)-1), 20, number)
            for i, column in enumerate(df.columns):
                sheet.write(0, i, column, header)
            if len(df.columns):
                sheet.autofilter(0, 0, len(df), len(df.columns)-1)
    return stream.getvalue()


def run():
    # Cloud runs in UTC by default; both applications use Korean business dates.
    os.environ['TZ'] = 'Asia/Seoul'
    if hasattr(time, 'tzset'):
        time.tzset()
    st.set_page_config(page_title='생산 운영 현황 대시보드', layout='wide')
    st.markdown('''<style>.block-container{padding-top:2rem;max-width:1700px}
    [data-testid="stMetric"]{background:#f5f8ff;border:1px solid #e1e8f3;
    border-radius:12px;padding:16px}h1{color:#183756}</style>''', unsafe_allow_html=True)
    st.title('🏭 생산 운영 현황 대시보드')
    with connect(DB) as con:
        meta = json.loads(con.execute('SELECT value FROM mirror_meta WHERE key="metadata"').fetchone()[0])
    version = meta['content_hash']
    engine = Path(__file__).parent / 'tower_mirror/dashboard_snapshot.py'
    if hashlib.sha256(engine.read_text(encoding='utf-8').encode('utf-8')).hexdigest() != meta['engine_sha256']:
        st.error('게시 데이터와 계산 코드 버전이 다릅니다. 컨트롤타워에서 다시 게시해 주세요.')
        return
    presentation = Path(__file__).parent/'tower_mirror/presentation.py'
    if hashlib.sha256(presentation.read_text(encoding='utf-8').encode('utf-8')).hexdigest() != meta.get('presentation_sha256'):
        st.error('표시 기준 버전이 다릅니다. 컨트롤타워에서 다시 게시해 주세요.')
        return
    earliest, latest = available_analysis_range(DB)
    if not latest:
        st.info('조회 가능한 확정 데이터가 없습니다.'); return
    try:
        receipt = json.loads((DB.parent/'publication.json').read_text(encoding='utf-8'))
    except (OSError, ValueError):
        receipt = {}
    published = receipt.get('pushed_at') if receipt.get('content_hash') == version else None
    st.caption(f"마지막 게시 완료(KST): {published or '게시 완료 기록 확인 중'} · 실적 기준일: {latest} · 데이터 {version[:10]}")
    a, b, c = st.columns([1.2, 1, 2])
    mode = a.radio('조회 기간', ['당월', '전월', '월 선택', '기간조회'], horizontal=True)
    periods = available_analysis_periods(DB)
    newest = date.fromisoformat(latest)
    first = date.fromisoformat(earliest)
    if mode == '기간조회':
        dates = c.date_input('조회 기간 선택', (newest.replace(day=1), newest), min_value=first, max_value=newest)
        if len(dates) != 2:
            st.info('시작일과 종료일을 선택하세요.'); return
        start, end = (d.isoformat() for d in dates)
    else:
        import calendar
        if mode == '월 선택':
            year, month = b.selectbox('연·월', periods, format_func=lambda x: f'{x[0]}년 {x[1]}월')
        else:
            today = date.today()
            year, month = today.year, today.month
            if mode == '전월':
                year, month = (year-1, 12) if month == 1 else (year, month-1)
        start = date(year, month, 1).isoformat()
        end = min(date(year, month, calendar.monthrange(year, month)[1]).isoformat(), latest)
    if start > end:
        st.info('선택한 기간에 게시된 데이터가 없습니다. 월 선택에서 이전 자료를 조회하세요.'); return
    payload, response = load_view(version, start, end)
    if not payload or not response:
        st.info('선택한 기간에 데이터가 없습니다.'); return
    factory = st.radio('공장 구분', ['전체', *response['factories']], horizontal=True)
    selected = response['selected']['overall'] if factory == '전체' else response['selected']['factories'][factory]
    previous = response['previous']['overall'] if factory == '전체' else response['previous']['factories'][factory]
    columns = st.columns(5)
    columns[0].metric('총 생산량 (pcs)', f"{selected['actual_qty']:,.0f}")
    rounded = balanced_percentages([selected[k] for k in ('effective_qty', 'excess_qty', 'nonstandard_qty')])
    displayed = dict(zip(('exact_rate', 'excess_rate', 'nonstandard_rate'), rounded))
    for col, key in zip(columns[1:], METRICS.values()):
        col.metric(LABELS[key], f'{displayed.get(key, selected[key]):.1f}%', f'{selected[key]-previous[key]:+.1f}%p',
                   delta_color='inverse' if key in ('excess_rate', 'nonstandard_rate') else 'normal')
        quantity_key = {'exact_rate': 'effective_qty', 'excess_rate': 'excess_qty', 'nonstandard_rate': 'nonstandard_qty'}.get(key)
        if quantity_key:
            col.caption(f"{selected[quantity_key]:,.0f} pcs")
    st.caption(f"조회 {start} ~ {end} · 비교 {response['comparison_label']} {response['previous_date_from']} ~ {response['previous_date_to']}")
    if payload['stale_days']:
        st.warning(f"컨트롤타워 기준본과 분석본 확인 필요: {payload['stale_days']}일")
    tab = st.radio('화면', ['생산 현황', '공정 밸런스'], horizontal=True, label_visibility='collapsed')
    def filtered(rows):
        return [r for r in rows if factory == '전체' or r.get('factory') == factory]
    if tab == '생산 현황':
        metric = METRICS[st.radio('비교 지표', list(METRICS), horizontal=True)]
        bars = [{'factory': f, '선택 기간': item[metric],
                 '비교 기간': response['previous']['factories'][f][metric]}
                for f, item in response['selected']['factories'].items() if factory == '전체' or f == factory]
        left, right = st.columns(2)
        fig = go.Figure()
        fig.add_bar(x=[r['factory'] for r in bars], y=[r['선택 기간'] for r in bars], name='선택 기간',
                    marker_color=[COLORS.get(r['factory'], '#1976ef') for r in bars], texttemplate='%{y:.1f}%')
        fig.add_scatter(x=[r['factory'] for r in bars], y=[r['비교 기간'] for r in bars], name=response['comparison_label'],
                        mode='markers', marker={'symbol': 'line-ew', 'size': 30, 'color': '#52627a', 'line': {'width': 2}})
        fig.update_layout(title='공장별 '+LABELS[metric], yaxis_title='%', height=370)
        left.plotly_chart(fig, use_container_width=True)
        trend = [{'날짜': d['analysis_date'], '공장': f, '비율': values[metric]}
                 for d in response['daily'] for f, values in d['factories'].items()
                 if factory == '전체' or f == factory]
        fig = px.line(pd.DataFrame(trend), x='날짜', y='비율', color='공장', color_discrete_map=COLORS,
                      title='일별 '+LABELS[metric]+' 추이')
        fig.update_layout(height=370)
        right.plotly_chart(fig, use_container_width=True)
        tables = {'공장별 요약': filtered(payload['production']['factories']),
                  '공장별 일별 실적': filtered(payload['production']['factory_daily']),
                  '신규분류별 상세': filtered(payload['production']['classifications'])}
    else:
        factories = filtered(payload['balance']['factories'])
        summary = [{k: v for k, v in r.items() if k != 'processes'} for r in factories]
        matrix = [{'공장': r['factory'], '공정': p['process'], '점수': p['score']}
                  for r in factories for p in r['processes']]
        left, right = st.columns([1, 2])
        left.plotly_chart(px.bar(frame(summary), x='공장', y='점수', color='공장',
                                color_discrete_map=COLORS, title='공장별 운영 밸런스'), use_container_width=True)
        pivot = pd.DataFrame(matrix).pivot(index='공장', columns='공정', values='점수')
        order = ['사출', '분리', '하드레이션', '접착', '누수규격']
        right.plotly_chart(px.imshow(pivot.reindex(columns=order), text_auto='.1f',
                                    zmin=0, zmax=100, color_continuous_scale='RdYlGn',
                                    title='공장·공정 점수'), use_container_width=True)
        st.caption('컨트롤타워와 동일: 일별 공정점수의 생산량 가중평균 · 공장등급은 공정등급 다수결')
        tables = {'공장별 요약': summary, '공장별 공정 점수': [dict(factory=r['factory'], **p) for r in factories for p in r['processes']],
                  '일별 공정 상세': filtered(payload['balance']['daily']),
                  '일별 신규분류 상세': filtered(payload['balance']['classifications'])}
    st.dataframe(frame(tables['공장별 요약']), use_container_width=True, hide_index=True)
    if st.toggle('상세 테이블 표시', value=False, key='mirror_details_'+tab):
        table_name = st.selectbox('상세 테이블', [name for name in tables if name != '공장별 요약'])
        st.dataframe(frame(tables[table_name]), use_container_width=True, hide_index=True)
    # Excel formatting is deferred until explicitly requested, keeping every
    # ordinary date/factory interaction free of spreadsheet generation work.
    export_key = (version, start, end, factory, tab)
    if st.button('조회 결과 엑셀 준비'):
        with st.spinner('엑셀 파일 준비 중…'):
            st.session_state['mirror_export'] = (export_key, excel(tables))
    prepared = st.session_state.get('mirror_export')
    if prepared and prepared[0] == export_key:
        st.download_button('조회 결과 엑셀 다운로드', data=prepared[1],
                       file_name=f'생산운영_{tab}_{start}_{end}.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
