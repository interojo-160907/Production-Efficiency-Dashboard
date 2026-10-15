"""Keep an open dashboard in sync with the deployed publication."""
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import streamlit as st
import mirror_adapter as mirror


def lag_days(date_to, today):
    """Only completed days count; weekends are not silently assumed idle."""
    return max(0, ((today - timedelta(days=1)) - date.fromisoformat(date_to)).days)


@st.fragment(run_every=60)
def publication_status(rendered_revision, rendered_day):
    today = datetime.now(ZoneInfo("Asia/Seoul")).date()
    current = mirror.revision()
    if current != rendered_revision or today != rendered_day:
        st.rerun()
    meta = mirror.metadata(current)
    st.caption(f"마지막 게시 완료(KST): {mirror.published_at()} · 데이터: {meta['date_to']}까지")
    lag = lag_days(meta['date_to'], today)
    if lag:
        st.warning(
            f"데이터가 전일 기준보다 {lag}일 늦습니다. "
            f"현재 웹 데이터는 {meta['date_to']}까지입니다. "
            "새로 확인 후에도 같으면 컨트롤타워 게시 상태와 Streamlit 배포 상태를 확인하세요."
        )
    if st.button("최신 데이터 다시 확인", key="mirror_refresh"):
        st.rerun()
    st.caption("열어 둔 화면은 60초마다 새 게시본을 확인합니다. 서버에 배포된 자료만 표시합니다.")
