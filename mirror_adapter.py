"""Feed Control Tower aggregates into the existing dashboard's data contract.

No layout, chart, widget or export presentation is implemented here.
"""
import hashlib
import json
from pathlib import Path
import pandas as pd
import streamlit as st
from tower_mirror.database import connect
from tower_mirror.dashboard_snapshot import _score, _grade, PROCESS_NAMES
from tower_mirror.presentation import balanced_percentages

DB = Path(__file__).parent/'outputs/mirror/dashboard.sqlite'
TABLES = {'result1_daily', 'result1_factory', 'result1_spec_daily',
          'result1_spec_factory_daily', 'result1_spec_factory_class_daily',
          'result2_daily', 'result2_factory', 'result2_process_daily',
          'result2_proc_base', 'result2_det_base'}


def active(path=None):
    return DB.is_file() and (path is None or Path(path).resolve() == DB.parent.resolve())


def revision():
    return DB.stat().st_mtime_ns if DB.exists() else 0


@st.cache_data(show_spinner=False, max_entries=2)
def metadata(version):
    with connect(DB) as con:
        result = json.loads(con.execute('SELECT value FROM mirror_meta WHERE key="metadata"').fetchone()[0])
    for field, name in [('engine_sha256','dashboard_snapshot.py'), ('presentation_sha256','presentation.py')]:
        text = (Path(__file__).parent/'tower_mirror'/name).read_text(encoding='utf-8')
        if hashlib.sha256(text.encode('utf-8')).hexdigest() != result[field]:
            raise RuntimeError('컨트롤타워와 웹의 계산 기준 버전이 다릅니다.')
    return result


def published_at():
    meta = metadata(revision())
    try:
        receipt = json.loads((DB.parent/'publication.json').read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return '게시 완료 기록 확인 중'
    if receipt.get('content_hash') != meta['content_hash']:
        return '게시 완료 기록 확인 중'
    from datetime import datetime
    from zoneinfo import ZoneInfo
    return datetime.fromisoformat(receipt['pushed_at']).astimezone(ZoneInfo('Asia/Seoul')).strftime('%Y-%m-%d %H:%M:%S')


@st.cache_data(show_spinner=False, max_entries=2)
def profile(version):
    with connect(DB) as con:
        return dict(con.execute('SELECT * FROM score_profile WHERE is_active=1 ORDER BY profile_id DESC LIMIT 1').fetchone())


def months(table):
    source = 'analysis_factory_daily' if table.startswith('result1') else 'analysis_process_daily'
    with connect(DB) as con:
        return tuple(r[0] for r in con.execute(f'SELECT DISTINCT substr(analysis_date,1,7) FROM {source} ORDER BY 1'))


@st.cache_data(show_spinner=False, max_entries=12)
def _source(table, selected_months, version):
    metadata(version)
    if table not in ('analysis_factory_daily','analysis_process_daily','analysis_classification_daily'):
        raise ValueError(table)
    with connect(DB) as con:
        if not selected_months:
            return pd.read_sql_query(f'SELECT * FROM {table} WHERE 0', con)
        # Date predicates use the publication DB's index.
        clauses=[];params=[]
        import calendar
        for month in selected_months:
            year, num = map(int,month.split('-'))
            clauses.append('(analysis_date BETWEEN ? AND ?)')
            params.extend([month+'-01',month+f'-{calendar.monthrange(year,num)[1]:02d}'])
        return pd.read_sql_query(f'SELECT * FROM {table} WHERE '+ ' OR '.join(clauses)+' ORDER BY analysis_date,factory',con,params=params)


def _rates(df):
    denom=df['총실적'].replace(0,float('nan'))
    for qty,label in [('유효생산량','유효비율(%)'),('과생산량','과생산비율(%)'),('불필요생산량','불필요비율(%)')]:
        df[label]=(df[qty]/denom*100).fillna(0)
    return df


def load(table, selected_months, columns=None):
    version=revision()
    if table not in TABLES:
        return pd.DataFrame()
    process=_source('analysis_process_daily',selected_months,version)
    if table == 'result1_daily':
        source=_source('analysis_factory_daily',selected_months,version)
        needs=process[process['process_code']=='80'][['analysis_date','factory','need_qty']]
        source=source.merge(needs,on=['analysis_date','factory'],how='left')
        df=source.rename(columns={'analysis_date':'날짜','actual_qty':'총실적','need_qty':'총부족수량',
            'effective_qty':'유효생산량','excess_qty':'과생산량','nonstandard_qty':'불필요생산량'})
        df=df.groupby('날짜',as_index=False)[['총실적','총부족수량','유효생산량','과생산량','불필요생산량']].sum()
        return _rates(df)
    if table.startswith('result1_spec'):
        is_class=table.endswith('_class_daily')
        source=_source('analysis_classification_daily' if is_class else 'analysis_factory_daily',selected_months,version)
        if is_class:source=source[source['process_code']=='80'].copy()
        df=source.rename(columns={'analysis_date':'날짜','factory':'공장','classification':'신규분류요약',
            'production_sku_count':'생산SKU수','responded_sku_count':'필요대응SKU수'})
        keys=['날짜']+(['공장'] if table!='result1_spec_daily' else [])+(['신규분류요약'] if is_class else [])
        df=df.groupby(keys,as_index=False)[['생산SKU수','필요대응SKU수']].sum()
        df['규격대응률(%)']=(df['필요대응SKU수']/df['생산SKU수'].replace(0,float('nan'))*100).fillna(0)
        return df
    if table in ('result1_factory','result2_factory','result2_daily','result2_det_base'):
        source=_source('analysis_classification_daily',selected_months,version)
        if table=='result1_factory':source=source[source['process_code']=='80'].copy()
        df=source.rename(columns={'analysis_date':'날짜','factory':'공장','classification':'신규분류요약',
            'actual_qty':'_actual','need_qty':'_need','effective_qty':'유효생산량','excess_qty':'과생산량',
            'nonstandard_qty':'불필요생산량','shortage_qty':'부족수량'})
        df['공정']=df['process_code'].map(PROCESS_NAMES)
        if table=='result2_det_base':
            df=df.rename(columns={'_actual':'실적수량','_need':'필요수량'})
            df['공장그룹']=df['공장'].str[:2]
            return df[['날짜','공장','공장그룹','공정','신규분류요약','실적수량','필요수량','부족수량','유효생산량','과생산량','불필요생산량']]
        df=df.rename(columns={'_actual':'총실적','_need':'총부족수량'})
        qty=['총실적','총부족수량','유효생산량','과생산량','불필요생산량']
        if table=='result2_daily':return _rates(df.groupby(['날짜','공정'],as_index=False)[qty].sum())
        df=df.rename(columns={'날짜':'생산일자'})
        keys=['생산일자','공장']+(['공정'] if table=='result2_factory' else [])+['신규분류요약']
        return _rates(df[keys+qty].copy())
    df=process.rename(columns={'analysis_date':'날짜','factory':'공장','actual_qty':'실적수량',
        'need_qty':'필요수량','effective_qty':'유효생산량','over_qty':'과생산량',
        'nonstandard_qty':'불필요생산량','production_sku_count':'생산SKU수','responded_sku_count':'규격대응SKU수'})
    df['공정']=df['process_code'].map(PROCESS_NAMES)
    df['공장그룹']=df['공장'].str[:2]
    # Unmet quantity is additive over item/classification rows, not max of totals.
    classes=_source('analysis_classification_daily',selected_months,version)
    shortage=classes.groupby(['analysis_date','factory','process_code'],as_index=False)['shortage_qty'].sum()
    shortage=shortage.rename(columns={'analysis_date':'날짜','factory':'공장','shortage_qty':'부족수량'})
    df=df.merge(shortage,on=['날짜','공장','process_code'],how='left')
    if table=='result2_process_daily':
        return df[['날짜','공장','공정','실적수량','부족수량','과생산량']].rename(columns={'과생산량':'과생산수량'})
    criteria=profile(version)
    df['_tower_score']=[_score(float(r.spec_rate),float(r.exact_rate),float(r.over_rate),float(r.nonstandard_rate),criteria) for r in df.itertuples()]
    df['_tower_grade']=df['_tower_score'].map(lambda score:_grade(score,criteria))
    return df


def grade(score):
    return _grade(float(score),profile(revision()))
