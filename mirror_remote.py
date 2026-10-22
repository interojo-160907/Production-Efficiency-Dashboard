"""Fetch published aggregates independently of Streamlit's Git deployment.

Only data is downloaded, from the configured public repository. A publication
receipt pins the SQLite download to an immutable commit. Validated snapshots
live outside the Git checkout so this cannot conflict with Cloud deployments.
"""
from contextlib import closing
from datetime import date, datetime
import hashlib
import json
from pathlib import Path
import re
import sqlite3
import tempfile
import threading
import time
from urllib.request import Request, urlopen


BASE_URL = 'https://raw.githubusercontent.com/interojo-160907/Production-Efficiency-Dashboard'
MAX_DATABASE_BYTES = 45 * 1024 * 1024


def download(url, limit):
    request = Request(url, headers={'Cache-Control': 'no-cache', 'User-Agent': 'APS-Yield-Mirror/1'})
    with urlopen(request, timeout=15) as response:
        data = response.read(limit + 1)
    if len(data) > limit:
        raise ValueError('게시 파일 용량 한도를 초과했습니다.')
    return data


def read_metadata(path):
    with closing(sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)) as con:
        return json.loads(con.execute("SELECT value FROM mirror_meta WHERE key='metadata'").fetchone()[0])


def validate(path, receipt, code_dir):
    meta = read_metadata(path)
    if meta.get('schema_version') != 2 or meta.get('content_hash') != receipt['content_hash']:
        raise ValueError('게시 기록과 데이터가 일치하지 않습니다.')
    for field, filename in [('engine_sha256', 'dashboard_snapshot.py'),
                            ('presentation_sha256', 'presentation.py')]:
        actual = hashlib.sha256((code_dir / filename).read_text(encoding='utf-8').encode()).hexdigest()
        if meta.get(field) != actual:
            raise ValueError('계산 코드가 변경되어 웹 배포 갱신이 필요합니다.')
    date.fromisoformat(meta['date_to'])
    with closing(sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)) as con:
        if con.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
            raise ValueError('게시 데이터 무결성 검사에 실패했습니다.')
        for table, count in meta['tables'].items():
            if not re.fullmatch(r'[a-z_]+', table):
                raise ValueError('게시 테이블 이름이 잘못되었습니다.')
            if con.execute(f'SELECT count(*) FROM "{table}"').fetchone()[0] != count:
                raise ValueError('게시 데이터 행 수가 일치하지 않습니다.')
        latest = con.execute('SELECT max(analysis_date) FROM analysis_factory_daily').fetchone()[0]
        if latest != meta['date_to']:
            raise ValueError('게시 데이터 기준일이 일치하지 않습니다.')
    return meta


class PublicationSync:
    def __init__(self, deployed_db, code_dir, *, fetch=download, clock=time.monotonic, cache_dir=None):
        self.path = Path(deployed_db)
        self.code_dir = Path(code_dir)
        self.fetch = fetch
        self.clock = clock
        self.cache_dir = Path(cache_dir) if cache_dir else Path(tempfile.mkdtemp(prefix='aps-mirror-'))
        self.lock = threading.Lock()
        self.next_check = 0
        self.error = ''
        self.checked_at = ''
        self.latest_pushed_at = None
        try:
            receipt = json.loads((self.path.parent / 'publication.json').read_text(encoding='utf-8'))
            if receipt['content_hash'] == read_metadata(self.path)['content_hash']:
                self.latest_pushed_at = datetime.fromisoformat(receipt['pushed_at'])
        except (OSError, ValueError, KeyError, sqlite3.Error):
            pass

    def refresh(self, *, force=False):
        with self.lock:
            if not force and self.clock() < self.next_check:
                return self.path, self.error
            try:
                # Cache-busting is necessary: raw main URLs can stay cached after a push.
                raw = self.fetch(f'{BASE_URL}/main/outputs/mirror/publication.json?check={time.time_ns()}', 16384)
                receipt = json.loads(raw)
                if not re.fullmatch(r'[0-9a-f]{40}', receipt['data_commit']):
                    raise ValueError('게시 커밋이 잘못되었습니다.')
                if not re.fullmatch(r'[0-9a-f]{64}', receipt['content_hash']):
                    raise ValueError('게시 버전이 잘못되었습니다.')
                published = datetime.fromisoformat(receipt['pushed_at'])
                if published.tzinfo is None:
                    raise ValueError('게시 시각의 시간대가 없습니다.')
                current = read_metadata(self.path)
                if self.latest_pushed_at and published < self.latest_pushed_at:
                    raise ValueError('이전 게시 기록이 응답되어 현재 데이터를 유지합니다.')
                if receipt['content_hash'] != current['content_hash']:
                    # Never overwrite a file being read by another viewer/session.
                    folder = self.cache_dir / receipt['content_hash']
                    folder.mkdir(parents=True, exist_ok=True)
                    target = folder / 'dashboard.sqlite'
                    candidate = folder / 'download.sqlite'
                    try:
                        candidate.write_bytes(self.fetch(
                            f"{BASE_URL}/{receipt['data_commit']}/outputs/mirror/dashboard.sqlite", MAX_DATABASE_BYTES))
                        validate(candidate, receipt, self.code_dir)
                        candidate.replace(target)
                        (folder / 'publication.json').write_text(json.dumps(receipt), encoding='utf-8')
                    finally:
                        candidate.unlink(missing_ok=True)
                    self.path = target
                self.latest_pushed_at = published
                self.error = ''
                self.checked_at = datetime.now().astimezone().isoformat(timespec='seconds')
            except Exception as exc:
                # Keep the last verified snapshot. Never render a half download.
                self.error = str(exc) if isinstance(exc, ValueError) else 'GitHub 최신 게시 확인에 실패했습니다. 잠시 후 자동 재시도합니다.'
            finally:
                self.next_check = self.clock() + 60
            return self.path, self.error
