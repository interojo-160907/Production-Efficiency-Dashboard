"""Read-only adapter for the unchanged Control Tower calculation module."""
import sqlite3
from contextlib import contextmanager
from pathlib import Path


def initialize_database(path):
    if not Path(path).is_file():
        raise FileNotFoundError('컨트롤타워 게시 데이터가 없습니다.')


@contextmanager
def connect(path):
    con = sqlite3.connect(Path(path).resolve().as_uri() + '?mode=ro', uri=True)
    con.row_factory = sqlite3.Row
    try:
        yield con
    finally:
        con.close()
