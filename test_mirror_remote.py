from contextlib import closing
import hashlib
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import Mock

from mirror_remote import PublicationSync, MAX_DATABASE_BYTES


class RemoteTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        for name in ('dashboard_snapshot.py', 'presentation.py'):
            (self.root / name).write_text('# calculation\n', encoding='utf-8')
        self.old = self.make_database('old.sqlite', 'a' * 64, '2026-10-19', 10)
        self.new = self.make_database('new.sqlite', 'b' * 64, '2026-10-21', 20)
        self.receipt = dict(content_hash='b' * 64, data_commit='c' * 40,
                            pushed_at='2026-10-22T10:17:03+09:00')
        self.fetch = Mock(side_effect=self.respond)
        self.now = 0
        self.sync = PublicationSync(self.old, self.root, fetch=self.fetch,
                                    clock=lambda: self.now, cache_dir=self.root / 'cache')

    def make_database(self, filename, version, day, qty):
        path = self.root / filename
        with closing(sqlite3.connect(path)) as con:
            con.execute('CREATE TABLE analysis_factory_daily(analysis_date TEXT, actual_qty INTEGER)')
            con.execute('INSERT INTO analysis_factory_daily VALUES (?,?)', (day, qty))
            meta = dict(schema_version=2, content_hash=version, date_to=day,
                        tables={'analysis_factory_daily': 1})
            for field, name in [('engine_sha256', 'dashboard_snapshot.py'), ('presentation_sha256', 'presentation.py')]:
                meta[field] = hashlib.sha256((self.root / name).read_text(encoding='utf-8').encode()).hexdigest()
            con.execute('CREATE TABLE mirror_meta(key TEXT, value TEXT)')
            con.execute('INSERT INTO mirror_meta VALUES (?,?)', ('metadata', json.dumps(meta)))
            con.commit()
        return path

    def respond(self, url, limit):
        return json.dumps(self.receipt).encode() if 'publication.json' in url else self.new.read_bytes()

    def test_new_publication_without_deploy_or_restart(self):
        path, error = self.sync.refresh()
        self.assertEqual(error, '')
        self.assertNotEqual(path, self.old)
        with closing(sqlite3.connect(path)) as con:
            self.assertEqual(con.execute('SELECT * FROM analysis_factory_daily').fetchone(), ('2026-10-21', 20))
        self.assertEqual(self.fetch.call_args.args,
                         (f'https://raw.githubusercontent.com/interojo-160907/Production-Efficiency-Dashboard/{"c" * 40}/outputs/mirror/dashboard.sqlite', MAX_DATABASE_BYTES))
        self.assertEqual(json.loads((path.parent / 'publication.json').read_text())['content_hash'], 'b' * 64)
        self.assertTrue(self.old.exists())

    def test_unchanged_only_fetches_small_receipt_and_throttles(self):
        self.receipt['content_hash'] = 'a' * 64
        self.assertEqual(self.sync.refresh(), (self.old, ''))
        self.sync.refresh()
        self.assertEqual(self.fetch.call_count, 1)
        self.now = 61
        self.sync.refresh()
        self.assertEqual(self.fetch.call_count, 2)

    def test_manual_check_bypasses_throttle(self):
        self.sync.refresh()
        self.fetch.reset_mock()
        self.sync.refresh(force=True)
        self.assertEqual(self.fetch.call_count, 1)

    def test_network_error_preserves_verified_snapshot_then_recovers(self):
        self.fetch.side_effect = TimeoutError()
        path, error = self.sync.refresh()
        self.assertEqual(path, self.old)
        self.assertTrue(error)
        self.fetch.side_effect = self.respond
        self.now = 61
        path, error = self.sync.refresh()
        self.assertNotEqual(path, self.old)
        self.assertEqual(error, '')

    def test_incomplete_download_never_replaces_current_data(self):
        self.fetch.side_effect = [json.dumps(self.receipt).encode(), b'partial sqlite']
        path, error = self.sync.refresh()
        self.assertEqual(path, self.old)
        self.assertTrue(error)
        self.assertFalse(list((self.root / 'cache').rglob('download.sqlite')))

    def test_incompatible_calculation_code_rejected(self):
        (self.root / 'dashboard_snapshot.py').write_text('# changed', encoding='utf-8')
        path, error = self.sync.refresh()
        self.assertEqual(path, self.old)
        self.assertIn('계산 코드', error)

    def test_receipt_database_mismatch_rejected(self):
        self.receipt['content_hash'] = 'd' * 64
        path, error = self.sync.refresh()
        self.assertEqual(path, self.old)
        self.assertIn('일치하지', error)

    def test_old_cdn_receipt_cannot_roll_back_data(self):
        path, _ = self.sync.refresh()
        self.receipt.update(content_hash='a' * 64, pushed_at='2026-10-20T10:17:03+09:00')
        self.assertEqual(self.sync.refresh(force=True)[0], path)
        self.assertIn('이전 게시', self.sync.error)

    def test_same_day_correction_and_old_reader_remains_consistent(self):
        prior, _ = self.sync.refresh()
        self.new = self.make_database('corrected.sqlite', 'd' * 64, '2026-10-21', 30)
        self.receipt.update(content_hash='d' * 64, pushed_at='2026-10-22T10:27:03+09:00')
        latest, error = self.sync.refresh(force=True)
        self.assertEqual(error, '')
        for path, expected in [(prior, 20), (latest, 30)]:
            with closing(sqlite3.connect(path)) as con:
                self.assertEqual(con.execute('SELECT actual_qty FROM analysis_factory_daily').fetchone()[0], expected)


if __name__ == '__main__':
    unittest.main()
