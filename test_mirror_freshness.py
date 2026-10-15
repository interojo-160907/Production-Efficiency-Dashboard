from datetime import date
import unittest
from unittest.mock import patch
from mirror_freshness import lag_days, publication_status


class FreshnessTests(unittest.TestCase):
    def test_completed_day_and_month_boundary(self):
        self.assertEqual(lag_days('2026-10-14', date(2026,10,15)), 0)
        self.assertEqual(lag_days('2026-10-13', date(2026,10,15)), 1)
        self.assertEqual(lag_days('2026-09-30', date(2026,10,1)), 0)
        self.assertEqual(lag_days('2026-10-15', date(2026,10,15)), 0)

    def test_new_publication_requests_full_rerun(self):
        with patch('mirror_freshness.mirror.revision', return_value=2), \
             patch('mirror_freshness.st.rerun', side_effect=RuntimeError('rerun')):
            with self.assertRaisesRegex(RuntimeError, 'rerun'):
                publication_status.__wrapped__(1, date(2026,10,15))


if __name__ == '__main__':
    unittest.main()
