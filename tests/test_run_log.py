from datetime import datetime, timezone, timedelta
from decimal import Decimal
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import test_prices
from daily_sales import process
from run_log import RunLog


class RunLogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        env = patch.dict(os.environ, LOCALAPPDATA=str(self.root / 'appdata'))
        env.start()
        self.addCleanup(env.stop)

    def records(self):
        paths = list((self.root / 'appdata' / 'sold_of_date' / 'logs').glob('*.jsonl'))
        self.assertEqual(len(paths), 1)
        return [json.loads(line) for line in paths[0].read_text(encoding='utf-8').splitlines()]

    def source(self):
        return test_prices.PriceTests().make_source(self.root)

    def test_success_modes_append_counts_and_absolute_paths(self):
        source = self.source()
        first = process(source)
        with patch('exchange_rates.fetch_rates', return_value={'德国': Decimal('7.63')}):
            second = process(source, all_calculations=True)
        daily, prices = self.records()
        self.assertEqual(daily['mode'], '仅日销')
        self.assertEqual(prices['mode'], '日销与价格')
        self.assertEqual(prices['status'], 'success')
        self.assertEqual(prices['total_rows'], 5)
        self.assertEqual(prices['daily_calculated'], 4)
        self.assertEqual(prices['daily_skipped_no_sales'], 1)
        self.assertEqual(prices['price_calculated'], 3)
        self.assertEqual(prices['price_skipped_no_rate'], 1)
        self.assertEqual(prices['price_skipped_empty'], 1)
        self.assertEqual(daily['price_calculated'], 0)
        self.assertEqual(prices['input_path'], str(source.resolve()))
        self.assertEqual(prices['input_filename'], source.name)
        self.assertEqual(prices['output_path'], second['output'])
        self.assertTrue(Path(first['log_path']).is_file())
        self.assertIsNotNone(datetime.fromisoformat(prices['started_at']).utcoffset())
        self.assertIsNotNone(datetime.fromisoformat(prices['ended_at']).utcoffset())
        self.assertGreaterEqual(prices['duration_seconds'], 0)

    def test_failure_retains_counts_without_exception_secrets(self):
        source = self.source()
        with patch('exchange_rates.fetch_rates', side_effect=ValueError('password=DO_NOT_LOG')):
            with self.assertRaises(ValueError):
                process(source, all_calculations=True)
        record, = self.records()
        self.assertEqual(record['status'], 'failed')
        self.assertEqual(record['daily_calculated'], 4)
        self.assertEqual(record['stage'], '价格计算与汇率读取')
        self.assertNotIn('DO_NOT_LOG', json.dumps(record))
        self.assertFalse(Path(record['output_path']).exists())

    def test_early_failure_unknown_count(self):
        with self.assertRaises(ValueError):
            process(self.root / 'bad.csv')
        record, = self.records()
        self.assertIsNone(record['total_rows'])
        self.assertIsNone(record['output_path'])
        self.assertEqual(record['status'], 'failed')

    def test_correction_and_missing_history_counts(self):
        source = self.source()
        with ZipFile(source) as z:
            parts = {name: z.read(name) for name in z.namelist()}
        sheet = 'xl/worksheets/sheet1.xml'
        parts[sheet] = parts[sheet].replace(b'<v>107</v>', b'<v>90</v>', 1).replace(
            b'<c r="E3"><v>100</v></c>', b'')
        with ZipFile(source, 'w') as z:
            for name, content in parts.items():
                z.writestr(name, content)
        process(source)
        record, = self.records()
        self.assertEqual(record['sales_corrected'], 1)
        self.assertEqual(record['daily_skipped_no_history'], 1)
        self.assertEqual(record['daily_calculated'], 3)

    def test_cross_month_uses_start_month_and_monotonic_duration(self):
        tz = timezone(timedelta(hours=8))
        with patch('run_log.now', side_effect=[datetime(2026, 10, 31, 23, 59, 59, tzinfo=tz),
                                              datetime(2026, 11, 1, tzinfo=tz)]), \
                patch('run_log.time.perf_counter', side_effect=[10, 11.25]):
            path = RunLog(self.root / 'a.xlsx', None, False).finish()
        self.assertEqual(path.name, '2026-10.jsonl')
        record, = self.records()
        self.assertEqual(record['duration_seconds'], 1.25)
        self.assertTrue(record['ended_at'].endswith('+08:00'))

    def test_logging_failure_preserves_success_and_original_error(self):
        source = self.source()
        with patch('run_log.RunLog.finish', side_effect=PermissionError('private detail')), \
                patch('sys.stderr') as stderr:
            result = process(source)
            self.assertTrue(Path(result['output']).exists())
            with self.assertRaisesRegex(ValueError, '仅支持'):
                process(self.root / 'bad.csv')
            self.assertNotIn('private detail', str(stderr.write.call_args_list))


if __name__ == '__main__':
    unittest.main()
