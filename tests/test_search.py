from pathlib import Path
import sys
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from core import Dataset
from data_search import SearchCancelled, search_records


class SearchTests(unittest.TestCase):
    def setUp(self):
        self.data = Dataset(['姓名', '备注'], [{'姓名': '张三', '备注': 'Straße'}, {'姓名': '李四', '备注': 'Hello WORLD'}], [2, 42], 'CSV')

    def test_unicode_and_case_insensitive_search(self):
        self.assertEqual(search_records(self.data, 'STRASSE'), [0])
        self.assertEqual(search_records(self.data, ' world '), [1])
        self.assertEqual(search_records(self.data, '李四'), [1])

    def test_source_row_search(self):
        self.assertEqual(search_records(self.data, '42'), [1])

    def test_empty_query_restores_all_records(self):
        self.assertEqual(search_records(self.data, ' '), [0, 1])

    def test_cancelled_search_returns_no_partial_result(self):
        stop = threading.Event()
        with self.assertRaises(SearchCancelled):
            search_records(self.data, '不存在', cancel=stop, progress=lambda *args: stop.set())

    def test_search_leaves_business_data_unchanged(self):
        before = [dict(row) for row in self.data.records], list(self.data.row_numbers)
        self.assertEqual(search_records(self.data, '不存在'), [])
        self.assertEqual((self.data.records, self.data.row_numbers), before)

    def test_large_search_does_not_flood_ui_with_progress_events(self):
        data = Dataset(['字段'], [{'字段': '内容'}] * 50000, list(range(2, 50002)), 'CSV')
        progress = []
        self.assertEqual(search_records(data, '不存在', progress=lambda *args: progress.append(args)), [])
        self.assertLessEqual(len(progress), 101)
        self.assertEqual(progress[0], (0, 50000))
        self.assertEqual(progress[-1], (50000, 50000))


if __name__ == '__main__':
    unittest.main()
