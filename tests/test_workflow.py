"""Regression coverage for data integrity, retries, recovery and portable results."""
import csv
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from core import (CheckCancelled, Dataset, InputError, Template, check_records,
                  continue_batch, file_stem, generate_batch, read_data, retry_batch)
from batch_state import load_batch_state, save_batch_state, validate_batch_files
from check_report import write_check_report
from result_archive import ArchiveCancelled, archive_batch


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        base = ROOT / '.build/test-temp'
        base.mkdir(parents=True, exist_ok=True)
        self.temporary = tempfile.TemporaryDirectory(dir=base)
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.template = Template(ROOT / '示例文件/通知模板.docx')
        self.dataset = read_data(ROOT / '示例文件/人员数据.xlsx')
        self.mapping = {name: name for name in self.template.fields}
        self.pattern = '{{序号}}_{{姓名}}'
        self.job = self.template, self.dataset, self.mapping, self.pattern, ('姓名',)

    def generate(self, **kwargs):
        return generate_batch(self.template, self.dataset, self.mapping, self.folder, self.pattern, **kwargs)

    def test_excel_and_csv_agree(self):
        other = read_data(ROOT / '示例文件/人员数据.csv')
        self.assertEqual(self.dataset.records, other.records)
        self.assertEqual(self.dataset.row_numbers, [2, 3, 4])

    def test_duplicate_headers_rejected(self):
        source = self.folder / 'bad.csv'
        source.write_text('姓名,姓名\n甲,乙\n', encoding='utf-8')
        with self.assertRaises(InputError):
            read_data(source)

    def test_multiline_csv_and_original_row_numbers(self):
        source = self.folder / 'data.csv'
        source.write_text('姓名;备注\n甲;"第一行\n第二行"\n;\n乙;结尾\n', encoding='utf-8', newline='')
        data = read_data(source)
        self.assertEqual(data.records[0]['备注'], '第一行\n第二行')
        self.assertEqual(data.row_numbers, [2, 4])

    def test_required_field_checks_only_selected_scope(self):
        self.dataset.records[0]['姓名'] = ' '
        issues = check_records(self.template, self.dataset, self.mapping, self.pattern, required_fields=['姓名'])
        self.assertTrue(any(i['级别'] == '错误' and i['行号'] == 2 for i in issues))
        self.assertFalse(check_records(self.template, self.dataset, self.mapping, self.pattern,
                                       indices=[1], required_fields=['姓名']))

    def test_invalid_xml_character_blocked(self):
        self.dataset.records[0]['姓名'] = '\x01'
        issues = check_records(self.template, self.dataset, self.mapping, self.pattern)
        self.assertTrue(any(i['级别'] == '错误' for i in issues))
        result = self.generate(indices=[0])
        self.assertEqual(result.items[0]['状态'], '失败')
        self.assertFalse(list(result.folder.glob('*.docx')))

    def test_cancelled_check_is_not_complete(self):
        stop = threading.Event()
        stop.set()
        with self.assertRaises(CheckCancelled):
            check_records(self.template, self.dataset, self.mapping, self.pattern, cancel=stop)

    def test_docx_content_and_other_parts_preserved(self):
        result = self.generate()
        self.assertEqual(result.succeeded, 3)
        for item, record in zip(result.items, self.dataset.records):
            with zipfile.ZipFile(result.folder / item['Word 文件']) as archive:
                self.assertIsNone(archive.testzip())
                text = '\n'.join(archive.read(n).decode('utf-8') for n in self.template.parts)
                self.assertIn(record['姓名'], text)
                self.assertNotIn('{{', text)
                for entry, content in self.template.entries:
                    if entry.filename not in self.template.parts:
                        self.assertEqual(archive.read(entry.filename), content)

    def test_duplicate_names_are_unique(self):
        result = generate_batch(self.template, self.dataset, self.mapping, self.folder, '相同名称')
        self.assertEqual([i['Word 文件'] for i in result.items], ['相同名称.docx', '相同名称_2.docx', '相同名称_3.docx'])

    def test_windows_reserved_name_is_adjusted(self):
        self.assertEqual(file_stem('CON', {}, 1), '_CON')
        self.assertNotIn('/', file_stem('a/b', {}, 1))

    def test_progress_write_failure_stops_with_completed_files_preserved(self):
        def checkpoint(result, index=None, item=None):
            if index is not None:
                raise OSError('磁盘无法写入')
        result = self.generate(checkpoint=checkpoint)
        self.assertTrue(result.cancelled)
        self.assertEqual(result.succeeded, 1)
        self.assertEqual(result.remaining_indices, [1, 2])
        self.assertTrue((result.folder / result.items[0]['Word 文件']).is_file())
        self.assertIn('磁盘无法写入', result.progress_warning)

    def test_stop_reload_continue_does_not_rewrite_success(self):
        stop = threading.Event()
        result = self.generate(cancel=stop, progress=lambda *args: stop.set())
        self.assertEqual(result.remaining_indices, [1, 2])
        save_batch_state(result, self.job)
        loaded, job = load_batch_state(result.folder)
        first = result.folder / result.items[0]['Word 文件']
        before = first.read_bytes(), first.stat().st_mtime_ns
        resumed = continue_batch(loaded, *job[:4], required_fields=job[4])
        self.assertEqual(resumed.succeeded, 3)
        self.assertEqual((first.read_bytes(), first.stat().st_mtime_ns), before)

    def test_pdf_failure_keeps_word_and_retry_only_converts(self):
        with patch('core.pdf_converter', return_value=('test', '')), patch('core.PdfSession.convert', side_effect=RuntimeError('转换失败')):
            result = self.generate(export_pdf=True)
        self.assertEqual(result.partial, 3)
        docs = {p: (p.read_bytes(), p.stat().st_mtime_ns) for p in result.folder.glob('*.docx')}
        def convert(path):
            target = path.with_suffix('.pdf')
            target.write_bytes(b'%PDF-test')
            return target
        with patch('core.pdf_converter', return_value=('test', '')), patch('core.PdfSession.convert', side_effect=convert):
            result = retry_batch(result, self.template, self.dataset, self.mapping, self.pattern)
        self.assertEqual(result.succeeded, 3)
        self.assertTrue(all((p.read_bytes(), p.stat().st_mtime_ns) == before for p, before in docs.items()))

    def test_changed_results_rejected(self):
        result = self.generate()
        save_batch_state(result, self.job)
        document = result.folder / result.items[0]['Word 文件']
        document.write_bytes(document.read_bytes() + b'changed')
        with self.assertRaises(InputError):
            load_batch_state(result.folder)
        with self.assertRaises(InputError):
            validate_batch_files(result)

    def test_archive_includes_only_results_and_manifest(self):
        result = self.generate()
        save_batch_state(result, self.job)
        (result.folder / 'private.txt').write_text('private', encoding='utf-8')
        target = self.folder / 'share.zip'
        archive_batch(result, target)
        with zipfile.ZipFile(target) as archive:
            expected = {i['Word 文件'] for i in result.items} | {result.manifest_path.name}
            self.assertEqual(set(archive.namelist()), expected)
            self.assertIsNone(archive.testzip())

    def test_archive_cancellation_preserves_existing_zip(self):
        result = self.generate()
        target = self.folder / 'share.zip'
        target.write_bytes(b'original')
        stop = threading.Event()
        with self.assertRaises(ArchiveCancelled):
            archive_batch(result, target, cancel=stop, progress=lambda *args: stop.set())
        self.assertEqual(target.read_bytes(), b'original')

    def test_report_exports_all_rows(self):
        issues = [{'行号': i + 2, '字段': '姓名', '数据列': '姓名', '级别': '错误', '说明': '必填字段不能为空'} for i in range(251)]
        path = self.folder / 'report.csv'
        write_check_report(path, self.folder / 'data.xlsx', 'Sheet1', issues)
        with path.open(encoding='utf-8-sig', newline='') as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual(len(rows), 251)
        self.assertEqual(rows[-1]['行号'], '252')


if __name__ == '__main__':
    unittest.main()
