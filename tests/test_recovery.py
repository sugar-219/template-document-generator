from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from core import InputError, Template, continue_batch, generate_batch, read_data, retry_batch
from batch_journal import BatchJournal, JOURNAL_FILE
from batch_state import DIRTY_FILE, STATE_DIR, load_batch_state, save_batch_state, validate_batch_files


class RecoveryTests(unittest.TestCase):
    def setUp(self):
        base = ROOT / '.build/test-temp'
        base.mkdir(parents=True, exist_ok=True)
        temporary = tempfile.TemporaryDirectory(dir=base)
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        template = Template(ROOT / '示例文件/通知模板.docx')
        dataset = read_data(ROOT / '示例文件/人员数据.xlsx')
        self.job = template, dataset, {name: name for name in template.fields}, '{{序号}}_{{姓名}}', ('姓名',)

    def crash(self, mode='generate', folder=None):
        folder = folder or self.folder
        completed = subprocess.run([sys.executable, '-B', str(ROOT / 'tests/crash_worker.py'), str(folder), mode],
                                   timeout=20, capture_output=True)
        self.assertEqual(completed.returncode, 73, completed.stderr.decode(errors='replace'))
        return Path((folder / 'fixture-folder.txt').read_text(encoding='utf-8'))

    def finish(self, loaded, job):
        journal = BatchJournal(job, loaded)
        try:
            result = continue_batch(loaded, *job[:4], checkpoint=journal)
            save_batch_state(result, job)
            journal.close(True)
        finally:
            journal.close()
        return result

    def test_forced_exit_restores_completed_record(self):
        folder = self.crash()
        loaded, job = load_batch_state(folder)
        self.assertEqual(loaded.succeeded, 1)
        self.assertEqual(loaded.remaining_indices, [1, 2])
        validate_batch_files(loaded)
        first = folder / loaded.items[0]['Word 文件']
        before = first.read_bytes(), first.stat().st_mtime_ns
        result = self.finish(loaded, job)
        self.assertEqual(result.succeeded, 3)
        self.assertEqual((first.read_bytes(), first.stat().st_mtime_ns), before)
        self.assertFalse((folder / DIRTY_FILE).exists())
        self.assertFalse((folder / STATE_DIR / JOURNAL_FILE).exists())
        self.assertEqual(load_batch_state(folder)[0].succeeded, 3)

    def test_partial_docx_is_preserved_without_overwrite(self):
        folder = self.crash('pending')
        partial = next(folder.glob('*.docx'))
        before = partial.read_bytes()
        loaded, job = load_batch_state(folder)
        self.assertEqual(len(loaded.items), 0)
        result = self.finish(loaded, job)
        self.assertEqual(result.succeeded, 3)
        self.assertEqual(partial.read_bytes(), before)
        self.assertNotIn(partial.name, [item['Word 文件'] for item in result.items])

    def test_crash_during_continuation_preserves_previous_successes(self):
        stop = threading.Event()
        result = generate_batch(*self.job[:3], self.folder, self.job[3], cancel=stop, progress=lambda *a: stop.set())
        save_batch_state(result, self.job)
        folder = self.crash('continue', result.folder)
        loaded, job = load_batch_state(folder)
        self.assertEqual(loaded.succeeded, 2)
        self.assertEqual(loaded.remaining_indices, [2])
        self.assertEqual(self.finish(loaded, job).succeeded, 3)

    def test_crash_during_retry_restores_only_completed_retry(self):
        with patch('core.pdf_converter', return_value=None):
            result = generate_batch(*self.job[:3], self.folder, self.job[3], export_pdf=True)
        save_batch_state(result, self.job)
        folder = self.crash('retry', result.folder)
        loaded, _ = load_batch_state(folder)
        self.assertEqual(loaded.succeeded, 1)
        self.assertEqual(loaded.partial, 2)
        self.assertFalse(loaded.remaining_indices)

    def test_crash_after_manifest_write_still_recovers(self):
        loaded, _ = load_batch_state(self.crash('manifest'))
        self.assertEqual(loaded.succeeded, 3)
        self.assertIsNone(loaded.manifest_path)
        self.assertTrue(loaded.manifest_warning)

    def test_unconfirmed_pdf_does_not_block_retry_or_overwrite_word(self):
        with patch('core.pdf_converter', return_value=None):
            result = generate_batch(*self.job[:3], self.folder, self.job[3], export_pdf=True)
        save_batch_state(result, self.job)
        folder = self.crash('pdf-publish', result.folder)
        orphan = next(folder.glob('*.pdf'))
        before = orphan.read_bytes()
        documents = {p: (p.read_bytes(), p.stat().st_mtime_ns) for p in folder.glob('*.docx')}
        loaded, job = load_batch_state(folder)
        def convert(path, destination=None):
            target = destination or path.with_suffix('.pdf')
            with target.open('xb') as stream:
                stream.write(b'%PDF-confirmed')
            return target
        with patch('core.pdf_converter', return_value=('test', '')), patch('core.PdfSession.convert', side_effect=convert):
            retried = retry_batch(loaded, *job[:4])
        self.assertEqual(retried.succeeded, 3)
        self.assertEqual(orphan.read_bytes(), before)
        self.assertNotIn(orphan.name, [item['PDF 文件'] for item in retried.items])
        self.assertTrue(all((p.read_bytes(), p.stat().st_mtime_ns) == before for p, before in documents.items()))

    def test_selected_record_order_and_source_rows_survive(self):
        loaded, job = load_batch_state(self.crash('selected'))
        self.assertEqual(loaded.items[0]['数据行'], '4')
        result = self.finish(loaded, job)
        self.assertEqual([item['数据行'] for item in result.items], ['4', '2'])
        self.assertEqual([item['Word 文件'][:3] for item in result.items], ['001', '002'])

    def test_torn_final_event_is_not_accepted(self):
        folder = self.crash()
        with (folder / STATE_DIR / JOURNAL_FILE).open('ab') as stream:
            stream.write(b'{"index":')
        loaded, _ = load_batch_state(folder)
        self.assertEqual(loaded.succeeded, 1)
        self.assertEqual(loaded.remaining_indices, [1, 2])

    def test_remaining_record_order_is_preserved(self):
        loaded, job = load_batch_state(self.crash('reversed'))
        result = self.finish(loaded, job)
        self.assertEqual([item['数据行'] for item in result.items], ['4', '3', '2'])

    def test_corrupt_complete_event_rejected(self):
        folder = self.crash()
        with (folder / STATE_DIR / JOURNAL_FILE).open('ab') as stream:
            stream.write(b'corrupt\n')
        with self.assertRaises(InputError):
            load_batch_state(folder)

    def test_external_change_rejected_before_and_after_recovery(self):
        folder = self.crash()
        loaded, _ = load_batch_state(folder)
        document = folder / loaded.items[0]['Word 文件']
        document.write_bytes(document.read_bytes() + b'changed')
        with self.assertRaises(InputError):
            load_batch_state(folder)
        with self.assertRaises(InputError):
            validate_batch_files(loaded)

    def test_old_unsaved_marker_without_journal_still_blocks(self):
        result = generate_batch(*self.job[:3], self.folder, self.job[3])
        save_batch_state(result, self.job)
        (result.folder / DIRTY_FILE).touch()
        with self.assertRaises(InputError):
            load_batch_state(result.folder)


if __name__ == '__main__':
    unittest.main()
