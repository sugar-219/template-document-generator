"""Subprocess fixture: exit without finally blocks at an actual persistence boundary."""
import os
from pathlib import Path
import sys
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from batch_journal import BatchJournal
from batch_state import load_batch_state
from core import Template, continue_batch, generate_batch, read_data, retry_batch
import core

folder, mode = Path(sys.argv[1]), sys.argv[2]
if mode in ('continue', 'retry', 'pdf-publish'):
    previous, job = load_batch_state(folder)
else:
    previous = None
    template = Template(ROOT / '示例文件/通知模板.docx')
    dataset = read_data(ROOT / '示例文件/人员数据.xlsx')
    job = template, dataset, {name: name for name in template.fields}, '{{序号}}_{{姓名}}', ('姓名',)
journal = BatchJournal(job, previous)

def checkpoint(result, index=None, item=None):
    journal(result, index, item)
    (folder / 'fixture-folder.txt').write_text(str(result.folder), encoding='utf-8')
    if index is not None and mode not in ('manifest', 'pending'):
        os._exit(73)

def crash_render(values, destination):
    with destination.open('xb') as stream:
        stream.write(b'incomplete DOCX')
    os._exit(73)

def crash_manifest(result):
    original_manifest(result)
    os._exit(73)

def fake_pdf(path):
    target = path.with_suffix('.pdf')
    target.write_bytes(b'%PDF-fixture')
    if mode == 'pdf-publish':
        os._exit(73)
    return target

if mode == 'continue':
    continue_batch(previous, *job[:4], required_fields=job[4], checkpoint=checkpoint)
elif mode in ('retry', 'pdf-publish'):
    with patch('core.pdf_converter', return_value=('test', '')), patch('core.PdfSession.convert', side_effect=fake_pdf):
        retry_batch(previous, *job[:4], required_fields=job[4], checkpoint=checkpoint)
elif mode == 'pending':
    with patch.object(job[0], 'render', side_effect=crash_render):
        generate_batch(*job[:3], folder, job[3], checkpoint=checkpoint)
elif mode == 'manifest':
    original_manifest = core.save_manifest
    with patch('core.save_manifest', side_effect=crash_manifest):
        generate_batch(*job[:3], folder, job[3], checkpoint=checkpoint)
else:
    selected = [2, 1, 0] if mode == 'reversed' else [2, 0] if mode == 'selected' else None
    generate_batch(*job[:3], folder, job[3], indices=selected, checkpoint=checkpoint)
