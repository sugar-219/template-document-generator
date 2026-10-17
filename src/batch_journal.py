"""Durable completed-record events over an unchanged batch snapshot."""
import json
import os
from pathlib import Path

from batch_state import DIRTY_FILE, STATE_DIR, _batch_files, _file, save_batch_state
from core import InputError
from local_lock import FileLock

JOURNAL_FILE = '进度.jsonl'


class BatchJournal:
    def __init__(self, job, previous=None):
        self.job = job
        self.previous = previous
        self.folder = None
        self.lease = None
        self.lookup = {}

    def __call__(self, result, index=None, item=None):
        if self.folder is None:
            self.folder = result.folder.resolve()
            if self.previous is None:
                self.lease = FileLock(self.folder / '.批次操作.lock', timeout=0).__enter__()
            wanted = sorted(set(result.indices) | set(result.remaining_indices))
            self.lookup = {old: new for new, old in enumerate(wanted)}
            save_batch_state(result, self.job)
            path = _file(self.folder / STATE_DIR, JOURNAL_FILE)
            temporary = path.with_suffix('.tmp')
            try:
                with temporary.open('wb') as stream:
                    stream.write((json.dumps({'schema': 1, 'base': result.state_token}) + '\n').encode('utf-8'))
                    stream.flush()
                    os.fsync(stream.fileno())
                temporary.replace(path)
            finally:
                temporary.unlink(missing_ok=True)
            marker = _file(self.folder, DIRTY_FILE)
            marker.touch()
        if index is not None:
            # Hash only this record, rather than all previously completed documents.
            from core import BatchResult
            files = _batch_files(BatchResult(self.folder, [item], False), self.folder.resolve())
            event = {'index': self.lookup[index], 'item': item, 'files': files}
            path = _file(self.folder / STATE_DIR, JOURNAL_FILE)
            with path.open('ab') as stream:
                stream.write((json.dumps(event, ensure_ascii=False, separators=(',', ':')) + '\n').encode('utf-8'))
                stream.flush()
                os.fsync(stream.fileno())

    def close(self, saved=False):
        try:
            if saved and self.folder is not None:
                try:
                    _file(self.folder / STATE_DIR, JOURNAL_FILE).unlink(missing_ok=True)
                except OSError:
                    pass  # A completed state ignores a leftover journal without its dirty marker.
        finally:
            if self.lease is not None:
                self.lease.close()
                self.lease = None


def replay_journal(folder, data, token):
    """Ignore only a torn final line; completed corrupt events fail closed."""
    path = _file(folder / STATE_DIR, JOURNAL_FILE)
    if not path.is_file():
        raise InputError('该批次上次任务状态未保存，请在原窗口重试保存；已生成文件保留。')
    try:
        lines = path.read_bytes().splitlines(keepends=True)
        if not lines or not lines[0].endswith(b'\n'):
            raise ValueError('进度记录不完整')
        header = json.loads(lines[0])
        if header != {'schema': 1, 'base': token}:
            raise ValueError('进度记录与批次不匹配')
        saved = data['result']
        items, indices = saved['items'], saved['indices']
        remaining = set(saved['remaining_indices'])
        slots = {index: slot for slot, index in enumerate(indices)}
        # A manifest interrupted during replacement is rebuilt from verified results.
        manifest = saved['manifest']
        expected = dict(data['files'])
        if manifest is not None:
            expected.pop(manifest, None)
        stems = set(saved['used_stems'])
        for line in lines[1:]:
            if not line.endswith(b'\n'):
                break
            event = json.loads(line)
            index, item, files = event['index'], event['item'], event['files']
            if (type(index) is not int or not 0 <= index < len(data['dataset']['records'])
                    or not isinstance(item, dict) or not isinstance(files, dict)
                    or any(not isinstance(name, str) or not isinstance(digest, str) for name, digest in files.items())):
                raise ValueError('进度记录无效')
            names = {item[key] for key in ('Word 文件', 'PDF 文件') if item[key]}
            if set(files) != names:
                raise ValueError('进度文件校验无效')
            if index in slots:
                old = items[slots[index]]
                for key in ('Word 文件', 'PDF 文件'):
                    if old[key]:
                        expected.pop(old[key], None)
                items[slots[index]] = item
            elif index in remaining:
                slots[index] = len(items)
                indices.append(index)
                items.append(item)
                remaining.remove(index)
            else:
                raise ValueError('进度记录顺序无效')
            expected.update(files)
            stems.update(Path(name).stem.casefold() for name in names)
        saved.update(remaining_indices=[index for index in saved['remaining_indices'] if index in remaining],
                     used_stems=sorted(stems), cancelled=True,
                     manifest=None, manifest_warning='上次操作意外中断，继续或重试后将重新保存生成清单。')
        data['files'] = expected
    except (KeyError, TypeError, ValueError, UnicodeError) as exc:
        raise InputError('批次进度记录损坏或不匹配，请保留原文件后检查，不能直接续作。') from exc
    return data
