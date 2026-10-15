"""Persist completed/stopped batches using local template and selected-record snapshots."""
from __future__ import annotations

import hashlib
import json
import tempfile
import zipfile
from pathlib import Path

from core import BatchResult, Dataset, InputError, Template, file_stem, values_for
from local_lock import FileLock

STATE_DIR, STATE_FILE, DIRTY_FILE = ".任务状态", "状态.json", ".任务状态未保存"


def _digest(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _folder(folder):
    folder = Path(folder)
    if not folder.is_dir() or folder.is_symlink() or folder.is_junction():
        raise InputError("批次目录不存在或使用了链接，请选择原批次文件夹。")
    folder = folder.resolve()
    private = folder / STATE_DIR
    if private.is_symlink() or private.is_junction():
        raise InputError("批次任务状态目录存在链接，无法安全读取或保存。")
    return folder, private


def _file(folder, name):
    if (not isinstance(name, str) or not name or name in (".", "..")
            or any(char in name for char in ("/", "\\", ":", "\0")) or Path(name).name != name):
        raise InputError("批次任务状态包含无效文件名。")
    path = folder / name
    if path.is_symlink() or path.is_junction() or path.resolve().parent != folder:
        raise InputError("批次文件存在链接或路径越界：" + name)
    return path


def state_token(folder):
    _, private = _folder(folder)
    path = _file(private, STATE_FILE)
    return _digest(path) if path.is_file() else ""


def _batch_files(result, folder):
    files, seen = {}, set()
    for item in result.items:
        for key, suffix in (("Word 文件", ".docx"), ("PDF 文件", ".pdf")):
            name = item[key]
            if not name:
                continue
            path = _file(folder, name)
            if path.suffix.lower() != suffix or not path.is_file():
                raise InputError("批次结果文件缺失或格式不符：" + name)
            if name.casefold() in seen:
                raise InputError("批次结果文件名重复。")
            seen.add(name.casefold())
            files[name] = _digest(path)
    if result.manifest_path is not None:
        manifest = Path(result.manifest_path)
        if manifest.resolve().parent != folder or manifest.suffix.lower() != ".csv":
            raise InputError("批次清单路径无效。")
        path = _file(folder, manifest.name)
        if not path.is_file():
            raise InputError("批次生成清单已丢失。")
        files[path.name] = _digest(path)
    return files


def validate_batch_files(result):
    folder, private = _folder(result.folder)
    if (folder / DIRTY_FILE).exists():
        raise InputError("该批次有未保存的操作，请在原窗口重试保存后重新载入。已生成文件保留。")
    encoded = _file(private, STATE_FILE).read_bytes()
    if hashlib.sha256(encoded).hexdigest() != result.state_token:
        raise InputError("批次状态已变化，请重新载入历史批次后再操作。")
    try:
        expected = json.loads(encoded)["files"]
    except (KeyError, TypeError, ValueError) as exc:
        raise InputError("批次任务状态格式无效，请恢复原文件后再操作。") from exc
    if not isinstance(expected, dict) or _batch_files(result, folder) != expected:
        raise InputError("批次文档或清单已改变，请恢复原文件后再操作。当前任务未改动。")


def save_batch_state(result, job):
    folder, private = _folder(result.folder)
    template, dataset, mapping, pattern, required = job
    wanted = sorted(set(result.indices) | set(result.remaining_indices))
    if (len(result.indices) != len(result.items) or not wanted
            or any(type(index) is not int or not 0 <= index < len(dataset.records) for index in wanted)):
        raise InputError("批次记录索引无效，无法保存任务状态。")
    lookup = {old: new for new, old in enumerate(wanted)}
    private.mkdir(exist_ok=True)
    temporary = created_template = None
    with FileLock(private / ".写入.lock"):
        try:
            with tempfile.NamedTemporaryFile(dir=private, prefix=".template-", suffix=".docx", delete=False) as stream:
                temporary = Path(stream.name)
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
                for entry, content in template.entries:
                    archive.writestr(entry, content)
            template_sha = _digest(temporary)
            snapshot = _file(private, "模板_" + template_sha + ".docx")
            if snapshot.exists():
                if not snapshot.is_file() or _digest(snapshot) != template_sha:
                    raise InputError("批次模板快照已损坏，请检查批次目录。")
                temporary.unlink()
            else:
                temporary.replace(snapshot)
                created_template = snapshot
            temporary = None
            payload = {
                "schema": 1, "template": snapshot.name, "template_sha256": template_sha,
                "dataset": {"headers": dataset.headers, "records": [dataset.records[index] for index in wanted],
                            "row_numbers": [dataset.row_numbers[index] for index in wanted], "sheet": dataset.sheet},
                "mapping": mapping, "pattern": pattern, "required_fields": list(required),
                "result": {"items": result.items, "indices": [lookup[index] for index in result.indices],
                           "remaining_indices": [lookup[index] for index in result.remaining_indices],
                           "cancelled": result.cancelled, "export_pdf": result.export_pdf,
                           "used_stems": sorted(result.used_stems),
                           "manifest": Path(result.manifest_path).name if result.manifest_path else None,
                           "manifest_warning": result.manifest_warning},
                "files": _batch_files(result, folder),
            }
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=private, prefix=".state-", suffix=".tmp", delete=False) as stream:
                temporary = Path(stream.name)
                json.dump(payload, stream, ensure_ascii=False, separators=(",", ":"))
            target = _file(private, STATE_FILE)
            temporary.replace(target)
            temporary = created_template = None
            marker = folder / DIRTY_FILE
            if marker.is_symlink() or marker.is_junction():
                raise InputError("批次未保存标记存在链接，请检查批次目录。")
            marker.unlink(missing_ok=True)
            result.state_token = _digest(target)
            return target
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
            if created_template is not None:
                created_template.unlink(missing_ok=True)


def load_batch_state(folder):
    folder, private = _folder(folder)
    if (folder / DIRTY_FILE).exists():
        raise InputError("该批次上次任务状态未保存，请在原窗口重试保存；已生成文件保留。")
    path = _file(private, STATE_FILE)
    if not path.is_file():
        raise InputError("未找到可恢复的任务状态，请选择 v1.11 或更新版本完成/停止后的批次文件夹。")
    try:
        encoded = path.read_bytes()
        data = json.loads(encoded)
        if not isinstance(data, dict) or data.get("schema") != 1:
            raise ValueError("版本无效")
        template_path = _file(private, data["template"])
        if (not isinstance(data["template_sha256"], str) or not template_path.is_file()
                or _digest(template_path) != data["template_sha256"]):
            raise InputError("批次模板快照已丢失或改变，请恢复原文件后再载入。")
        template = Template(template_path)
        raw, saved = data["dataset"], data["result"]
        headers, records, rows, sheet = raw["headers"], raw["records"], raw["row_numbers"], raw["sheet"]
        if (not isinstance(headers, list) or not headers or any(not isinstance(h, str) or not h for h in headers)
                or len(headers) != len(set(headers)) or not isinstance(records, list) or not records
                or any(not isinstance(row, dict) or set(row) != set(headers) or any(not isinstance(v, str) for v in row.values()) for row in records)
                or not isinstance(rows, list) or len(rows) != len(records) or any(type(row) is not int or row < 2 for row in rows)
                or len(rows) != len(set(rows)) or not isinstance(sheet, str) or not sheet):
            raise ValueError("数据快照无效")
        dataset = Dataset(headers, records, rows, sheet)
        mapping, pattern, required = data["mapping"], data["pattern"], data["required_fields"]
        if (not isinstance(mapping, dict) or any(not isinstance(k, str) or not isinstance(v, str) for k,v in mapping.items())
                or not isinstance(pattern, str) or not pattern.strip() or not isinstance(required, list)
                or any(not isinstance(field, str) or field not in template.fields for field in required)):
            raise ValueError("任务设置无效")
        file_stem(pattern, values_for(template, mapping, records[0]), 1)
        items, indices, remaining = saved["items"], saved["indices"], saved["remaining_indices"]
        if (not isinstance(items, list) or not isinstance(indices, list) or not isinstance(remaining, list)
                or len(items) != len(indices) or any(type(index) is not int or not 0 <= index < len(records) for index in indices+remaining)
                or len(set(indices+remaining)) != len(records) or len(indices+remaining) != len(records)
                or any(not isinstance(item, dict) or set(item) != {"数据行", "Word 文件", "PDF 文件", "状态", "说明"}
                       or any(not isinstance(value, str) for value in item.values()) for item in items)
                or any(item["数据行"] != str(rows[index]) or item["状态"] not in ("成功", "部分成功", "失败") for item,index in zip(items,indices))
                or type(saved["cancelled"]) is not bool or type(saved["export_pdf"]) is not bool
                or not isinstance(saved["used_stems"], list) or any(not isinstance(stem, str) for stem in saved["used_stems"])
                or not isinstance(saved["manifest_warning"], str)):
            raise ValueError("批次处理状态无效")
        manifest_name = saved["manifest"]
        manifest = _file(folder, manifest_name) if manifest_name is not None else None
        result = BatchResult(folder, items, saved["cancelled"], indices, saved["export_pdf"], manifest,
                             saved["manifest_warning"], remaining, set(saved["used_stems"]))
        actual = _batch_files(result, folder)
        if not isinstance(data["files"], dict) or actual != data["files"]:
            raise InputError("批次结果或清单已改变，请恢复原文件后再载入，当前任务未改动。")
        if state_token(folder) != hashlib.sha256(encoded).hexdigest():
            raise InputError("读取时批次状态发生变化，请重新载入。")
        result.state_token = hashlib.sha256(encoded).hexdigest()
        return result, (template, dataset, mapping, pattern, tuple(required))
    except (KeyError, TypeError, ValueError, zipfile.BadZipFile) as exc:
        if isinstance(exc, InputError):
            raise
        raise InputError("批次任务状态格式无效，请选择原始完整的批次文件夹。") from exc
