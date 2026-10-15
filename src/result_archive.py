"""Archive only the generated files referenced by one batch and its latest manifest."""
from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path


class ArchiveCancelled(RuntimeError):
    """The user stopped an archive before it replaced the destination."""


def _check_cancel(cancel):
    if cancel is not None and cancel.is_set():
        raise ArchiveCancelled("打包已停止，已有 ZIP 文件未修改。")


def _signature(stat):
    return stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns


def archive_batch(result, destination, cancel=None, progress=None) -> Path:
    """Write an atomic ZIP; progress(done, total, filename) counts finished files."""
    _check_cancel(cancel)
    destination = Path(destination)
    if destination.suffix.lower() != ".zip":
        raise ValueError("打包文件必须使用 .zip 扩展名。")
    batch_folder = Path(result.folder)
    if not batch_folder.is_dir():
        raise ValueError("本批次输出目录不存在，请检查后重试。")
    if batch_folder.is_symlink() or batch_folder.is_junction():
        raise ValueError("本批次输出目录存在链接，无法安全打包。")
    folder = batch_folder.resolve(strict=True)
    names = []
    for item in result.items:
        for column, suffix in (("Word 文件", ".docx"), ("PDF 文件", ".pdf")):
            name = item.get(column, "")
            if not name:
                continue
            if (not isinstance(name, str) or ":" in name or "\0" in name
                    or Path(name).name != name or Path(name).suffix.lower() != suffix):
                raise ValueError(f"本批次的 {column} 路径无效，无法打包。")
            names.append(name)
    if not names:
        raise ValueError("本批次没有已生成的 Word 或 PDF 文件可打包。")
    if result.manifest_path is None:
        raise ValueError("本批次生成清单尚未保存，请先重新保存清单。")
    manifest = Path(result.manifest_path)
    if manifest.is_symlink() or manifest.resolve().parent != folder:
        raise ValueError("本批次生成清单位于输出目录之外，无法打包。")
    if ":" in manifest.name or manifest.suffix.lower() != ".csv":
        raise ValueError("本批次生成清单必须是 CSV 文件。")
    names.append(manifest.name)
    if len({name.casefold() for name in names}) != len(names):
        raise ValueError("本批次存在重复文件名，无法安全打包。")
    sources = []
    for name in names:
        source = folder / name
        if source.is_symlink() or source.resolve().parent != folder:
            raise ValueError(f"本批次文件存在链接或路径越界，无法打包：{name}")
        if not source.is_file():
            raise ValueError(f"本批次文件已丢失，请恢复后再打包：{name}")
        if destination.exists() and os.path.samefile(destination, source):
            raise ValueError(f"ZIP 保存位置不能覆盖本批次源文件：{name}")
        sources.append((source, _signature(source.stat())))
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=destination.parent, prefix=".batch-archive-", suffix=".tmp",
                                         delete=False) as stream:
            temporary = Path(stream.name)
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for done, (source, signature) in enumerate(sources, 1):
                _check_cancel(cancel)
                if source.is_symlink() or source.resolve().parent != folder:
                    raise ValueError(f"打包期间文件路径发生变化，请重试：{source.name}")
                info = zipfile.ZipInfo.from_file(source, source.name, strict_timestamps=False)
                info.compress_type = zipfile.ZIP_DEFLATED
                with source.open("rb") as reader, archive.open(info, "w", force_zip64=True) as writer:
                    if _signature(os.fstat(reader.fileno())) != signature:
                        raise ValueError(f"打包期间文件发生变化，请重试：{source.name}")
                    while True:
                        _check_cancel(cancel)
                        block = reader.read(1024 * 1024)
                        if not block:
                            break
                        writer.write(block)
                    if (_signature(os.fstat(reader.fileno())) != signature
                            or source.is_symlink() or _signature(source.stat()) != signature):
                        raise ValueError(f"打包期间文件发生变化，请重试：{source.name}")
                if progress:
                    progress(done, len(sources), source.name)
        _check_cancel(cancel)
        for source, signature in sources:
            if (not source.is_file() or source.is_symlink() or source.resolve().parent != folder
                    or _signature(source.stat()) != signature):
                raise ValueError(f"打包期间文件发生变化，请重试：{source.name}")
        temporary.replace(destination)
        return destination
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
