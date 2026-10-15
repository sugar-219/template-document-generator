"""Archive only the portable application and verify every archived file."""
import hashlib
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
release = ROOT / "桌面版"
archive_path = ROOT / "模板文档批量生成软件_v1.14.zip"
files = sorted(path for path in release.rglob("*") if path.is_file())
assert (release / "模板文档批量生成.exe").is_file()
assert (release / "_internal/python313.dll").is_file()
with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
    for path in files:
        archive.write(path, "模板文档批量生成软件/" + path.relative_to(release).as_posix())
with zipfile.ZipFile(archive_path) as archive:
    assert archive.testzip() is None
    assert len(archive.namelist()) == len(files)
    for path in files:
        name = "模板文档批量生成软件/" + path.relative_to(release).as_posix()
        assert hashlib.sha256(path.read_bytes()).digest() == hashlib.sha256(archive.read(name)).digest(), name
    archive.extractall(ROOT / ".build/portable")
summary = {
    "files": len(files),
    "release_bytes": sum(path.stat().st_size for path in files),
    "zip_bytes": archive_path.stat().st_size,
    "zip_sha256": hashlib.sha256(archive_path.read_bytes()).hexdigest(),
    "exe_sha256": hashlib.sha256((release / "模板文档批量生成.exe").read_bytes()).hexdigest(),
}
(ROOT / ".build/archive-test.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False))
