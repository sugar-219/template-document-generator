"""Preserve the notices shipped with runtime dependencies."""
from importlib.metadata import distribution
from pathlib import Path
import shutil
import sys

ROOT = Path(__file__).resolve().parent.parent
destination = ROOT / "桌面版/第三方许可"
destination.mkdir(parents=True, exist_ok=True)
for name in ("openpyxl", "et-xmlfile", "lxml", "pywin32", "pyinstaller", "Pillow", "pypdfium2"):
    package = distribution(name)
    for relative in package.files or []:
        if any(term in relative.as_posix().lower() for term in ("license", "licence", "copying", "notice")):
            source = Path(package.locate_file(relative))
            if source.is_file():
                target = destination / name / Path(*relative.parts)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
python_license = Path(sys.base_prefix) / "LICENSE.txt"
if python_license.is_file():
    shutil.copy2(python_license, destination / "Python-LICENSE.txt")
