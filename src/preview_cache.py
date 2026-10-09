"""Each app instance owns a preview directory until its process releases the lock."""
import re
import shutil
import time
import uuid
from pathlib import Path

from local_lock import FileLock


class PreviewCache:
    def __init__(self, base):
        self.base = Path(base)
        self.path = self.base / ("instance-" + uuid.uuid4().hex)
        self.lease = None

    def ensure(self):
        if self.lease is None:
            self.path.mkdir(parents=True, exist_ok=True)
            self.lease = FileLock(self.path / ".owner.lock", timeout=0).__enter__()

    def cleanup_stale(self):
        if not self.base.is_dir():
            return
        for folder in self.base.iterdir():
            if folder == self.path or not re.fullmatch(r"instance-[0-9a-f]{32}", folder.name):
                continue
            if folder.is_symlink() or folder.is_junction() or folder.resolve().parent != self.base.resolve():
                continue
            try:
                if not folder.is_dir() or time.time() - folder.stat().st_mtime < 30:
                    continue
                if any(path.name.startswith("~$") for path in folder.rglob("*")):
                    continue
                # A running owner holds this byte lock; process exit releases it.
                with FileLock(folder / ".owner.lock", timeout=0):
                    pass
                shutil.rmtree(folder)
            except OSError:
                pass

    def close(self):
        if self.lease is not None:
            self.lease.close()
            self.lease = None
        try:
            if all(path.name == ".owner.lock" for path in self.path.iterdir()):
                (self.path / ".owner.lock").unlink(missing_ok=True)
                self.path.rmdir()
        except OSError:
            pass

