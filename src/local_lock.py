"""Windows process locks for local settings and preview ownership."""
import msvcrt
import time
from pathlib import Path


class FileLock:
    def __init__(self, path, timeout=2):
        self.path = Path(path)
        self.timeout = timeout
        self.stream = None

    def __enter__(self):
        stream = self.path.open("a+b")
        deadline = time.monotonic() + self.timeout
        try:
            while True:
                stream.seek(0)
                try:
                    msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError:
                    if time.monotonic() >= deadline:
                        raise
                    time.sleep(0.02)
        except BaseException:
            stream.close()
            raise
        self.stream = stream
        return self

    def close(self):
        if self.stream is not None:
            try:
                self.stream.seek(0)
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
            finally:
                self.stream.close()
                self.stream = None

    def __exit__(self, *args):
        self.close()

