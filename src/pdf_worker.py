"""Isolated Office conversion. JSON files also work with a windowed frozen EXE."""
import ctypes
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def write_json(path, value):
    path = Path(path)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    temporary.replace(path)


def run_worker(request_path):
    request_path = Path(request_path)
    job = json.loads(request_path.read_text(encoding="utf-8"))
    answer = {"ok": False, "phase": "startup"}
    word = None
    initialized = False
    try:
        source, target = Path(job["source"]), Path(job["target"])
        if job["converter"][0] == "Word":
            import pythoncom
            import win32com.client
            import win32process
            before = set(win32process.EnumProcesses())
            pythoncom.CoInitialize()
            initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3
            scratch = word.Documents.Add()
            pid = win32process.GetWindowThreadProcessId(word.ActiveWindow.Hwnd)[1]
            write_json(request_path.parent / "office.json", {"pid": pid, "owned": pid not in before})
            scratch.Close(False)
            answer["phase"] = "convert"
            document = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            try:
                document.ExportAsFixedFormat(str(target), 17)
            finally:
                document.Close(False)
        else:
            answer["phase"] = "convert"
            profile = request_path.parent / "lo-profile"
            profile.mkdir(exist_ok=True)
            result = subprocess.run([
                job["converter"][1], f"-env:UserInstallation={profile.as_uri()}",
                "--headless", "--convert-to", "pdf", "--outdir", str(target.parent), str(source),
            ], stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            produced = target.parent / source.with_suffix(".pdf").name
            if result.returncode or not produced.is_file():
                raise RuntimeError("LibreOffice 转换失败：" + result.stderr.decode(errors="replace")[:200])
            if produced != target:
                produced.replace(target)
        if not target.is_file() or target.stat().st_size == 0:
            raise RuntimeError("转换器未生成有效 PDF。请检查 Office 是否正常可用。")
        answer["ok"] = True
    except Exception as exc:
        answer["error"] = str(exc)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()
        write_json(request_path.parent / "answer.json", answer)


def _office_handle(folder, started):
    """Pin only the newly created Word process reported by this worker."""
    if os.name != "nt":
        return None
    try:
        import pywintypes
        import win32api
        import win32process
    except ImportError:
        return None
    handle = None
    try:
        data = json.loads((folder / "office.json").read_text(encoding="utf-8"))
        if not data.get("owned"):
            return None
        handle = win32api.OpenProcess(0x1000 | 0x100000 | 0x1, False, data["pid"])
        created = win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
        buffer = ctypes.create_unicode_buffer(32768)
        size = ctypes.c_ulong(len(buffer))
        query = ctypes.windll.kernel32.QueryFullProcessImageNameW
        query.argtypes = [ctypes.c_void_p, ctypes.c_ulong, ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_ulong)]
        if created < started or not query(int(handle), 0, buffer, ctypes.byref(size)) or Path(buffer.value).name.upper() != "WINWORD.EXE":
            handle.Close()
            return None
        return handle
    except (OSError, ValueError, KeyError, pywintypes.error):
        if handle is not None:
            handle.Close()
        return None


def _stop_worker(process, office):
    if process.poll() is None:
        if os.name == "nt":
            try:
                subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            except (OSError, subprocess.TimeoutExpired):
                pass
        if process.poll() is None:
            process.kill()
        process.wait(timeout=5)
    if office is not None:
        import pywintypes
        import win32api
        import win32event
        try:
            win32api.TerminateProcess(office, 1)
            win32event.WaitForSingleObject(office, 5000)
        except (OSError, pywintypes.error):
            pass


class ConversionCancelled(RuntimeError):
    pass


class PdfSession:
    def __init__(self, converter, cancel=None, timeout=60):
        self.converter, self.cancel, self.timeout = converter, cancel, timeout
        self.startup_error = None

    def __enter__(self):
        return self

    def convert(self, docx, destination=None):
        docx = Path(docx).resolve()
        target = Path(destination).resolve() if destination is not None else docx.with_suffix(".pdf")
        if target.exists():
            raise FileExistsError(f"PDF 已存在，不会覆盖：{target.name}")
        if self.startup_error:
            raise RuntimeError(self.startup_error)
        if self.cancel and self.cancel.is_set():
            raise ConversionCancelled("PDF 转换已取消，Word 文档已保留。")
        with tempfile.TemporaryDirectory(prefix="template-pdf-") as temporary:
            folder = Path(temporary)
            request = folder / "request.json"
            job = {"source": str(docx), "target": str(folder / "converted.pdf"), "converter": self.converter}
            write_json(request, job)
            command = [sys.executable]
            if not getattr(sys, "frozen", False):
                command.append(str(Path(__file__).resolve()))
            command += ["--pdf-worker", str(request)]
            started = time.time()
            process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                       stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            office = None
            deadline = time.monotonic() + self.timeout
            try:
                while process.poll() is None:
                    if office is None:
                        office = _office_handle(folder, started)
                    if self.cancel and self.cancel.is_set():
                        raise ConversionCancelled("PDF 转换已取消，Word 文档已保留。")
                    if time.monotonic() >= deadline:
                        raise TimeoutError(f"PDF 转换超过 {self.timeout:g} 秒，已停止；Word 文档已保留。")
                    try:
                        process.wait(timeout=0.1)
                    except subprocess.TimeoutExpired:
                        pass
                answer_path = folder / "answer.json"
                if not answer_path.is_file():
                    raise RuntimeError(f"PDF 转换进程异常退出（{process.returncode}）。")
                answer = json.loads(answer_path.read_text(encoding="utf-8"))
                if not answer.get("ok"):
                    if answer.get("phase") == "startup":
                        self.startup_error = "转换器启动失败：" + answer.get("error", "未知错误")
                    raise RuntimeError(self.startup_error or answer.get("error", "PDF 转换失败。"))
                # Publish only a fully converted file, using exclusive creation.
                destination = target.open("xb")
                try:
                    with Path(job["target"]).open("rb") as source, destination:
                        import shutil
                        shutil.copyfileobj(source, destination)
                except Exception:
                    destination.close()
                    target.unlink(missing_ok=True)
                    raise
                return target
            except BaseException:
                if office is None:
                    office = _office_handle(folder, started)
                _stop_worker(process, office)
                raise
            finally:
                if office is not None:
                    office.Close()

    def __exit__(self, *args):
        return False


if __name__ == "__main__":
    run_worker(sys.argv[sys.argv.index("--pdf-worker") + 1])
