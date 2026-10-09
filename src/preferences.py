"""Small local preferences file; business records are never persisted here."""
import json
import os
import copy
import tempfile
from pathlib import Path

from local_lock import FileLock


class Preferences:
    def __init__(self, path=None):
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "TemplateBatch"
        self.path = Path(path) if path else base / "preferences.json"
        try:
            self.data = self._read()
        except OSError:
            self.data = {}
        self.local_data = copy.deepcopy(self.data)
        self.local_data.setdefault("animations", True)

    def _read(self):
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (FileNotFoundError, ValueError):
            return {}

    def profile(self, fingerprint):
        profiles = self.data.get("templates", {})
        profile = profiles.get(fingerprint, {}) if isinstance(profiles, dict) else {}
        return profile if isinstance(profile, dict) else {}

    def recent_files(self, kind):
        histories = self._read().get("recent_files", {})
        entries = histories.get(kind, []) if isinstance(histories, dict) else []
        return [path for path in entries if isinstance(path, str) and path][:5] if isinstance(entries, list) else []

    def save(self, output, pattern, pdf, fingerprint=None, mapping=None, recent_file=None, clear_recent=False, animations=None):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        requested = {"output": output, "pattern": pattern, "pdf": bool(pdf)}
        if animations is not None:
            requested["animations"] = bool(animations)
        mapping = dict(mapping or {})
        missing = object()
        with FileLock(self.path.with_suffix(".lock")):
            data = self._read()
            if clear_recent:
                data.pop("recent_files", None)
            elif recent_file:
                kind, source = recent_file
                path = str(Path(source).resolve())
                histories = data.get("recent_files", {})
                histories = dict(histories) if isinstance(histories, dict) else {}
                entries = histories.get(kind, [])
                entries = entries if isinstance(entries, list) else []
                histories[kind] = [path, *[entry for entry in entries if isinstance(entry, str) and entry
                                         and os.path.normcase(entry) != os.path.normcase(path)]][:5]
                data["recent_files"] = histories
            for key, value in requested.items():
                if self.local_data.get(key, missing) != value:
                    data[key] = value
            if fingerprint:
                previous = self.local_data.get("templates", {})
                previous = previous.get(fingerprint, {}) if isinstance(previous, dict) else {}
                previous = previous if isinstance(previous, dict) else {}
                old_mapping = previous.get("mapping", {})
                old_mapping = old_mapping if isinstance(old_mapping, dict) else {}
                profiles = data.get("templates", {})
                if not isinstance(profiles, dict):
                    profiles = {}
                current = profiles.get(fingerprint, {})
                current = current if isinstance(current, dict) else {}
                merged = current.get("mapping", {})
                merged = dict(merged) if isinstance(merged, dict) else {}
                if previous.get("pattern", missing) != pattern:
                    current["pattern"] = pattern
                for name in old_mapping.keys() | mapping.keys():
                    if old_mapping.get(name, missing) != mapping.get(name, missing):
                        if name in mapping:
                            merged[name] = mapping[name]
                        else:
                            merged.pop(name, None)
                current["mapping"] = merged
                profiles[fingerprint] = current
                data["templates"] = profiles
            temporary = None
            try:
                with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent,
                                                 prefix=".preferences-", suffix=".tmp", delete=False) as stream:
                    temporary = Path(stream.name)
                    stream.write(json.dumps(data, ensure_ascii=False, indent=2))
                temporary.replace(self.path)
            finally:
                if temporary:
                    temporary.unlink(missing_ok=True)
            self.data = data
            self.local_data.update(requested)
            if fingerprint:
                if not isinstance(self.local_data.get("templates"), dict):
                    self.local_data["templates"] = {}
                self.local_data["templates"][fingerprint] = {"pattern": pattern, "mapping": mapping}
