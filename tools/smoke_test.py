"""Repeatable desktop checks, separate from the application UI."""
import ctypes
import json
import os
from pathlib import Path
import sys
import threading
import time
import traceback
from tkinter import filedialog, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from core import Template, check_records, generate_batch, preview_filenames, read_data, sheet_names
from batch_state import load_batch_state
from app import resource
from preferences import Preferences
from core import Dataset

def verify_cell_menu(app, tree, row, column):
    """Exercise a Windows menu and dismiss its native message loop in UI checks."""
    from ctypes import wintypes
    class MenuInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
                    ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND), ("hwndCapture", wintypes.HWND),
                    ("hwndMenuOwner", wintypes.HWND), ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
                    ("rcCaret", wintypes.RECT)]
    api = ctypes.windll.user32
    thread_id = api.GetWindowThreadProcessId(app.winfo_id(), None)
    shown = threading.Event()
    def dismiss():
        for _ in range(100):
            info = MenuInfo()
            info.cbSize = ctypes.sizeof(info)
            if api.GetGUIThreadInfo(thread_id, ctypes.byref(info)) and info.hwndMenuOwner:
                if api.GetWindowThreadProcessId(info.hwndMenuOwner, None) == thread_id:
                    shown.set()
                    api.PostMessageW(info.hwndMenuOwner, 0x0100, 0x1B, 0)
                    return
            time.sleep(0.02)
    x, y, width, height = tree.bbox(row, column)
    x, y = x + width // 2, y + height // 2
    helper = threading.Thread(target=dismiss, daemon=True)
    helper.start()
    tree.event_generate("<Button-3>", x=x, y=y, rootx=tree.winfo_rootx()+x, rooty=tree.winfo_rooty()+y)
    helper.join(timeout=3)
    assert shown.is_set(), "原生右键菜单未显示"
    assert tree.cell_menu.entrycget(0, "label") == "查看完整内容"
    tree.cell_menu.invoke(0)
    return app.cell_window


def verify_file_workflow(app, folder):
    """Verify refreshed input and recent-file actions in source and packaged UI checks."""
    def wait_idle():
        deadline = time.monotonic() + 10
        while app.busy:
            assert time.monotonic() < deadline, app.status.get()
            app.update()
            time.sleep(0.01)
        app.update()
    source = folder / "refresh-source.csv"
    source.write_text(resource("示例文件/人员数据.csv").read_text(encoding="utf-8-sig"), encoding="utf-8-sig")
    app._data_loaded((source, ["CSV"], read_data(source)))
    original = app.dataset
    app.mapping["姓名"].set("部门")
    app.data_tree.selection_set("0", "1")
    app.update()
    source.write_text(source.read_text(encoding="utf-8-sig").replace("林晓", "刷新验收姓名"), encoding="utf-8-sig")
    app.refresh_button.invoke()
    wait_idle()
    assert app.dataset.records[0]["姓名"] == "刷新验收姓名"
    assert original.records[0]["姓名"] == "林晓"
    assert app._mapping()["姓名"] == "部门"
    assert not app.selected_indices and not app.data_tree.selection()
    assert "已刷新" in app.status.get()
    window = app.open_recent_files()
    app.update()
    data_item = next(iid for iid in window.file_tree.get_children() if window.file_tree.set(iid, "path") == str(source))
    window.file_tree.selection_set(data_item)
    app.update()
    window.open_button.invoke()
    wait_idle()
    assert app.data_path == source and app.dataset.records[0]["姓名"] == "刷新验收姓名"
    window = app.open_recent_files()
    template_item = next(iid for iid in window.file_tree.get_children() if window.file_tree.set(iid, "kind") == "模板")
    window.file_tree.selection_set(template_item)
    app.update()
    window.open_button.invoke()
    wait_idle()
    window = app.open_recent_files()
    window.clear_button.invoke()
    assert not window.file_tree.get_children()
    assert not Preferences(app.preferences.path).recent_files("data")
    assert not Preferences(app.preferences.path).recent_files("template")
    assert source.is_file()
    window.destroy()
    assert app.save_task_profile("发布验收方案")
    source.write_text(source.read_text(encoding="utf-8-sig").replace("刷新验收姓名", "方案重新读取姓名"), encoding="utf-8-sig")
    app._check_source_changes()
    assert "数据" in app.source_changes and app.generate_button.instate(["disabled"])
    app.start_generation()
    assert not app.busy
    window = app.open_task_profiles()
    window.profile_tree.selection_set("发布验收方案")
    app.update()
    window.load_button.invoke()
    wait_idle()
    assert app.dataset.records[0]["姓名"] == "方案重新读取姓名"
    assert app._mapping()["姓名"] == "部门"
    assert not app.source_changes and app.scope.get() == "全部记录"
    window.destroy()
    source = resource("示例文件/人员数据.xlsx")
    app._data_loaded((source, sheet_names(source), read_data(source)))
    app.mapping["姓名"].set("姓名")
    app._save_preferences()
    return {"refresh_reads_saved_file_and_keeps_mapping": True, "refresh_clears_old_selection": True,
            "recent_template_and_data_reopened": True, "recent_history_cleared_without_deleting_files": True,
            "changed_source_blocks_new_batch": True, "task_profile_load_reads_current_file_and_keeps_mapping": True}


def verify_check_report(app, folder):
    """Export genuine multi-page check results through the desktop controls."""
    import csv
    from unittest.mock import patch
    def wait_idle():
        deadline = time.monotonic() + 10
        while app.busy:
            assert time.monotonic() < deadline, app.status.get()
            app.update()
            time.sleep(0.01)
        app.update()
    source = folder / "check-source.csv"
    sample = read_data(resource("示例文件/人员数据.xlsx"))
    with source.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=sample.headers)
        writer.writeheader()
        for index in range(251):
            writer.writerow(dict(sample.records[0], 姓名=f"报告验证{index:04d}", 备注=""))
            writer.writerow({})
    app._data_loaded((source, ["CSV"], read_data(source)))
    app.check_data()
    wait_idle()
    assert len(app.issues) == 251, app.issues[:3]
    app._change_issue_page(1)
    assert len(app.check_tree.get_children()) == 51
    target = folder / "数据检查报告.csv"
    with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
        app.export_report_button.invoke()
        wait_idle()
    assert target.read_bytes().startswith(b"\xef\xbb\xbf")
    with target.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    assert len(rows) == 251 and rows[-1]["行号"] == "502"
    assert rows[-1]["说明"] == app.issues[-1]["说明"]
    assert all(row["数据文件"] == str(source) and row["工作表"] == "CSV" for row in rows)
    source = resource("示例文件/人员数据.xlsx")
    app._data_loaded((source, sheet_names(source), read_data(source)))
    assert app.export_report_button.instate(["disabled"])
    return {"actual_check_report_export": True, "all_251_issues_exported": True,
            "source_row_numbers_and_context_correct": True, "report_cleared_after_data_reload": True}


def verify_delivery_workflow(app, folder, history=None):
    from unittest.mock import patch
    import hashlib
    import zipfile
    import threading

    def wait_idle():
        deadline = time.monotonic() + 120
        app.update()
        while app.busy:
            assert time.monotonic() < deadline, app.status.get()
            app.update()
            time.sleep(.01)
        app.update()

    errors = []
    with patch.object(messagebox, "showerror", side_effect=lambda *a, **k: errors.append(a)):
        if history:
            assert app.template is None and app.dataset is None
            app.load_history_batch(history)
            wait_idle()
            assert app.last_batch and app.last_batch.remaining_indices == [1, 2], errors
            assert "姓名" in app._job_snapshot()[4]
            first = app.last_batch.folder / app.last_batch.items[0]["Word 文件"]
            original = first.read_bytes(), first.stat().st_mtime_ns
            app.start_continue()
            wait_idle()
            assert app.last_batch.succeeded == 3 and not app.last_batch.remaining_indices, errors
            assert (first.read_bytes(), first.stat().st_mtime_ns) == original
        else:
            app.output.set(str(folder / "delivery-output"))
            app.pdf.set(False)
            import core as core_module
            original_values = core_module.values_for
            previous = app.last_batch
            existing = set(Path(app.output.get()).glob("批次_*"))
            for action in (app.check_data, app.start_generation):
                entered, release = threading.Event(), threading.Event()
                def blocked_values(*args):
                    entered.set()
                    if not release.wait(10):
                        raise TimeoutError("检查停止验证超时")
                    return original_values(*args)
                with patch.object(core_module, "values_for", side_effect=blocked_values):
                    action()
                    try:
                        deadline = time.monotonic() + 10
                        while not entered.is_set():
                            assert time.monotonic() < deadline
                            app.update()
                            time.sleep(.01)
                        assert not app.cancel_button.instate(["disabled"])
                        assert app.cancel_button.cget("text") == "停止检查"
                        app.cancel_button.invoke()
                    finally:
                        release.set()
                        wait_idle()
                assert not app.check_completed and not app.issues
                assert app.last_batch is previous and not errors, errors
                assert set(Path(app.output.get()).glob("批次_*")) == existing
            original_check = check_records
            def late_cancel(*args, **kwargs):
                issues = original_check(*args, **kwargs)
                app.cancel_event.set()
                return issues
            with patch.object(ui, "check_records", side_effect=late_cancel):
                app.start_generation()
                wait_idle()
            assert app.last_batch is previous and not app.check_completed and not errors, errors
            assert set(Path(app.output.get()).glob("批次_*")) == existing
            app.required["姓名"].set(True)
            original = app.dataset.records[0]["姓名"]
            app.dataset.records[0]["姓名"] = ""
            app.check_data()
            wait_idle()
            assert any(issue["级别"] == "错误" for issue in app.issues)
            app.start_generation()
            wait_idle()
            assert app.last_batch is None
            app.scope.set("选中记录")
            app.data_tree.selection_set("1")
            app.update()
            app.check_data()
            wait_idle()
            assert not any(issue["级别"] == "错误" for issue in app.issues)
            app.dataset.records[0]["姓名"] = original
            assert app.save_task_profile("发布验证")
            app.required["姓名"].set(False)
            app.load_task_profile("发布验证")
            wait_idle()
            assert app.required["姓名"].get() and app.scope.get() == "全部记录"
            stop = threading.Event()
            job = app.last_job = (app.template, app.dataset, app._mapping(), app.pattern.get(), app._required_fields())
            app._run(lambda: app._execute_batch(
                lambda checkpoint: generate_batch(*job[:3], Path(app.output.get()), job[3], cancel=stop,
                                       progress=lambda *a: stop.set(), required_fields=job[4], checkpoint=checkpoint), job),
                app._batch_done, "正在验证发布功能…", generation=True)
            wait_idle()
            assert app.last_batch.succeeded == 1 and app.last_batch.remaining_indices == [1, 2], errors
        batch = app.last_batch
        assert not getattr(batch, "recovery_warning", ""), app.status.get()
        loaded, job = load_batch_state(batch.folder)
        assert loaded.succeeded == batch.succeeded and "姓名" in job[4]
        target = folder / "批次分享.zip"
        with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
            app.export_batch_archive()
            wait_idle()
        expected = {item[key] for item in batch.items for key in ("Word 文件", "PDF 文件") if item[key]}
        expected.add(batch.manifest_path.name)
        with zipfile.ZipFile(target) as archive:
            assert set(archive.namelist()) == expected and archive.testzip() is None
            assert all(hashlib.sha256(archive.read(name)).digest() == hashlib.sha256((batch.folder / name).read_bytes()).digest()
                       for name in expected)
        assert not errors, errors
        packed = target.read_bytes()
        document = batch.folder / batch.items[0]["Word 文件"]
        original = document.read_bytes()
        try:
            document.write_bytes(original + b"external change")
            with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
                app.export_batch_archive()
                wait_idle()
            assert len(errors) == 1 and target.read_bytes() == packed
        finally:
            document.write_bytes(original)
        errors.clear()
        marker = batch.folder / ".任务状态未保存"
        marker.touch()
        try:
            with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
                app.export_batch_archive()
                wait_idle()
            assert len(errors) == 1 and target.read_bytes() == packed
        finally:
            marker.unlink()
        errors.clear()
        return {"required_fields_and_scoped_checks": not bool(history), "snapshot_and_zip_verified": True,
                "changed_results_blocked": True, "foreign_unsaved_state_blocked": True,
                "manual_and_preflight_checks_stoppable": not bool(history), "late_check_stop_prevents_generation": not bool(history),
                "fresh_process_resume": bool(history), "history_fixture": str(batch.folder)}


def verify_large_search(app):
    from unittest.mock import patch
    original = app.dataset
    record = {name: '无匹配内容' * 40 for name in original.headers}
    records = [dict(record) for _ in range(50000)]
    records[40001][original.headers[1]] = '唯一搜索词'
    app.dataset = Dataset(original.headers, records, list(range(2, 50002)), original.sheet)
    app.selected_indices = {250, 40001}
    app._fill_data(reset=True)
    app.update()
    heartbeats = []
    ticking = [True]
    def tick():
        heartbeats.append(time.monotonic())
        if ticking[0]:
            app.after(10, tick)
    app.after(10, tick)
    app.data_query.set('唯一搜索词')
    started = time.monotonic()
    app._apply_data_search()
    assert app.busy
    while app.busy:
        assert time.monotonic() - started < 20
        app.update()
        time.sleep(.005)
    elapsed = time.monotonic() - started
    assert app.data_indices == [40001]
    assert app.selected_indices == {250, 40001}
    assert app.scope.get() == '全部记录' and '50000' in app.generate_button.cget('text')
    assert heartbeats
    before = list(app.data_indices), set(app.selected_indices)
    entered = threading.Event()
    def wait_for_cancel(dataset, query, cancel, progress):
        from data_search import SearchCancelled
        entered.set()
        assert cancel.wait(10)
        raise SearchCancelled('搜索已停止，保留原来的显示和选择。')
    with patch.object(ui, 'search_records', side_effect=wait_for_cancel):
        app.data_query.set('取消测试')
        app._apply_data_search()
        assert entered.wait(10)
        app.cancel_button.invoke()
        deadline = time.monotonic() + 10
        while app.busy:
            assert time.monotonic() < deadline
            app.update()
            time.sleep(.005)
    assert (app.data_indices, app.selected_indices) == before
    assert app.data_query.get() == '唯一搜索词'
    ticking[0] = False
    app.dataset = original
    app.selected_indices.clear()
    app._fill_data(reset=True)
    app.update()
    return {'large_search_records': len(records), 'large_search_seconds': round(elapsed, 3),
            'search_ui_heartbeats': len(heartbeats), 'hidden_selection_and_scope_preserved': True,
            'search_cancellation_preserves_previous_results': True}


def run(application_module):
    global ui
    ui = application_module
    report_path = Path(sys.argv[sys.argv.index("--smoke-test") + 1]) if "--smoke-test" in sys.argv else None
    peer_root = Path(sys.argv[sys.argv.index("--smoke-peer") + 1]) if report_path and "--smoke-peer" in sys.argv else None
    app = ui.Application(settings_path=peer_root / "preferences.json" if peer_root else report_path.parent / "smoke-settings.json" if report_path else None)
    # UI smoke testing runs the same callbacks and background workers as a normal session.
    if "--smoke-test" in sys.argv:
        report_path = Path(sys.argv[sys.argv.index("--smoke-test") + 1])
        if "--smoke-resume" in sys.argv:
            def resume_smoke():
                try:
                    folder = Path(sys.argv[sys.argv.index("--smoke-resume") + 1])
                    summary = verify_delivery_workflow(app, report_path.parent, folder)
                    summary["status"] = "PASS"
                    report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    report_path.write_text(traceback.format_exc(), encoding="utf-8")
                app._close()
            app.after(250, resume_smoke)
            app.mainloop()
            return
        if peer_root:
            app.withdraw()
            role = report_path.parent.name
            assert role in ("a", "b")
            (peer_root / (role + ".started")).write_text("started")
            def load_when_both_started():
                if all((peer_root / (name + ".started")).exists() for name in ("a", "b")):
                    app.load_demo()
                else:
                    app.after(100, load_when_both_started)
            app.after(250, load_when_both_started)
        else:
            app.after(250, app.load_demo)
        deadline = __import__("time").monotonic() + 120
        stage = [0]
        summary = {}
        if "--smoke-workflow" in sys.argv:
            summary["worker_phases"] = []
            report_phase = app._report_phase
            def capture_phase(phase, current, total):
                if phase not in summary["worker_phases"]:
                    summary["worker_phases"].append(phase)
                report_phase(phase, current, total)
            app._report_phase = capture_phase
        def smoke():
            try:
                if __import__("time").monotonic() > deadline:
                    raise TimeoutError("界面测试超时")
                if app.busy and "已用" in app.action_hint.get():
                    summary["live_elapsed_display"] = True
                if stage[0] == 0 and not app.busy and app.dataset:
                    if "--smoke-search" in sys.argv:
                        summary.update(verify_large_search(app))
                    if "--smoke-files" in sys.argv:
                        summary.update(verify_file_workflow(app, report_path.parent))
                    if "--smoke-report" in sys.argv:
                        summary.update(verify_check_report(app, report_path.parent))
                    if "--smoke-delivery" in sys.argv:
                        summary.update(verify_delivery_workflow(app, report_path.parent))
                    if "--smoke-template" in sys.argv:
                        path = Path(sys.argv[sys.argv.index("--smoke-template") + 1])
                        app._template_loaded(Template(path))
                        summary["template_source"] = str(path)
                    summary["template_fingerprint"] = app.template.fingerprint
                    assert not app.generate_button.instate(["disabled"]), app._mapping()
                    app.notebook.select(app.preview_page)
                    assert "林晓" in app.preview_text.get("1.0", "end")
                    app.output.set(str(report_path.parent / "gui-output"))
                    app.pdf.set("--smoke-pdf" in sys.argv)
                    if "--smoke-workflow" in sys.argv:
                        app.nav_buttons[1].invoke()
                        app.update()
                        assert app.notebook.select() == str(app.data_page)
                        assert app.nav_buttons[1].cget("style") == "ActiveNavigation.TButton"
                        summary["navigation_buttons_and_active_state"] = True
                        summary["animations_enabled"] = app.nav_motion.enabled
                        assert app.generate_button.cget("text") == "生成全部 3 份"
                        geometry, minimum = app.geometry(), app.minsize()
                        app.minsize(1, 1)
                        app.geometry(f"{app.winfo_width()}x{app.px(550)}")
                        app.update()
                        for page, tree in ((app.data_page, app.data_tree), (app.result_page, app.result_tree)):
                            app.notebook.select(page)
                            app.update()
                            assert tree.winfo_height() >= app.px(34) * 3
                            assert tree.xview() == (0.0, 1.0)
                        summary["sample_tables_fit_width"] = True
                        assert not app.export_panel.winfo_ismapped()
                        settings = app.open_export_settings()
                        app.update()
                        assert settings.winfo_ismapped()
                        settings.destroy()
                        app.minsize(*minimum)
                        app.geometry(geometry)
                        summary["compact_layout_and_export_settings"] = True
                        app.notebook.select(app.data_page)
                        app.update()
                        tree = app.data_tree
                        tree.selection_set("0", "1")
                        app.update()
                        cell = verify_cell_menu(app, tree, "0", "c6")
                        content = app.dataset.records[0]["备注"]
                        assert cell.content_text.get("1.0", "end-1c") == content
                        assert app.selected_indices == {0, 1}
                        from unittest.mock import patch
                        with patch.object(app, "clipboard_clear"), patch.object(app, "clipboard_append") as copied:
                            cell.copy_button.invoke()
                            copied.assert_called_once_with(content)
                        cell.destroy()
                        app.clear_selection()
                        summary["native_cell_menu_and_full_copy"] = True
                        app.data_query.set("陈雨")
                        app._apply_data_search()
                        assert app.data_indices == [1]
                        assert app.scope.get() == "全部记录"
                        app.locate_row.set("4")
                        app._locate_row()
                        assert app.data_query.get() == ""
                        assert app.preview_index.get() == "3"
                        app.scope.set("选中记录")
                        assert app.generate_button.cget("text") == "生成选中 1 份"
                        app.clear_selection()
                        assert not app.selected_indices and not app.data_tree.selection()
                        assert app.generate_button.cget("text") == "生成选中 0 份"
                        assert app.generate_button.instate(["disabled"])
                        assert "选择记录" in app.action_hint.get()
                        app.generate_button.event_generate("<Enter>")
                        app.update()
                        assert app.reason_popup is not None
                        assert "选择记录" in app.reason_popup.reason_text.cget("text")
                        app.generate_button.event_generate("<Leave>")
                        summary["disabled_action_reason"] = True
                        app.scope.set("全部记录")
                        last_folder = app.last_folder
                        preview = app.show_filename_preview()
                        names = preview_filenames(app.template, app.dataset, app._mapping(), app.pattern.get())
                        summary["filename_preview_names"] = [item["Word 文件"] for item in names]
                        content = preview.preview_text.get("1.0", "end-1c")
                        assert all(name in content for name in summary["filename_preview_names"])
                        assert app.last_folder == last_folder
                        preview.destroy()
                        summary["generation_scope_and_count"] = True
                        summary["clear_all_selection"] = True
                        summary["search_and_source_row_location"] = True
                        template, dataset, mapping, pattern = app.template, app.dataset, app._mapping(), app.pattern.get()
                        output, pdf = Path(app.output.get()), app.pdf.get()
                        app.last_job = (template, dataset, mapping, pattern)
                        def stopped_progress(number, total, item):
                            app.events.put(("progress", number, total, item, number - 1))
                            if number == 1:
                                time.sleep(.25)  # Let the UI display an in-flight operation even on fast machines.
                                app.cancel_event.set()
                        app._run(lambda: generate_batch(template, dataset, mapping, output, pattern,
                                                        export_pdf=pdf, cancel=app.cancel_event, progress=stopped_progress, phase=app._report_phase),
                                 app._batch_done, "正在验证停止与继续…", generation=True)
                        stage[0] = 3
                    else:
                        app.start_generation()
                        stage[0] = 1
                elif stage[0] == 3 and not app.busy and app.last_batch:
                    assert app.last_batch.succeeded == 1, app.last_batch.items
                    assert app.last_batch.remaining_indices == [1, 2]
                    assert not app.resume_button.instate(["disabled"])
                    original_files = [path for path in app.last_folder.iterdir() if path.suffix in (".docx", ".pdf")]
                    summary["_original_state"] = {path: (path.read_bytes(), path.stat().st_mtime_ns) for path in original_files}
                    summary["continued_folder"] = str(app.last_folder)
                    app.start_continue()
                    stage[0] = 1
                elif stage[0] == 1 and not app.busy and app.last_folder:
                    if "--smoke-workflow" in sys.argv:
                        assert str(app.last_folder) == summary["continued_folder"]
                        assert app.last_batch.succeeded == 3, app.last_batch.items
                        assert all("成功" in app.result_tree.item(str(i), "tags") for i in range(3))
                        colors = {str(app.result_tree.tag_configure(status, "foreground")) for status in ("成功", "部分成功", "失败")}
                        assert len(colors) == 3 and "" not in colors
                        summary["result_status_display"] = True
                        assert not app.last_batch.remaining_indices
                        assert app.resume_button.instate(["disabled"])
                        for path, snapshot in summary.pop("_original_state").items():
                            assert (path.read_bytes(), path.stat().st_mtime_ns) == snapshot
                        app.result_filter.set("仅失败和部分成功")
                        app._filter_results()
                        assert not app.result_tree.get_children()
                        app.result_filter.set("全部结果")
                        app._filter_results()
                        app.result_tree.selection_set("1")
                        app._refresh_ready()
                        assert not app.word_result_button.instate(["disabled"])
                        assert not app.manifest_button.instate(["disabled"])
                        if app.pdf.get():
                            assert not app.pdf_result_button.instate(["disabled"])
                        summary["same_batch_continued_and_success_files_unchanged"] = True
                        summary["result_filter_and_file_actions"] = True
                        details = app.show_result_details()
                        content = details.detail_text.get("1.0", "end-1c")
                        assert str(app.last_folder / app.result_items[1]["Word 文件"]) in content
                        assert "数据行：3" in content
                        from unittest.mock import patch
                        with patch.object(app, "clipboard_clear"), patch.object(app, "clipboard_append") as copied:
                            details.copy_button.invoke()
                            copied.assert_called_once_with(content)
                        details.destroy()
                        summary["result_details_and_copy"] = True
                        assert summary["filename_preview_names"] == [item["Word 文件"] for item in app.result_items]
                        assert all(phase in summary["worker_phases"] for phase in ("填充 Word", "保存清单"))
                        if app.pdf.get():
                            assert "转换 PDF" in summary["worker_phases"]
                        assert summary["live_elapsed_display"]
                        assert "上次操作耗时" in app.action_hint.get()
                        summary["filename_preview_matches_generation"] = True
                        summary["phase_and_elapsed_feedback"] = True
                    files = list(app.last_folder.glob("*.docx"))
                    assert len(files) == 3
                    pdf_count = len(list(app.last_folder.glob("*.pdf")))
                    if "--smoke-pdf" in sys.argv:
                        assert pdf_count == 3, app.status.get()
                    summary.update(status="PASS", documents=len(files), pdf_documents=pdf_count, folder=str(app.last_folder))
                    if "--smoke-preview" in sys.argv:
                        app.open_page_preview()
                        stage[0] = 2
                    else:
                        report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                        app._close()
                        return
                elif stage[0] == 2 and not app.busy and app.preview_windows and app.preview_windows[-1].photo:
                    window = app.preview_windows[-1]
                    folder = window.folder
                    summary["preview_pages"] = len(window.document)
                    if peer_root:
                        from pdf_worker import write_json
                        os.utime(folder, (time.time()-60, time.time()-60))
                        os.utime(app.preview_root, (time.time()-60, time.time()-60))
                        write_json(peer_root / (role + ".ready.json"), {"folder": str(folder), "instance": str(app.preview_root)})
                        stage[0] = 4
                    else:
                        window.close()
                        summary["preview_cleaned"] = not folder.exists()
                        assert summary["preview_cleaned"]
                        report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                        app._close()
                        return
                elif stage[0] == 4 and all((peer_root / (name + ".ready.json")).exists() for name in ("a", "b")):
                    peer = "b" if role == "a" else "a"
                    peer_info = json.loads((peer_root / (peer + ".ready.json")).read_text(encoding="utf-8"))
                    assert peer_info["instance"] != str(app.preview_root)
                    app._cleanup_previews(schedule=False)
                    assert (Path(peer_info["folder"]) / "文档预览.docx").is_file()
                    assert (Path(peer_info["folder"]) / "文档预览.pdf").is_file()
                    summary["other_process_preview_kept"] = True
                    (peer_root / (role + ".checked")).write_text("checked")
                    stage[0] = 5
                elif stage[0] == 5 and all((peer_root / (name + ".checked")).exists() for name in ("a", "b")):
                    window = app.preview_windows[-1]
                    folder = window.folder
                    window.close()
                    summary["preview_cleaned"] = not folder.exists()
                    assert summary["preview_cleaned"]
                    report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                    app._close()
                    return
            except Exception:
                report_path.write_text(traceback.format_exc(), encoding="utf-8")
                app._close()
                return
            app.after(100, smoke)
        app.after(400, smoke)
    app.mainloop()


if __name__ == "__main__":
    import app
    run(app)
