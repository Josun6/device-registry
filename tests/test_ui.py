import importlib.util
import json
import shutil
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

SOURCE = Path(__file__).resolve().parents[1] / "device_registry.py"


@pytest.fixture
def module(tmp_path):
    copy = tmp_path / "device_registry.py"
    shutil.copy2(SOURCE, copy)
    shutil.copytree(SOURCE.parent / "assets", tmp_path / "assets")
    old_argv0 = sys.argv[0]
    sys.argv[0] = str(copy)
    try:
        spec = importlib.util.spec_from_file_location("registry_test_module", copy)
        imported = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(imported)
        yield imported
    finally:
        sys.argv[0] = old_argv0


def sample(mod, name="Printer", code="P-001", company="Example Co"):
    data = {key: "" for key, _ in mod.FIELDS}
    data.update({
        "name": name, "code": code, "company": company,
        "status": "在用", "purchase_date": "2024-03-01",
        "qty": 1.0, "value": 1000.0, "warranty_months": 0.0,
    })
    return data


def test_database_and_obsolete_preferences_are_retained(module, tmp_path):
    db = module.DB(tmp_path / "test.db")
    try:
        db.set_setting("ui_visible_fields_v2", json.dumps(["name", "code"]))
        db.set_setting("ui_form_fields", json.dumps(["name", "code"]))
        db.set_setting("ui_list_fields", json.dumps(["name", "code"]))
        new_id = db.save(sample(module))
        assert db.get(new_id)["company"] == "Example Co"
        assert db.distinct("company") == ["Example Co"]
    finally:
        db.close()


def test_fixed_columns_and_entry_fields_match_requested_layout(module):
    expected = (
        "name", "code", "model", "site", "owner", "location", "status",
        "qty", "unit", "value", "purchase_date", "warranty_months", "net", "remark",
    )
    assert module.LIST_COLUMNS == expected
    assert module.FORM_FIELDS == tuple(key for key in expected if key != "net") + ("image",)
    app = module.App()
    try:
        app.update()
        assert isinstance(app.split, module.tk.PanedWindow)
        assert len(app.split.panes()) == 2
        assert list(app.tree["columns"]) == list(expected)
        assert list(app.dv) == list(expected)
        assert not hasattr(app, "open_ui_settings")
        assert not hasattr(app, "visible_fields")
        assert not hasattr(app, "form_fields")
        assert not hasattr(app, "list_fields")
        assert hasattr(app, "search_canvas")
        app.db.save(sample(module))
        app.refresh()
        assert app.default_company_value() == "Example Co"
        with patch.object(module.tk.Toplevel, "wait_window", lambda self, *_: None):
            dialog = module.EditDialog(app)
        assert list(dialog.field_rows) == list(module.FORM_FIELDS)
        assert "company" not in dialog.vars
        assert "status_date" not in dialog.vars
        assert "net" not in dialog.vars
        assert dialog.vars["purchase_date"].get()
        dialog.vars["name"].set("New Device")
        dialog.vars["code"].set("P-002")
        dialog.ok()
        assert dialog.result["company"] == "Example Co"
        assert dialog.result["name"] == "New Device"
        assert not any("显示字段" in item for item in app.lock_menu.entrycget(0, "label"))
    finally:
        app.on_close()


def test_public_title_and_about_are_generic():
    contents = SOURCE.read_text(encoding="utf-8")
    assert 'self.title(tr(f"Device Registry' in contents
    assert "Copyright ©" not in contents



def test_password_old_format_and_new_format(module, tmp_path):
    import hashlib
    db = module.DB(tmp_path / "auth.db")
    try:
        salt = bytes.fromhex("01" * 16)
        old = hashlib.pbkdf2_hmac("sha256", b"abc123", salt, 200_000)
        db.set_setting("pw", f"{salt.hex()}:{old.hex()}")
        assert db.check_password("abc123")
        assert not db.check_password("wrong")
        db.set_password("new-password")
        assert db._get("pw").startswith("pbkdf2_sha256$")
        assert db.check_password("new-password")
        assert not db.check_password("abc123")
    finally:
        db.close()


def test_invalid_numbers_and_unsafe_paths(module, tmp_path):
    db = module.DB(tmp_path / "numbers.db")
    try:
        for invalid in (float("nan"), float("inf"), -1):
            data = sample(module, code=f"BAD-{str(invalid)}")
            data["value"] = invalid
            with pytest.raises(ValueError):
                db.save(data)
        safe = module.stored_file_path(tmp_path, "abc123.png")
        assert safe == str(tmp_path / "abc123.png")
        for unsafe in ("../secret", "..\\\\secret", "/etc/passwd", ".", ""):
            with pytest.raises(ValueError):
                module.stored_file_path(tmp_path, unsafe)
    finally:
        db.close()


def test_restore_rejects_unsafe_zip_and_recovers_valid_backup(module, tmp_path):
    import zipfile
    app = module.App()
    try:
        app.unlocked = True
        app.update_mode()
        app.db.save(sample(module))
        valid_backup = tmp_path / "valid.zip"
        with zipfile.ZipFile(valid_backup, "w") as z:
            z.write(app.db.path, "devices.db")
        app.db.save(sample(module, name="Monitor", code="P-002"))
        assert len(app.db.query()) == 2

        bad_backup = tmp_path / "malicious.zip"
        with zipfile.ZipFile(bad_backup, "w") as z:
            z.write(app.db.path, "devices.db")
            z.writestr("images/../../escape.txt", "bad")

        errors = []
        with patch.object(module.filedialog, "askopenfilename", return_value=str(bad_backup)), \
             patch.object(module.messagebox, "askyesno", return_value=True), \
             patch.object(module.messagebox, "showerror", side_effect=lambda *a: errors.append(a)):
            app.restore()
        assert errors
        assert len(app.db.query()) == 2

        with patch.object(module.filedialog, "askopenfilename", return_value=str(valid_backup)), \
             patch.object(module.messagebox, "askyesno", return_value=True), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.restore()
        assert len(app.db.query()) == 1
        assert (tmp_path / "devices.db.bak").exists()
        assert app.unlocked and not app.protection_enabled
    finally:
        app.on_close()


def test_bilingual_runtime_preserves_filters_and_database(module):
    app = module.App()
    try:
        new_id = app.db.save(sample(module, name="Desktop", code="D-100"))
        app.refresh()
        app.tree.selection_set(str(new_id))
        app.kw.set("Desktop")
        app.fv["status"].set("在用")
        assert len(app.rows()) == 1
        app.change_language("en")
        assert app.language == "en"
        assert "Asset Manager" in app.title()
        assert app.tree.heading("name")["text"] == "Device Name"
        assert app.tree.heading("code")["text"] == "Asset Code"
        assert app.tree.heading("status")["text"] == "Status"
        assert app.fv["status"].get() == "In Use"
        assert app.tree.set(str(new_id), "status") == "In Use"
        assert len(app.rows()) == 1
        assert app.db.get(new_id)["status"] == "在用"
        assert app.db._get(module.UI_LANGUAGE_KEY) == "en"
        assert app.language_choice.get() == "en"
        detail_labels = [w.cget("text") for w in app.detail.winfo_children()
                         if isinstance(w, module.ttk.Label)]
        assert "Device Name: " in detail_labels
        assert "Asset Code: " in detail_labels
        app.change_language("zh")
        assert "设备资产管理" in app.title()
        assert app.fv["status"].get() == "在用"
        assert app.tree.heading("name")["text"] == "设备名称"
        assert app.db.get(new_id)["status"] == "在用"
    finally:
        app.on_close()


def test_english_dialog_uses_english_labels_but_saves_canonical_status(module):
    app = module.App()
    try:
        app.change_language("en")
        with patch.object(module.tk.Toplevel, "wait_window", lambda self, *_: None):
            dialog = module.EditDialog(app)
        assert dialog.title() == "Add Device"
        assert dialog.vars["status"].get() == "In Use"
        assert dialog.name_input.winfo_exists()
        dialog.vars["name"].set("Laptop")
        dialog.vars["code"].set("LT-01")
        dialog.ok()
        assert dialog.result["status"] == "在用"
        assert dialog.result["name"] == "Laptop"
    finally:
        app.on_close()


def test_icon_loaded_in_main_window_and_ico_has_small_sizes(module):
    from PIL import Image
    ico = SOURCE.parent / "assets" / "icon.ico"
    png = SOURCE.parent / "assets" / "icon.png"
    assert ico.is_file() and png.is_file()
    with Image.open(ico) as img:
        sizes = img.info.get("sizes", set())
        assert (16, 16) in sizes
        assert (256, 256) in sizes
    app = module.App()
    try:
        assert bool(getattr(app, "_registry_icon", None))
    finally:
        app.on_close()


def test_language_pref_persists_on_restart(module):
    app = module.App()
    app.change_language("en")
    app.on_close()
    reopened = module.App()
    try:
        assert reopened.language == "en"
        assert reopened.tree.heading("name")["text"] == "Device Name"
        assert "Asset Manager" in reopened.title()
    finally:
        reopened.on_close()


def test_english_csv_round_trip_and_external_english_status(module, tmp_path):
    import csv
    app = module.App()
    try:
        app.unlocked = True
        app.update_mode()
        app.db.save(sample(module, name="Monitor", code="M-1"))
        app.change_language("en")
        app.update_mode()
        output = tmp_path / "device-list.csv"
        with patch.object(module.filedialog, "asksaveasfilename", return_value=str(output)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.export_file()
        with open(output, encoding="utf-8-sig", newline="") as stream:
            data = list(csv.reader(stream))
        assert "Device Name" in data[0]
        assert "Asset Code" in data[0]
        assert "Status" in data[0]
        assert data[1][data[0].index("Status")] == "在用"

        app.db.delete([1])
        assert len(app.db.query()) == 0
        with patch.object(module.filedialog, "askopenfilename", return_value=str(output)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.import_file()
        assert len(app.db.query()) == 1
        assert app.db.query()[0]["status"] == "在用"

        external = tmp_path / "manual-english.csv"
        with external.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Device Name", "Asset Code", "Status", "Purchase Date", "Quantity", "Unit Price (CNY)"])
            writer.writerow(["Keyboard", "K-10", "In Use", "2026-01-01", "1", "49.5"])
        with patch.object(module.filedialog, "askopenfilename", return_value=str(external)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.import_file()
        assert app.db.find_code("K-10")
        assert app.db.get(app.db.find_code("K-10")["id"])["status"] == "在用"
    finally:
        app.on_close()


def test_hidden_old_settings_do_not_hide_columns_or_discard_data(module):
    db = module.DB(module.DB_PATH)
    legacy = sample(module, name="Tablet", code="T-100", company="Retained Company")
    legacy["model"] = "Original Model"
    device_id = db.save(legacy)
    db.set_setting("ui_visible_fields_v2", json.dumps(["name", "code"]))
    db.set_setting("ui_form_fields", json.dumps(["name", "code"]))
    db.set_setting("ui_list_fields", json.dumps(["name", "code"]))
    db.close()

    app = module.App()
    try:
        assert list(app.tree["columns"]) == list(module.LIST_COLUMNS)
        assert "model" in app.dv
        with patch.object(module.tk.Toplevel, "wait_window", lambda self, *_: None):
            edit = module.EditDialog(app, row=app.db.get(device_id))
        assert "company" not in edit.vars
        assert edit.original["company"] == "Retained Company"
        edit.vars["name"].set("Tablet 2")
        edit.ok()
        app.db.save(edit.result, device_id)
        reloaded = app.db.get(device_id)
        assert reloaded["company"] == "Retained Company"
        assert reloaded["model"] == "Original Model"
        assert reloaded["name"] == "Tablet 2"
        assert reloaded["status_date"] == "2026-01-01" or reloaded["status_date"]
    finally:
        app.on_close()


def test_user_options_only_change_company_and_alerts(module):
    app = module.App()
    try:
        before = tuple(app.tree["columns"])
        win = app.open_options()
        assert bool(getattr(win, "_registry_icon", None))
        assert not hasattr(win, "field_vars")
        assert not hasattr(win, "visible_fields")
        texts = []
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    texts.append(child.cget("text"))
                except module.tk.TclError:
                    pass
                walk(child)
        walk(win)
        assert module.tr("默认所属公司：") in texts
        assert module.tr("启用保修到期提醒") in texts
        assert module.tr("启用维修超期提醒") in texts
        assert tuple(app.tree["columns"]) == before
        win.destroy()
    finally:
        app.on_close()


def test_responsive_column_widths_keep_full_headings(module):
    app = module.App()
    try:
        for index in range(3):
            app.db.save(sample(module, name=f"Special Large Asset {index}",
                                  code=f"K-{index:03d}"))
        app.refresh()
        app.update()
        width = app.split.winfo_width()
        app.split.sash_place(0, int(width * 0.47), 0)
        app.update()
        app._autofit_columns()
        assert tuple(app.tree["columns"]) == module.LIST_COLUMNS
        for field in module.LIST_COLUMNS:
            minimum = app._heading_font.measure(module.tr(module.HEAD[field])) + module.S(28)
            assert app.tree.column(field, "width") >= minimum
            assert app.tree.column(field, "minwidth") >= minimum
            assert not app.tree.column(field, "stretch")
        assert app.tree.xview()[1] < 1, "Horizontal scrolling needed for fixed columns"
        app.change_language("en")
        app.update()
        app._autofit_columns()
        for field in module.LIST_COLUMNS:
            minimum = app._heading_font.measure(module.tr(module.HEAD[field])) + module.S(28)
            assert app.tree.column(field, "width") >= minimum
    finally:
        app.on_close()


def test_wide_list_expands_columns_when_space_is_available(module):
    app = module.App()
    try:
        app.geometry("4000x840+0+0")
        app.update()
        width = app.split.winfo_width()
        app.split.sash_place(0, int(width * 0.85), 0)
        app.update()
        app._autofit_columns()
        target = app.tree.winfo_width() - module.S(4)
        actual = sum(app.tree.column(key, "width") for key in module.LIST_COLUMNS)
        assert abs(actual - target) <= 6, (target, actual)
    finally:
        app.on_close()


def test_new_install_editing_is_enabled_without_password_prompt(module):
    app = module.App()
    try:
        app.update()
        assert not app.protection_enabled
        assert app.unlocked
        assert app.need_unlock()
        assert app.db._get(module.UI_EDIT_PROTECT_KEY) is None
        assert app.lock_btn.cget("text") == module.tr("编辑保护：已关闭")
        assert "disabled" not in app.edit_btns[0].state()
        assert app.lock_menu.entrycget(module.tr("修改密码"), "state") == "disabled"
        assert "无需密码" in app.mode_lbl.cget("text")
    finally:
        app.on_close()


def test_opt_in_password_can_be_locked_and_turned_off_with_verification(module):
    app = module.App()
    try:
        dialog = app.configure_edit_protection()
        assert dialog.protection_choice.get() is False
        assert bool(getattr(dialog, "_registry_icon", None))
        dialog.protection_choice.set(True)
        with patch.object(module.simpledialog, "askstring", side_effect=["my-password", "my-password"]), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            dialog.save_protection()
        assert app.db.check_password("my-password")
        assert app.db._get(module.UI_EDIT_PROTECT_KEY) == "1"
        assert app.protection_enabled and app.unlocked
        assert app.lock_menu.entrycget(module.tr("修改密码"), "state") == "normal"

        app.toggle_lock()
        assert not app.unlocked
        assert "disabled" in app.edit_btns[0].state()
        with patch.object(module.messagebox, "showinfo", return_value=None):
            assert not app.need_unlock()
        with patch.object(module.simpledialog, "askstring", return_value="wrong"), \
             patch.object(module.messagebox, "showerror", return_value=None):
            app.toggle_lock()
        assert not app.unlocked
        with patch.object(module.simpledialog, "askstring", return_value="my-password"):
            app.toggle_lock()
        assert app.unlocked

        # Existing password is required even when editing has already been unlocked.
        off = app.configure_edit_protection()
        off.protection_choice.set(False)
        with patch.object(module.simpledialog, "askstring", return_value="wrong"), \
             patch.object(module.messagebox, "showerror", return_value=None):
            off.save_protection()
        assert app.protection_enabled and app.db.has_password()
        with patch.object(module.simpledialog, "askstring", return_value="my-password"), \
             patch.object(module.messagebox, "askyesno", return_value=False):
            off.save_protection()
        assert app.protection_enabled
        with patch.object(module.simpledialog, "askstring", return_value="my-password"), \
             patch.object(module.messagebox, "askyesno", return_value=True), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            off.save_protection()
        assert not app.protection_enabled
        assert app.unlocked
        assert not app.db.has_password()
        assert app.db._get(module.UI_EDIT_PROTECT_KEY) == "0"
        assert "disabled" not in app.edit_btns[0].state()
    finally:
        app.on_close()
    reopened = module.App()
    try:
        assert reopened.unlocked and not reopened.protection_enabled
        assert reopened.need_unlock()
    finally:
        reopened.on_close()


def test_old_password_database_is_still_locked_by_default(module):
    original = module.DB(module.DB_PATH)
    original.set_password("legacy-pass")
    original.close()
    app = module.App()
    try:
        assert app.protection_enabled
        assert not app.unlocked
        assert "disabled" in app.edit_btns[0].state()
        assert app.db._get(module.UI_EDIT_PROTECT_KEY) is None
        with patch.object(module.simpledialog, "askstring", return_value="legacy-pass"):
            app.toggle_lock()
        assert app.unlocked
        assert app.db.check_password("legacy-pass")
    finally:
        app.on_close()


def test_visible_splitter_drag_and_persistence(module):
    app = module.App()
    try:
        app.geometry("1350x790+0+0")
        app.update()
        assert isinstance(app.split, module.tk.PanedWindow)
        assert app.split.cget("showhandle") in (True, 1, "1")
        assert int(app.split.cget("sashwidth")) >= module.S(9)
        assert app.split.cget("cursor") == "sb_h_double_arrow"
        assert len(app.split.panes()) == 2
        width = app.split.winfo_width()
        assert width >= 600
        app.split.sash_place(0, int(width * 0.68), 0)
        app.update()
        before = app.split.sash_coord(0)
        # Tk's native bindings: click on the visible separator, drag left and release.
        x, y = before
        app.split.event_generate("<ButtonPress-1>", x=x + 3, y=y + 18)
        app.split.event_generate("<B1-Motion>", x=x - module.S(105), y=y + 18)
        app.split.event_generate("<ButtonRelease-1>", x=x - module.S(105), y=y + 18)
        app.update()
        after = app.split.sash_coord(0)
        assert after[0] < before[0] - module.S(35), (before, after)
        saved = float(app.db._get(module.UI_SPLIT_KEY))
        assert abs(saved - after[0] / width) < 0.02
    finally:
        app.on_close()
    app2 = module.App()
    try:
        app2.geometry("1350x790+0+0")
        app2.update()
        assert abs(app2.split.sash_coord(0)[0] / app2.split.winfo_width() - saved) < .05
    finally:
        app2.on_close()


def test_dialogs_and_build_scripts_have_icons(module):
    app = module.App()
    try:
        assert getattr(app, "_registry_icon", None) is not None
        options = app.open_options()
        assert bool(getattr(options, "_registry_icon", None))
        options.destroy()
        with patch.object(module.tk.Toplevel, "wait_window", lambda _self, *_args: None):
            dialog = module.EditDialog(app)
        assert bool(getattr(dialog, "_registry_icon", None))
        dialog.destroy()
    finally:
        app.on_close()
    repo = SOURCE.parent
    for name in ("build_exe.bat", ".github/workflows/windows-exe.yml"):
        text = (repo / name).read_text(encoding="utf-8")
        assert "icon.ico" in text
        assert "icon.png" in text




def test_fixed_layout_in_both_languages_and_no_field_controls(module):
    app = module.App()
    try:
        app.update()
        assert tuple(app.tree["columns"]) == module.LIST_COLUMNS
        assert [app.tree.heading(key)["text"] for key in module.LIST_COLUMNS] == [
            module.HEAD[key] for key in module.LIST_COLUMNS
        ]
        app.change_language("en")
        assert tuple(app.tree["columns"]) == module.LIST_COLUMNS
        assert [app.tree.heading(key)["text"] for key in module.LIST_COLUMNS] == [
            module.tr(module.HEAD[key]) for key in module.LIST_COLUMNS
        ]
        assert not hasattr(app, "open_ui_settings")
        assert "net" not in module.FORM_FIELDS
        assert "image" in module.FORM_FIELDS
        assert "company" not in module.LIST_COLUMNS
    finally:
        app.on_close()


def test_editing_preserves_legacy_company_and_status_date(module):
    app = module.App()
    try:
        row = sample(module, name="Switch", code="S-200", company="Archive Division")
        row["status_date"] = "2020-10-03"
        row["model"] = "24-port"
        row["site"] = "HQ"
        device_id = app.db.save(row)
        assert app.db.get(device_id)["status_date"] == "2020-10-03"
        with patch.object(module.tk.Toplevel, "wait_window", lambda *args: None):
            dialog = module.EditDialog(app, row=app.db.get(device_id))
        dialog.vars["model"].set("48-port")
        dialog.ok()
        app.db.save(dialog.result, device_id)
        current = app.db.get(device_id)
        assert current["company"] == "Archive Division"
        assert current["status_date"] == "2020-10-03"
        assert current["model"] == "48-port"
        with patch.object(module.tk.Toplevel, "wait_window", lambda *args: None):
            another = module.EditDialog(app, row=current)
        another.vars["status"].set(module.tr("维修"))
        another.ok()
        app.db.save(another.result, device_id)
        changed = app.db.get(device_id)
        assert changed["company"] == "Archive Division"
        assert changed["status"] == "维修"
        assert changed["status_date"] == module.date.today().isoformat()
    finally:
        app.on_close()


def test_fixed_form_rejects_invalid_values_without_writing(module):
    app = module.App()
    try:
        with patch.object(module.tk.Toplevel, "wait_window", lambda *args: None):
            dialog = module.EditDialog(app)
        dialog.vars["name"].set("Test")
        dialog.vars["code"].set("TX-1")
        dialog.vars["value"].set("inf")
        with patch.object(module.messagebox, "showwarning", return_value=None):
            dialog.ok()
        assert dialog.result is None
        dialog.vars["value"].set("99.5")
        dialog.ok()
        assert dialog.result["value"] == 99.5
        assert dialog.result["company"] == ""
    finally:
        app.on_close()


def test_default_company_setting_applies_to_new_devices(module):
    app = module.App()
    try:
        app.db.set_setting(module.UI_COMPANY_KEY, "Team A")
        app._load_preferences()
        with patch.object(module.tk.Toplevel, "wait_window", lambda *args: None):
            dialog = module.EditDialog(app)
        dialog.vars["name"].set("Barcode Reader")
        dialog.vars["code"].set("BR-1")
        dialog.ok()
        assert dialog.result["company"] == "Team A"
        app.db.save(dialog.result)
        assert app.db.query(company="Team A")
    finally:
        app.on_close()


def test_generated_archive_is_free_of_business_data():
    root = SOURCE.parent
    assert not list(root.glob("devices.db"))
    assert not list(root.glob("*.db.bak"))
    assert not list(root.glob("*.xlsx"))
    assert not list(root.glob("*.csv"))
    assert (root / "LICENSE").exists()
    assert (root / "README_EN.md").exists()
    assert (root / "SECURITY.md").exists()


def test_restore_rolls_back_overwritten_attachment_on_late_failure(module, tmp_path):
    """A failed database swap must not leave new files behind old records."""
    import zipfile

    app = module.App()
    try:
        device_id = app.db.save(sample(module))
        stored = "invoice-example.pdf"
        source = Path(module.ATT_DIR) / stored
        source.write_bytes(b"OLD INVOICE CONTENT")
        app.db.conn.execute(
            "INSERT INTO attachments(device_id, orig_name, stored, added) VALUES(?,?,?,?)",
            (device_id, "invoice.pdf", stored, "2026-01-01"),
        )
        app.db.conn.commit()
        snapshot = tmp_path / "old.db"
        with module.sqlite3.connect(snapshot) as target:
            app.db.conn.backup(target)
        restore_zip = tmp_path / "replacement.zip"
        with zipfile.ZipFile(restore_zip, "w") as archive:
            archive.write(snapshot, "devices.db")
            archive.writestr("attachments/" + stored, b"NEW INVOICE CONTENT")

        original_replace = module.os.replace
        failures = []

        def fail_database_swap(source_file, destination):
            if str(destination) == str(module.DB_PATH) and ".device-registry-restore-" in str(source_file):
                raise OSError("Simulated database swap failure")
            return original_replace(source_file, destination)

        with patch.object(module.filedialog, "askopenfilename", return_value=str(restore_zip)), \
             patch.object(module.messagebox, "askyesno", return_value=True), \
             patch.object(module.messagebox, "showerror", side_effect=lambda *a: failures.append(a)), \
             patch.object(module.os, "replace", side_effect=fail_database_swap):
            app.restore()
        assert failures, "Restore should report the injected failure"
        assert source.read_bytes() == b"OLD INVOICE CONTENT"
        assert len(app.db.query()) == 1
        assert app.db.conn.execute("PRAGMA quick_check").fetchone()[0] == "ok"
    finally:
        app.on_close()


def test_backup_includes_safe_files_but_skips_symlinks(module, tmp_path):
    import zipfile

    app = module.App()
    try:
        app.db.save(sample(module))
        safe = Path(module.IMG_DIR) / "example.png"
        safe.write_bytes(b"png placeholder")
        outside = tmp_path / "outside-secret.txt"
        outside.write_text("not for inclusion", encoding="utf-8")
        link = Path(module.IMG_DIR) / "linked-secret.txt"
        try:
            link.symlink_to(outside)
        except OSError:
            pytest.skip("OS does not support symlinks")
        output = tmp_path / "backup.zip"
        with patch.object(module.filedialog, "asksaveasfilename", return_value=str(output)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.backup()
        with zipfile.ZipFile(output) as archive:
            assert "devices.db" in archive.namelist()
            assert "images/example.png" in archive.namelist()
            assert "images/linked-secret.txt" not in archive.namelist()
    finally:
        app.on_close()


def test_failed_save_shows_message_instead_of_crashing_ui(module):
    app = module.App()
    try:
        saved = sample(module)
        saved["qty"] = float("nan")
        messages = []
        with patch.object(module.messagebox, "showerror", side_effect=lambda *a, **kw: messages.append(a)):
            assert app.try_save(saved) is None
        assert messages
        assert not app.db.query()
    finally:
        app.on_close()


def test_excel_round_trip_when_openpyxl_is_installed(module, tmp_path):
    pytest.importorskip("openpyxl")
    app = module.App()
    try:
        record = sample(module, name="Excel Asset", code="XL-01", company="Original Company")
        record["model"] = "Keyboard"
        device_id = app.db.save(record)
        app.refresh()
        exported = tmp_path / "assets.xlsx"
        with patch.object(module.filedialog, "asksaveasfilename", return_value=str(exported)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.export_file()
        assert exported.exists()
        app.db.delete([device_id])
        assert not app.db.query()
        with patch.object(module.filedialog, "askopenfilename", return_value=str(exported)), \
             patch.object(module.messagebox, "showinfo", return_value=None):
            app.import_file()
        rows = app.db.query()
        assert len(rows) == 1
        assert rows[0]["name"] == "Excel Asset"
        assert rows[0]["company"] == "Original Company"
        assert rows[0]["model"] == "Keyboard"
    finally:
        app.on_close()
