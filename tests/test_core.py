"""Data layer tests that also run without a desktop session (Windows CI)."""

import hashlib
import importlib.util
import shutil
import sqlite3
import sys
from datetime import date
from pathlib import Path

import pytest

SOURCE = Path(__file__).resolve().parents[1] / "device_registry.py"


@pytest.fixture
def registry(tmp_path):
    program = tmp_path / "device_registry.py"
    shutil.copy2(SOURCE, program)
    old_argv = sys.argv[0]
    sys.argv[0] = str(program)
    try:
        spec = importlib.util.spec_from_file_location("registry_core_tests", program)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        yield module
    finally:
        sys.argv[0] = old_argv


def device(registry, code="ASSET-01", status="在用"):
    row = {field: "" for field, _ in registry.FIELDS}
    row.update(
        name="Example Device", code=code, status=status, qty=1,
        value=1499.0, warranty_months=24, purchase_date="2025-01-10",
    )
    return row


def test_legacy_schema_is_extended_without_dropping_records(registry, tmp_path):
    path = tmp_path / "legacy.db"
    with sqlite3.connect(path) as conn:
        conn.execute("""CREATE TABLE devices (
            id INTEGER PRIMARY KEY, name TEXT NOT NULL, code TEXT UNIQUE NOT NULL,
            model TEXT, owner TEXT, location TEXT, status TEXT,
            qty REAL, unit TEXT, value REAL, remark TEXT, image TEXT
        )""")
        conn.execute("""INSERT INTO devices
            (name, code, model, status, qty, value)
            VALUES ('Old Laptop', 'L-1', 'T480', '在用', 1, 999)""")
    db = registry.DB(path)
    try:
        assert len(db.query()) == 1
        record = db.query()[0]
        assert record["model"] == "T480"
        assert record["company"] is None
        assert record["site"] is None
        assert record["purchase_date"] is None
        assert record["status_date"]
        assert db.conn.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        assert len(db.conn.execute("PRAGMA table_info(devices)").fetchall()) >= 16
    finally:
        db.close()


def test_code_uniqueness_is_case_insensitive(registry, tmp_path):
    db = registry.DB(tmp_path / "codes.db")
    try:
        first = db.save(device(registry, "l-900"))
        assert db.find_code("L-900")["id"] == first
        with pytest.raises(sqlite3.IntegrityError):
            db.save(device(registry, "L-900"))
        assert len(db.query()) == 1
    finally:
        db.close()


def test_status_date_resets_only_when_status_changes(registry, tmp_path):
    db = registry.DB(tmp_path / "status.db")
    try:
        record = device(registry)
        record["status_date"] = "2020-01-05"
        row_id = db.save(record)
        record["value"] = 1999
        record.pop("status_date", None)
        db.save(record, row_id)
        assert db.get(row_id)["status_date"] == "2020-01-05"
        record["status"] = "维修"
        db.save(record, row_id)
        assert db.get(row_id)["status_date"] == date.today().isoformat()
        history = db.logs_for(row_id)
        assert any(item["action"] == "编辑" for item in history)
    finally:
        db.close()


@pytest.mark.parametrize("invalid", ["nan", "inf", "-10", "not-a-number"])
def test_numeric_validation_rejects_invalid_values(registry, tmp_path, invalid):
    db = registry.DB(tmp_path / "numbers.db")
    try:
        row = device(registry)
        row["value"] = invalid
        with pytest.raises(ValueError):
            db.save(row)
        assert not db.query()
    finally:
        db.close()


def test_backup_password_compatibility(registry, tmp_path):
    db = registry.DB(tmp_path / "password.db")
    try:
        salt = b"0123456789ABCDEF"
        digest = hashlib.pbkdf2_hmac("sha256", b"legacy", salt, 200_000)
        db.set_setting("pw", f"{salt.hex()}:{digest.hex()}")
        assert db.edit_protection_enabled()
        assert db.check_password("legacy")
        db.set_password("new-password")
        assert db.check_password("new-password")
        assert not db.check_password("legacy")
        db.disable_edit_protection()
        assert not db.edit_protection_enabled()
        assert not db.has_password()
    finally:
        db.close()


def test_fixed_columns_have_translations_and_export_keeps_full_schema(registry):
    assert len(registry.LIST_COLUMNS) == 14
    assert registry.LIST_COLUMNS[0] == "name"
    assert registry.LIST_COLUMNS[-1] == "remark"
    assert registry.LIST_COLUMNS[-2] == "net"
    assert "company" not in registry.LIST_COLUMNS
    assert "net" not in registry.FORM_FIELDS
    assert "company" not in registry.FORM_FIELDS
    assert len(registry.DISPLAY) > len(registry.LIST_COLUMNS)
    registry.set_ui_language("en")
    assert registry.tr("设备名称") == "Device Name"
    assert registry.tr("默认公司与提醒") == "Default Company & Alerts"


def test_date_helpers_handle_leap_year(registry):
    leap = date(2024, 1, 31)
    assert registry.add_months(leap, 1) == date(2024, 2, 29)
    assert registry.parse_date("20240229") == date(2024, 2, 29)
    assert registry.parse_date("2024/02/29") == date(2024, 2, 29)
    assert registry.parse_date("2023-02-29") is None
    assert registry.add_months(date(2024, 2, 29), 12) == date(2025, 2, 28)


def test_file_path_validation_prevents_directory_escape(registry, tmp_path):
    assert registry.stored_file_path(tmp_path, "image.png") == str(tmp_path / "image.png")
    for name in ("../secret.txt", "..\\secret.txt", "folder/child.txt", ".", "", "/etc/shadow"):
        with pytest.raises(ValueError):
            registry.stored_file_path(tmp_path, name)


def test_attachment_reference_counting_prevents_early_delete(registry, tmp_path):
    db = registry.DB(tmp_path / "attachments.db")
    try:
        first = db.save(device(registry, code="A-1"))
        second = db.save(device(registry, code="A-2"))
        attachment = tmp_path / "invoice.txt"
        attachment.write_text("Sample invoice (not a real purchase)", encoding="utf-8")
        db.add_attachment(first, str(attachment))
        ref = db.attachments(first)[0]["stored"]
        assert (Path(registry.ATT_DIR) / ref).is_file()
        db.copy_attachments(first, second)
        assert db.attachments(second)[0]["stored"] == ref
        db.delete([first])
        assert (Path(registry.ATT_DIR) / ref).is_file()
        db.delete([second])
        assert not (Path(registry.ATT_DIR) / ref).exists()
    finally:
        db.close()


def test_reminder_priority_for_overdue_repair(registry, tmp_path):
    db = registry.DB(tmp_path / "reminders.db")
    try:
        record = device(registry, status="维修")
        record["status_date"] = "2020-01-01"
        row_id = db.save(record)
        alerts = registry.device_alerts(db.get(row_id), registry.DEFAULT_CFG)
        assert any(level == "high" and "维修超期" in text for level, text in alerts)
    finally:
        db.close()
