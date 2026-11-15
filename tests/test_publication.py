"""Checks for public-release metadata and privacy safeguards."""

from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]
EMAIL_PATTERN = re.compile(r"(?i)[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}")


def test_public_project_does_not_publish_contact_email():
    readable_extensions = {".md", ".txt", ".yml", ".yaml", ".py", ".bat", ".toml"}
    excluded_folders = {"__pycache__", ".pytest_cache", ".venv", "build", "dist"}
    for path in ROOT.rglob("*"):
        if not path.is_file() or any(part in excluded_folders for part in path.parts):
            continue
        if path.suffix.lower() in readable_extensions or path.name in {"LICENSE", ".gitignore"}:
            content = path.read_text(encoding="utf-8", errors="replace")
            assert not EMAIL_PATTERN.search(content), f"Email found in {path.relative_to(ROOT)}"


def test_license_only_contains_public_copyright_name():
    license_text = (ROOT / "LICENSE").read_text(encoding="utf-8")
    assert license_text.startswith("MIT License\n\nCopyright (c) 2026 sunly56\n")
    assert "Permission is hereby granted, free of charge" in license_text
    assert "LICENSE.template" not in {p.name for p in ROOT.iterdir()}


def test_windows_release_has_only_standard_edition():
    workflow = (ROOT / ".github/workflows/windows-exe.yml").read_text(encoding="utf-8")
    assert "dist/DeviceRegistry.exe" in workflow
    assert "DeviceRegistry-Lite" not in workflow
    assert "matrix:" not in workflow
    assert not (ROOT / "build_lite.bat").exists()
    assert not (ROOT / "requirements-build-lite.txt").exists()


def test_wechat_donation_uses_a_real_image_only_when_provided():
    donation = (ROOT / "DONATE.md").read_text(encoding="utf-8")
    assert "assets/wechat-donate.png" in donation
    assert "```markdown" in donation
    assert "赞助完全自愿" in donation
    assert (ROOT / "SUPPORT.md").is_file()
