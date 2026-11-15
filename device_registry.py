#!/usr/bin/env python3
"""Device Registry: local asset inventory with optional editing protection.

The core uses Python's standard library. Excel and extended image support
are optional dependencies.
"""
import calendar
import csv
import getpass
import hashlib
import hmac
import importlib.util
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import uuid
import zipfile
from datetime import date, datetime

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# Optional libraries are lazy-loaded so basic operations and startup stay lean.
_openpyxl = None
_pillow = None


def has_excel():
    return importlib.util.find_spec("openpyxl") is not None


def excel_module():
    global _openpyxl
    if _openpyxl is None:
        try:
            import openpyxl
            _openpyxl = openpyxl
        except ImportError:
            _openpyxl = False
    return _openpyxl or None


def pillow_modules():
    global _pillow
    if _pillow is None:
        try:
            from PIL import Image, ImageTk
            _pillow = (Image, ImageTk)
        except ImportError:
            _pillow = False
    return _pillow or None

# User data stays next to the portable executable, never inside PyInstaller's
# temporary resource folder. Bundled icon resources use _MEIPASS separately.
BASE = os.path.dirname(os.path.abspath(sys.argv[0]))


def resource_path(*parts):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def set_windows_app_id():
    """Use our own taskbar group/icon instead of Python's generic icon on Windows."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "DeviceRegistry.Portable.App")
    except (AttributeError, OSError):
        pass


def set_window_icon(window):
    """Set the Tk title-bar/task-switcher icon and Windows .ico resource."""
    png = resource_path("assets", "icon.png")
    ico = resource_path("assets", "icon.ico")
    try:
        if os.path.isfile(png):
            photo = tk.PhotoImage(master=window, file=png)
            window.iconphoto(True, photo)
            window._registry_icon = photo  # keep Tcl image alive
    except (tk.TclError, OSError):
        pass
    if sys.platform == "win32" and os.path.isfile(ico):
        try:
            window.iconbitmap(default=ico)
        except tk.TclError:
            pass

DB_PATH = os.path.join(BASE, "devices.db")
IMG_DIR = os.path.join(BASE, "images")
ATT_DIR = os.path.join(BASE, "attachments")  # 采购单据等附件，与数据库同级，备份时一并打包
os.makedirs(IMG_DIR, exist_ok=True)
os.makedirs(ATT_DIR, exist_ok=True)
MAX_ATT_MB = 100
APP_VERSION = "1.3.0"

# ---------------- 高分屏(DPI)适配 ----------------
SF = 1.0  # 界面缩放系数: 屏幕DPI/96


def enable_dpi():
    """启动 Tk 之前启用 Windows 高分屏感知，优先 Per-Monitor V2。"""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def init_scale(root):
    global SF
    try:
        SF = max(1.0, root.winfo_fpixels("1i") / 96.0)
    except Exception:
        SF = 1.0
    root.tk.call("tk", "scaling", SF * 96 / 72)  # 字号(pt)随 DPI 同步放大


def S(px):
    """把按 96DPI 设计的像素值换算成当前屏幕的像素值。"""
    return int(round(px * SF))


def open_path(p):
    if sys.platform == "win32":
        os.startfile(p)
    elif sys.platform == "darwin":
        subprocess.call(["open", p])
    else:
        subprocess.call(["xdg-open", p])

def stored_file_path(directory, name):
    """仅允许数据库记录中的单个文件名，禁止跨目录读取。"""
    if not isinstance(name, str) or not name or name in (".", ".."):
        raise ValueError("文件名无效")
    if "/" in name or "\\" in name or "\0" in name or os.path.isabs(name):
        raise ValueError("存在不安全的文件路径")
    root = os.path.realpath(directory)
    path = os.path.realpath(os.path.join(root, name))
    if os.path.commonpath((root, path)) != root:
        raise ValueError("文件路径越界")
    return path


# (字段, 表头)
FIELDS = [("name", "设备名称"), ("code", "设备编码"), ("model", "型号"),
          ("company", "所属公司"), ("site", "办公地点"), ("owner", "产权人"), ("location", "存放位置"), ("status", "状态"),
          ("status_date", "状态日期"), ("qty", "数量"), ("unit", "单位"), ("value", "单价(元)"),
          ("purchase_date", "购买日期"), ("warranty_months", "保修期(月)"),
          ("remark", "备注"), ("image", "图片")]
HEAD = dict(FIELDS)
# 计算列(不入库): 由购买时间、保修时长、状态、今天日期和预警设置推算
HEAD.update(warranty="保修状态", net="当前净值(元)", alert="提醒")
# 导入时兼容旧版表头
OLD_HEAD = {"购买时间": "purchase_date", "保修时长(月)": "warranty_months", "价值": "value", "状态变更日期": "status_date"}
# 界面/导出用的显示列
DISPLAY = [(f, h) for f, h in FIELDS if f != "image"]
_i = [f for f, _ in DISPLAY].index("warranty_months") + 1
DISPLAY[_i:_i] = [("warranty", HEAD["warranty"]), ("net", HEAD["net"]), ("alert", HEAD["alert"])]
STATUSES = ["在用", "闲置", "维修", "报废", "已处置", "借出"]  # 报废=待处理，处理完毕改为"已处置"
# Language strings are UI-only. Database statuses, audit actions, and field IDs
# stay in their historical format to preserve compatibility with existing files.
UI_LANGUAGE = "zh"
UI_LANGUAGE_KEY = "ui_language"
EN_TEXT = {
    "文件": "File", "设备": "Devices", "统计": "Reports", "设置": "Settings",
    "帮助": "Help", "语言 / Language": "Language", "简体中文": "Simplified Chinese",
    "关于": "About", "退出": "Exit", "中文": "Chinese", "英语": "English",
    "设备资产管理": "Asset Manager", "免费、离线使用的设备资产管理工具。": "A free, offline device asset manager.",
    "设备名称": "Device Name", "设备编码": "Asset Code", "型号": "Model",
    "所属公司": "Company", "公司": "Company", "办公地点": "Office", "产权人": "Owner",
    "存放位置": "Location", "位置": "Location", "状态": "Status", "状态日期": "Status Date",
    "数量": "Quantity", "单位": "Unit", "单价(元)": "Unit Price (CNY)", "单价": "Unit Price",
    "购买日期": "Purchase Date", "保修期(月)": "Warranty (Months)", "保修期": "Warranty",
    "备注": "Notes", "图片": "Photo", "保修状态": "Warranty Status",
    "当前净值(元)": "Current Net Value (CNY)", "提醒": "Alerts", "提醒内容": "Alert",
    "在用": "In Use", "闲置": "Idle", "维修": "Repair", "报废": "Pending Disposal",
    "已处置": "Disposed", "借出": "Loaned", "未设置": "Not Set",
    "在保": "In Warranty", "即将到期": "Expiring Soon", "已过保": "Expired",
    "维修超期/报废待处理": "Overdue Repair / Pending Disposal",
    "折旧年限(年)": "Depreciation (Years)", "残值率(%)": "Residual Value (%)",
    "保修到期提前提醒(天)": "Warranty Alert Lead (Days)",
    "维修超过多少天提醒": "Repair Alert After (Days)",
    "新增设备": "Add Device", "编辑设备": "Edit Device", "复制设备": "Duplicate Device",
    "编辑选中": "Edit Selected", "删除选中": "Delete Selected", "复制新增": "Duplicate Added",
    "导入表单": "Import Data", "导出表单": "Export Data", "导出备份": "Export Backup",
    "恢复备份": "Restore Backup", "统计报表": "Statistics", "提醒中心": "Alert Center",
    "操作日志": "Activity Log", "修改密码": "Change Password",
    "预警与折旧设置": "Alert & Depreciation Settings",
    "解锁编辑": "Unlock Editing", "锁定(退出编辑)": "Lock (Read Only)",
    "模式: 可编辑": "Mode: Editable", "模式: 只读": "Mode: Read Only",
    "只读模式": "Read-only Mode", "颜色说明：": "Color Key:",
    "设备详情": "Device Details", "选中设备可预览图片": "Select a device to preview its photo",
    "详情": "Details", "单据": "Documents", "变更记录": "Change History",
    "文件名": "Filename", "添加日期": "Added On", "打开": "Open",
    "添加单据": "Add Document", "删除单据": "Remove Document",
    "搜索": "Search", "查询": "Search", "重置": "Reset",
    "选择日期": "Select Date", "年": "Year", "月": "Month", "今天": "Today",
    "一": "Mon", "二": "Tue", "三": "Wed", "四": "Thu", "五": "Fri", "六": "Sat", "日": "Sun",
    "设备名称、设备编码为必填项，其余信息按需填写。": "Device name and code are required; other fields are optional.",
    "选择": "Browse", "清除": "Clear", "日历选择": "Pick Date",
    "复制份数": "Copies", "多份时编码附加 -001 等序号": "Multiple copies get codes like -001, -002.",
    "采购发票、合同等单据请先保存设备，再从右侧「单据」页面添加。":
        "Save the device first, then add invoices or contracts in Documents.",
    "保存": "Save", "取消": "Cancel",
    "默认公司与提醒": "Default Company & Alerts",
    "默认所属公司：": "Default Company:",
    "留空时，如果历史记录只有一个公司，新增设备也会自动填写该公司。":
        "Leave blank to auto-fill when only one company exists in records.",
    "提醒开关": "Alert Options", "启用保修到期提醒": "Enable Warranty Expiry Alerts",
    "启用维修超期提醒": "Enable Overdue Repair Alerts", "保存设置": "Save Settings",
    "提示": "Notice", "错误": "Error", "完成": "Done", "确认": "Confirm",
    "图片保存失败": "Photo Save Failed", "文件名无效": "Invalid filename",
    "存在不安全的文件路径": "Unsafe file path", "文件路径越界": "File path outside data directory",
    "设备名称和设备编码为必填项": "Device name and asset code are required.",
    "购买日期必须填写且格式为 YYYY-MM-DD": "Purchase date is required (YYYY-MM-DD).",
    "购买日期不能晚于今天": "Purchase date cannot be in the future.",
    "数量必须是有效的正数": "Quantity must be a valid positive number.",
    "单价必须是有效的非负数": "Unit price must be a valid non-negative number.",
    "保修期必须是有效的非负数": "Warranty months must be a valid non-negative number.",
    "复制份数需为 1 到 500": "Copies must be between 1 and 500.",
    "设置密码": "Set Password", "首次使用，请设置编辑密码:": "First use: set an editing password:",
    "请再次输入密码:": "Confirm password:", "解锁": "Unlock",
    "请输入编辑密码:": "Enter the editing password:", "密码错误": "Incorrect password.",
    "请先点击“解锁编辑”输入密码": "Choose Unlock Editing and enter your password first.",
    "新密码:": "New password:", "再次输入新密码:": "Confirm new password:",
    "两次输入不一致": "Passwords do not match.", "密码已修改": "Password updated.",
    "(暂无记录)": "(No history)", "选中单个设备可查看变更记录": "Select one device to see its history.",
    "选中单个设备可查看详情和图片": "Select one device to view its details and photo.",
    "图片路径不安全": "Unsafe photo path", "(无图片)": "(No photo)", "（无）": "(None)",
    "无法预览(安装 Pillow 可支持更多格式)": "Cannot preview (install Pillow for more formats).",
    "请先在列表中选中一条设备": "Select one device from the list first.",
    "选择采购单据(可多选)": "Select Purchase Documents (multiple allowed)",
    "请先选中一个单据": "Select a document first.",
    "单据文件不存在(可能被移动或删除)": "Document not found (moved or deleted).",
    "删除选中的单据？(若其他设备共用同一文件，文件会保留)":
        "Remove the selected document? Shared files will be retained.",
    "折旧: 按购买时间起直线折旧，到期后保留残值。\n报废设备一律提醒“待处理”，处理完改为“已处置”。":
        "Straight-line depreciation starts at purchase and retains residual value.\nPending-disposal devices stay in alerts until marked Disposed.",
    "请输入有效数字(年限>0，残值率0-100，天数>=0)":
        "Enter valid values (years > 0, residual 0–100%, days >= 0).",
    "修改预警与折旧设置": "Alert & Depreciation Settings Changed",
    "设备提醒": "Device Alerts", "设备条数": "Device Count", "数量合计": "Total Quantity",
    "价值合计": "Total Value", "价值占比": "Share of Value",
    "资产总价值": "Total Asset Value", "折旧后净值": "Net Value After Depreciation",
    "按公司": "By Company", "按办公地点": "By Office", "按状态": "By Status",
    "按产权人": "By Owner", "按存放位置": "By Location",
    "(未填写)": "(Blank)", "(全部)": "(All)",
    "搜索(编码/名称/动作/内容/操作人)": "Search (Code/Name/Action/Details/User)",
    "时间": "Time", "操作人": "User", "动作": "Action", "内容": "Details",
    "请选中一条记录": "Select a record.",
    "请选中一条要复制的设备": "Select a device to duplicate.",
    "设备编码已存在，未复制：\n": "Code already exists; no devices duplicated:\n",
    "复制": "Duplicate", "复制自": "Copied from",
    "新增": "Added", "编辑": "Edited", "删除": "Deleted",
    "导入新增": "Imported (New)", "导入更新": "Imported (Updated)",
    "已复制": "Duplicated", "导入完成": "Import Complete",
    "Excel": "Excel", "CSV": "CSV", "所有文件": "All Files",
    "表单": "Data Files", "设备清单": "Device List", "备份包": "Backup ZIP",
    "导入 xlsx 需要安装 openpyxl: pip install openpyxl":
        "XLSX import needs openpyxl: pip install openpyxl",
    "CSV 编码无法识别（支持 UTF-8 和 GBK）": "Unrecognized CSV encoding (UTF-8 and GBK supported).",
    "表头需包含“设备名称”和“设备编码”":
        "The header must contain Device Name and Asset Code.",
    "备份失败": "Backup Failed", "恢复失败": "Restore Failed",
    "恢复将覆盖当前所有数据（包括密码），并创建当前数据库备份。继续？":
        "Restoring replaces all data (including the password) and backs up the current database. Continue?",
    "备份文件体积超出安全限制": "Backup exceeds the size limit.",
    "缺少 devices.db，不是有效的设备备份": "Not a valid device backup (devices.db is missing).",
    "ZIP 校验失败，备份内容可能损坏": "ZIP verification failed; backup may be damaged.",
    "备份数据库完整性校验失败": "Backup database integrity check failed.",
    "备份缺少设备数据表": "Backup has no devices table.",
    "备份已恢复；原数据库保存在 devices.db.bak，请重新解锁编辑":
        "Backup restored. Previous database saved as devices.db.bak. Unlock editing again.",
    "可添加采购发票、合同、验收单等任意文件。批量“复制设备”时，单据会一并带上。":
        "Attach invoices, contracts and acceptance documents. Duplicated devices retain their linked documents.",
    "单据文件不存在": "Document not found", "到期 ": "Expires ",
    "关闭": "Close",
    "编辑密码保护": "Editing Password Protection",
    "编辑保护设置": "Editing Protection Settings",
    "启用编辑密码保护…": "Enable Editing Password...",
    "编辑保护：已关闭": "Protection: Off (Optional)",
    "启用编辑密码保护": "Enable Editing Password Protection",
    "未启用密码保护，可直接新增、编辑、删除设备。":
        "Password protection is off. You may add, edit and delete devices directly.",
    "启用后可使用密码锁定编辑；仅控制软件内操作，不会加密数据库。":
        "When enabled, a password locks editing in the app; it does not encrypt the database.",
    "关闭密码保护需要验证现有密码。已保存的设备数据不会受到影响。":
        "Disabling protection requires your current password. Device data is preserved.",
    "模式: 可编辑（无需密码）": "Mode: Editable (No Password)",
    "输入当前密码以关闭编辑保护：": "Enter your current password to disable editing protection:",
    "确认关闭密码保护？关闭后任何能访问此程序的人都能在软件中修改记录。":
        "Disable password protection? Anyone with access to the app will be able to edit records.",
    "编辑密码保护已启用": "Editing password protection enabled.",
    "编辑密码保护已关闭": "Editing password protection disabled.",
    "请先解锁编辑后重试": "Unlock editing and try again.",
    "恢复完成；备份启用了密码保护，请使用备份的编辑密码解锁。":
        "Restored. The backup has editing password protection; unlock with its password.",
    "恢复完成；未启用密码保护，可以直接编辑。":
        "Restored. Password protection is off; editing is available immediately.",
    "恢复前数据库已另存为 devices.db.bak。":
        "The previous database is saved as devices.db.bak.",
    "拖动中间把手调整清单和详情的宽度": "Drag the center handle to resize list and details",
}
# Longest phrases first: fallback for message templates that contain numbers or filenames.
EN_PARTS = [
    ("已存在(编码不区分大小写，全公司唯一)", " already exists (codes are case-insensitive and globally unique)"),
    ("删除选中的", "Delete selected"), ("条记录？", " records?"),
    ("共 ", "Total "), (" 条 | 数量 ", " items | Quantity "),
    (" | 价值合计 ", " | Total Value "), (" | 净值合计 ", " | Net Value "),
    ("条设备提醒(其中紧急", " device alerts (urgent:"),
    ("条)，\n可在“提醒中心”查看详情。", ").\nSee Alert Center for details."),
    ("条提醒 (紫色=维修超期/报废待处理，橙黄=保修即将到期)。双击跳转到该设备。",
     " alerts (purple=overdue/pending disposal, amber=warranty expiring). Double-click to open."),
    ("显示最近 ", "Showing latest "), (" 条(最多 2000 条)", " entries (up to 2,000)"),
    ("设备 ", "Devices "), (" 条 | 数量合计 ", " | Total Quantity "),
    (" | 资产总价值 ", " | Asset Value "), (" | 折旧后净值 ", " | Depreciated Net Value "),
    ("    (筛选: ", "    (Filtered: "), ("    (全部)", "    (All)"),
    ("已复制 ", "Duplicated "), (" 条设备", " devices"),
    ("已导出 ", "Exported "), (" 条(当前筛选结果)", " records (current filter)"),
    ("新增 ", "Added "), (" 条，按编码更新 ", "; updated by code "),
    (" 条，更新 ", "; updated "), (" 条，跳过 ", "; skipped "),
    (" 条购买日期缺失或无法识别，已留空，请补充。", " purchase dates were blank or invalid; please complete them."),
    ("其中 ", "Of these, "), ("备份已保存:\n", "Backup saved:\n"),
    ("读取失败: ", "Read failed: "), ("无法打开: ", "Could not open: "),
    ("添加失败: ", "Add failed: "), ("备份未能安全恢复：", "Could not safely restore backup: "),
    ("备份包含非法目录：", "Backup contains forbidden directory: "),
    ("备份包含不安全或重复的文件名：", "Backup contains unsafe or duplicate filename: "),
    (" 超过 ", " exceeds "), ("MB，已跳过", "MB; skipped"),
    ("已过保", "Expired"), ("到期 ", "Expires "),
    ("设备编码 ", "Asset code "), ("有 ", "There are "),
    (" 等", " etc."), (" 条", " records"),
]

def set_ui_language(code):
    global UI_LANGUAGE
    UI_LANGUAGE = "en" if code == "en" else "zh"

def tr(value):
    """Translate visible interface text; never rewrite persisted device data."""
    if UI_LANGUAGE != "en" or not isinstance(value, str) or not value:
        return value
    if value in EN_TEXT:
        return EN_TEXT[value]
    m = re.fullmatch(r"在保·即将到期\(剩余(\d+)天\)", value)
    if m:
        return f"Warranty Expiring ({m.group(1)} days left)"
    m = re.fullmatch(r"在保\(剩余(\d+)天\)", value)
    if m:
        return f"In Warranty ({m.group(1)} days left)"
    m = re.fullmatch(r"保修即将到期\(剩余(\d+)天\)", value)
    if m:
        return f"Warranty Expiring ({m.group(1)} days left)"
    m = re.fullmatch(r"(维修超期|报废待处理)\(已(\d+)天\)", value)
    if m:
        return ("Repair Overdue" if m.group(1) == "维修超期" else "Pending Disposal") + f" ({m.group(2)} days)"
    if "保修即将到期" in value:
        value = value.replace("保修即将到期", "Warranty Expiring")
    if value.startswith("Device Registry · 设备资产管理"):
        return value.replace("设备资产管理", "Asset Manager")
    for cn, en in EN_PARTS:
        if cn in value:
            value = value.replace(cn, en)
    return value


def ui_status_to_db(value):
    for canonical in STATUSES:
        if value == EN_TEXT.get(canonical):
            return canonical
    return value

DEFAULT_CFG = dict(life_years=5.0, residual_pct=5.0, warn_days=30, repair_days=30)

# Stable list layout. Computed net value is read-only; photo is edited in the form.
# Former field-visibility preferences are intentionally ignored on upgrade.
LIST_COLUMNS = (
    "name", "code", "model", "site", "owner", "location", "status",
    "qty", "unit", "value", "purchase_date", "warranty_months", "net", "remark",
)
FORM_FIELDS = tuple(key for key in LIST_COLUMNS if key != "net") + ("image",)
REQUIRED_FORM_FIELDS = ("name", "code")
UI_COMPANY_KEY = "ui_default_company"
UI_SPLIT_KEY = "ui_split_ratio"
UI_EDIT_PROTECT_KEY = "ui_edit_protection"
COLUMN_BASE_WIDTHS = dict(
    name=170, code=135, model=120, company=140, site=120, owner=115,
    location=135, status=90, status_date=125, qty=75, unit=75,
    value=120, purchase_date=130, warranty_months=115, warranty=165,
    net=135, alert=205, remark=210,
)
CFG_LABELS = [("life_years", "折旧年限(年)"), ("residual_pct", "残值率(%)"),
              ("warn_days", "保修到期提前提醒(天)"), ("repair_days", "维修超过多少天提醒")]


# 行颜色 (背景, 文字): 在保=绿，即将到期=橙黄，已过保=红，维修超期/报废待处理=紫
TAG_COLORS = {"in": ("#c8efd0", "#0b5a1f"), "soon": ("#ffd966", "#6b4700"),
              "out": ("#ffc9c9", "#a31515"), "high": ("#dcc8f5", "#4a148c")}
LEGEND = [("in", "在保"), ("soon", "即将到期"), ("out", "已过保"), ("high", "维修超期/报废待处理")]


def parse_date(v):
    """解析 2024-01-31 / 2024/1/31 / 2024.1.31 / 20240131 / Excel 日期，失败返回 None。"""
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    t = str(v or "").strip()
    m = re.match(r"^(\d{4})\D+(\d{1,2})\D+(\d{1,2})", t) or re.match(r"^(\d{4})(\d{2})(\d{2})$", t)
    try:
        return date(*map(int, m.groups())) if m else None
    except ValueError:
        return None


def add_months(d, n):
    y, m = divmod(d.month - 1 + int(round(n)), 12)
    y += d.year; m += 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def warranty_info(row, warn_days=30):
    """返回 (状态文字, 到期日或None, 标签)。"""
    d = parse_date(row["purchase_date"])
    try:
        months = float(row["warranty_months"])
    except (TypeError, ValueError):
        months = 0
    if not d or months <= 0:
        return "未设置", None, ""
    end = add_months(d, months)
    left = (end - date.today()).days
    if left >= 0:
        if left <= warn_days:
            return f"在保·即将到期(剩余{left}天)", end, "soon"
        return f"在保(剩余{left}天)", end, "in"
    return "已过保", end, "out"


def net_value(row, cfg):
    """直线法折旧后的单价净值；无购买时间或价值时返回 None。"""
    d = parse_date(row["purchase_date"])
    if not d or row["value"] is None:
        return None
    total = max(cfg["life_years"], 0.1) * 12
    used = max(0, (date.today().year - d.year) * 12 + date.today().month - d.month)
    res = min(max(cfg["residual_pct"], 0), 100) / 100
    return row["value"] * max(res, 1 - (1 - res) * used / total)


def device_alerts(row, cfg):
    """返回 [(级别, 文字)]，级别 high/warn。"""
    out = []
    st = row["status"] or ""
    days = (date.today() - parse_date(row["status_date"])).days if parse_date(row["status_date"]) else 0
    if st == "报废":
        out.append(("high", f"报废待处理(已{days}天)"))
    elif st == "维修" and days > cfg["repair_days"]:
        out.append(("high", f"维修超期(已{days}天)"))
    w = warranty_info(row, cfg["warn_days"])
    if w[2] == "soon" and st not in ("报废", "已处置"):
        out.append(("warn", w[0].replace("在保·", "保修")))
    return out


# ---------------- 数据层 ----------------
class DB:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("""CREATE TABLE IF NOT EXISTS devices(
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
            code TEXT UNIQUE NOT NULL, model TEXT, owner TEXT, location TEXT,
            status TEXT, qty REAL DEFAULT 1, unit TEXT, value REAL DEFAULT 0,
            remark TEXT, image TEXT)""")
        self.conn.execute("CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT)")
        have = {r[1] for r in self.conn.execute("PRAGMA table_info(devices)")}  # 旧库自动升级
        self.conn.execute("""CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, operator TEXT, action TEXT NOT NULL,
            device_id INTEGER, code TEXT, name TEXT, detail TEXT)""")
        self.conn.execute("CREATE INDEX IF NOT EXISTS ix_log_dev ON audit_log(device_id)")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS attachments(
            id INTEGER PRIMARY KEY AUTOINCREMENT, device_id INTEGER NOT NULL,
            orig_name TEXT, stored TEXT NOT NULL, added TEXT)""")
        self.conn.execute("CREATE INDEX IF NOT EXISTS ix_att_dev ON attachments(device_id)")
        for col, typ in (("purchase_date", "TEXT"), ("warranty_months", "REAL"), ("status_date", "TEXT"),
                         ("company", "TEXT"), ("site", "TEXT")):
            if col not in have:
                self.conn.execute(f"ALTER TABLE devices ADD COLUMN {col} {typ}")
        self.conn.execute("UPDATE devices SET status_date=? WHERE status_date IS NULL OR status_date=''",
                          (date.today().isoformat(),))  # 旧数据从升级当天开始计时
        try:  # 编码不区分大小写唯一(A001 与 a001 视为同一编码)；历史数据若已冲突则跳过，仅靠程序校验
            self.conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_code_nocase ON devices(code COLLATE NOCASE)")
        except sqlite3.IntegrityError:
            pass
        self.conn.commit()

    # 兼容旧格式 salt:hash，新密码以算法和迭代次数标识。
    def has_password(self):
        return self._get("pw") is not None

    def set_password(self, pw):
        salt = os.urandom(16)
        iterations = 260_000
        hashed = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt, iterations)
        encoded = f"pbkdf2_sha256${iterations}${salt.hex()}${hashed.hex()}"
        self.conn.execute("REPLACE INTO settings VALUES('pw',?)", (encoded,))
        self.conn.commit()

    def edit_protection_enabled(self):
        """Protect a database with an existing password unless explicitly disabled."""
        return self.has_password() and self._get(UI_EDIT_PROTECT_KEY) != "0"

    def disable_edit_protection(self):
        """Called only by the GUI after verifying the existing password."""
        with self.conn:
            self.conn.execute("DELETE FROM settings WHERE k='pw'")
            self.conn.execute(
                "REPLACE INTO settings(k,v) VALUES(?,?)", (UI_EDIT_PROTECT_KEY, "0"))

    def check_password(self, pw):
        encoded = self._get("pw")
        if not encoded:
            return False
        try:
            if encoded.startswith("pbkdf2_sha256$"):
                algo, rounds, salt, expected = encoded.split("$")
                iterations = int(rounds)
                if algo != "pbkdf2_sha256" or not 100_000 <= iterations <= 2_000_000:
                    return False
            else:
                salt, expected = encoded.split(":")
                iterations = 200_000
            calculated = hashlib.pbkdf2_hmac(
                "sha256", pw.encode(), bytes.fromhex(salt), iterations)
            return hmac.compare_digest(calculated.hex(), expected)
        except (ValueError, TypeError):
            return False

    def _get(self, k):
        r = self.conn.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
        return r[0] if r else None

    def set_setting(self, key, value):
        """界面偏好也写入本地数据库，随备份迁移。"""
        self.conn.execute("REPLACE INTO settings(k,v) VALUES(?,?)", (key, value))
        self.conn.commit()

    def load_cfg(self):
        cfg = dict(DEFAULT_CFG)
        for k in cfg:
            try: cfg[k] = float(self._get("cfg_" + k))
            except (TypeError, ValueError): pass
        return cfg

    def save_cfg(self, cfg):
        for k, v in cfg.items():
            self.conn.execute("REPLACE INTO settings VALUES(?,?)", ("cfg_" + k, str(v)))
        self.conn.commit()

    def query(self, keyword="", owner="", location="", status="", company="", site=""):
        sql, args = "SELECT * FROM devices WHERE 1=1", []
        for col, val in (("company", company), ("site", site), ("owner", owner),
                         ("location", location), ("status", status)):
            if val:
                sql += f" AND {col}=?"; args.append(val)
        if keyword:
            like = f"%{keyword}%"
            sql += " AND (name LIKE ? OR code LIKE ? OR model LIKE ? OR remark LIKE ?)"
            args += [like] * 4
        return self.conn.execute(sql + " ORDER BY id", args).fetchall()

    def distinct(self, col):
        return [r[0] for r in self.conn.execute(
            f"SELECT DISTINCT {col} FROM devices WHERE {col}<>'' ORDER BY {col}")]

    def get(self, id_):
        return self.conn.execute("SELECT * FROM devices WHERE id=?", (id_,)).fetchone()

    # ---- 变更记录 ----
    def log(self, action, detail="", device_id=None):
        code = name = None
        if device_id is not None:
            r = self.get(device_id)
            if r: code, name = r["code"], r["name"]
        try:
            who = getpass.getuser()
        except Exception:
            who = ""
        self.conn.execute("INSERT INTO audit_log(ts,operator,action,device_id,code,name,detail) VALUES(?,?,?,?,?,?,?)",
                          (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), who, action, device_id, code, name, detail))
        self.conn.commit()

    def logs_for(self, device_id):
        return self.conn.execute("SELECT * FROM audit_log WHERE device_id=? ORDER BY id DESC", (device_id,)).fetchall()

    def all_logs(self, keyword="", limit=2000):
        sql, args = "SELECT * FROM audit_log", []
        if keyword:
            like = f"%{keyword}%"
            sql += " WHERE code LIKE ? OR name LIKE ? OR action LIKE ? OR detail LIKE ? OR operator LIKE ?"
            args = [like] * 5
        return self.conn.execute(sql + " ORDER BY id DESC LIMIT ?", args + [limit]).fetchall()

    @staticmethod
    def diff_fields(old, new):
        def show(f, v):
            if f in ("qty", "value", "warranty_months"):
                return fmt_num(v) if v not in (None, "") else "(空)"
            v = str(v if v is not None else "").strip()
            if f == "remark" and len(v) > 30: v = v[:30] + "…"
            return v or "(空)"
        out = []
        for f, label in FIELDS:
            if f == "status_date":
                continue
            a, b = old[f], new.get(f)
            if f in ("qty", "value", "warranty_months"):
                if abs(num(a) - num(b)) < 1e-9: continue
            elif str(a or "").strip() == str(b or "").strip():
                continue
            if f == "image":
                out.append("图片: " + ("已清除" if not b else "已更换" if a else "已添加"))
            else:
                out.append(f"{label}: {show(f, a)} → {show(f, b)}")
        return out

    def save(self, d, id_=None, src="", note=""):
        """保存设备并写变更记录；返回设备 id。src: ""=手工, "复制", "导入"。"""
        cols = [f for f, _ in FIELDS]
        d = dict(d)
        for field in ("qty", "value", "warranty_months"):
            value = d.get(field)
            if value in (None, ""):
                value = 1.0 if field == "qty" else 0.0
            try:
                value = float(value)
            except (ValueError, TypeError, OverflowError):
                raise ValueError(f"{HEAD[field]}不是有效数字")
            if not math.isfinite(value) or value < 0 or (field == "qty" and value <= 0):
                raise ValueError(f"{HEAD[field]}超出有效范围")
            d[field] = value
        today = date.today().isoformat()
        given = parse_date(d.get("status_date"))
        old = self.get(id_) if id_ is not None else None
        if given:
            d["status_date"] = given.isoformat()
        elif old and old["status"] == d.get("status") and old["status_date"]:
            d["status_date"] = old["status_date"]  # 状态没变，保持原日期
        else:
            d["status_date"] = today  # 新增或状态变更，重新计时
        if old is None:
            cur = self.conn.execute(f"INSERT INTO devices({','.join(cols)}) VALUES({','.join('?'*len(cols))})",
                                    [d.get(c) for c in cols])
            id_ = cur.lastrowid
            self.conn.commit()
            self.log({"": "新增", "复制": "复制新增", "导入": "导入新增"}[src],
                     "；".join(x for x in (note, f"编码 {d['code']}，名称 {d['name']}") if x), id_)
        else:
            changes = self.diff_fields(old, d)
            self.conn.execute(f"UPDATE devices SET {','.join(c+'=?' for c in cols)} WHERE id=?",
                              [d.get(c) for c in cols] + [id_])
            self.conn.commit()
            if changes:
                self.log("导入更新" if src == "导入" else "编辑", "\n".join(changes), id_)
        return id_

    def delete(self, ids):
        for i in ids:
            r = self.get(i)
            if not r: continue
            stored = [x[0] for x in self.conn.execute("SELECT stored FROM attachments WHERE device_id=?", (i,))]
            info = "；".join(f"{lb}:{r[f]}" for f, lb in (("company", "公司"), ("site", "地点"), ("owner", "产权人"),
                                                          ("location", "位置"), ("status", "状态")) if r[f])
            self.log("删除", info + (f"；含单据 {len(stored)} 个" if stored else ""), i)
            self.conn.execute("DELETE FROM attachments WHERE device_id=?", (i,))
            self.conn.execute("DELETE FROM devices WHERE id=?", (i,))
            for st in stored: self._gc_file(st)
        self.conn.commit()

    def find_code(self, code):
        return self.conn.execute("SELECT id FROM devices WHERE code=? COLLATE NOCASE", (code.strip(),)).fetchone()

    # ---- 采购单据/附件 ----
    def attachments(self, device_id):
        return self.conn.execute("SELECT * FROM attachments WHERE device_id=? ORDER BY id", (device_id,)).fetchall()

    def add_attachment(self, device_id, src):
        stored = uuid.uuid4().hex + os.path.splitext(src)[1].lower()  # 保留扩展名，双击能用正确的软件打开
        shutil.copy2(src, stored_file_path(ATT_DIR, stored))
        self.conn.execute("INSERT INTO attachments(device_id,orig_name,stored,added) VALUES(?,?,?,?)",
                          (device_id, os.path.basename(src), stored, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.conn.commit()
        self.log("添加单据", os.path.basename(src), device_id)

    def remove_attachment(self, att_id):
        r = self.conn.execute("SELECT * FROM attachments WHERE id=?", (att_id,)).fetchone()
        if not r: return
        self.conn.execute("DELETE FROM attachments WHERE id=?", (att_id,)); self.conn.commit()
        self._gc_file(r["stored"])
        self.log("删除单据", r["orig_name"], r["device_id"])

    def copy_attachments(self, src_id, dst_id):
        """复制设备时共用同一份单据文件(同一批采购共用一张发票)，只复制记录不复制文件。"""
        self.conn.execute("INSERT INTO attachments(device_id,orig_name,stored,added) "
                          "SELECT ?,orig_name,stored,added FROM attachments WHERE device_id=?", (dst_id, src_id))
        self.conn.commit()

    def _gc_file(self, stored):  # 没有任何记录引用时才删除物理文件
        if not self.conn.execute("SELECT 1 FROM attachments WHERE stored=?", (stored,)).fetchone():
            try:
                os.remove(stored_file_path(ATT_DIR, stored))
            except (OSError, ValueError):
                pass

    def close(self):
        self.conn.close()


def num(v, default=0.0):
    try:
        value = float(str(v).replace(",", "").strip())
        return value if math.isfinite(value) else default
    except (ValueError, TypeError, OverflowError):
        return default


def fmt_num(x):
    return "" if x is None else (str(int(x)) if float(x) == int(x) else f"{x:.2f}")


# ---------------- 日历选择 ----------------
class DatePicker(tk.Toplevel):
    """简易日历弹窗，选中后把 YYYY-MM-DD 写入 var；未来日期不可选。"""
    def __init__(self, master, var):
        super().__init__(master)
        set_window_icon(self)
        self.title(tr("选择日期")); self.transient(master); self.resizable(False, False)
        self.var = var
        cur = parse_date(var.get()) or date.today()
        self.y = tk.IntVar(value=cur.year); self.m = tk.IntVar(value=cur.month); self.sel = cur
        top = ttk.Frame(self, padding=(8, 8, 8, 2)); top.pack(fill="x")
        ttk.Button(top, text=tr("<"), width=3, command=lambda: self.shift(-1)).pack(side="left")
        ttk.Combobox(top, textvariable=self.y, values=list(range(date.today().year, 1989, -1)),
                     width=6, state="readonly").pack(side="left", padx=4)
        ttk.Label(top, text=tr("年")).pack(side="left")
        ttk.Combobox(top, textvariable=self.m, values=list(range(1, 13)), width=4, state="readonly").pack(side="left", padx=4)
        ttk.Label(top, text=tr("月")).pack(side="left")
        ttk.Button(top, text=tr(">"), width=3, command=lambda: self.shift(1)).pack(side="right")
        for w in top.winfo_children():
            if isinstance(w, ttk.Combobox): w.bind("<<ComboboxSelected>>", lambda _: self.draw())
        self.grid_f = ttk.Frame(self, padding=8); self.grid_f.pack()
        ttk.Button(self, text=tr("今天"), command=lambda: self.pick(date.today())).pack(pady=(0, 8))
        self.draw(); self.grab_set(); self.wait_window()

    def shift(self, n):
        y, m = self.y.get(), self.m.get() + n
        if m < 1: y, m = y - 1, 12
        if m > 12: y, m = y + 1, 1
        self.y.set(y); self.m.set(m); self.draw()

    def draw(self):
        for w in self.grid_f.winfo_children(): w.destroy()
        for i, h in enumerate("一二三四五六日"):
            ttk.Label(self.grid_f, text=tr(h), anchor="center", width=4).grid(row=0, column=i)
        today = date.today()
        for r, week in enumerate(calendar.monthcalendar(self.y.get(), self.m.get()), 1):
            for c, day in enumerate(week):
                if not day: continue
                d = date(self.y.get(), self.m.get(), day)
                b = tk.Button(self.grid_f, text=tr(str(day)), width=4, relief="flat",
                              bg="#4a90d9" if d == self.sel else "#f5f5f5",
                              fg="white" if d == self.sel else "black",
                              state="disabled" if d > today else "normal",
                              command=lambda d=d: self.pick(d))
                b.grid(row=r, column=c, padx=1, pady=1)

    def pick(self, d):
        self.var.set(d.isoformat()); self.destroy()


# ---------------- 设备录入窗口 ----------------
class EditDialog(tk.Toplevel):
    def __init__(self, master, row=None, copy=False):
        super().__init__(master)
        set_window_icon(self)
        self.copy = copy
        self.result = None
        self.image = row["image"] if row else ""
        self.pending_image = None
        self.count = tk.StringVar(value="1")
        self.vars = {}
        self.field_rows = {}
        self.original = {key: row[key] for key, _ in FIELDS} if row else {}
        self.title(tr("复制设备" if copy else "编辑设备" if row else "新增设备"))
        self.transient(master)
        self.resizable(True, True)
        self.minsize(S(470), S(340))

        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        ttk.Label(top, text=tr("设备名称、设备编码为必填项，其余信息按需填写。"),
                  foreground="#395879").pack(anchor="w")

        outer = ttk.Frame(self, padding=(10, 0, 10, 0))
        outer.pack(fill="both", expand=True)
        view_height = min(590, max(300, len(FORM_FIELDS) * 37 + 35))
        self.canvas = tk.Canvas(outer, highlightthickness=0, height=S(view_height),
                                width=S(515))
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.form_frame = ttk.Frame(self.canvas, padding=(4, 8))
        inner_id = self.canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        self.form_frame.bind("<Configure>",
                             lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(inner_id, width=e.width))
        self.bind("<MouseWheel>", self._scroll_form)
        self.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-2, "units"))
        self.bind("<Button-5>", lambda e: self.canvas.yview_scroll(2, "units"))
        self.form_frame.grid_columnconfigure(0, weight=1)

        for index, f in enumerate(FORM_FIELDS):
            label = HEAD[f]
            line = ttk.Frame(self.form_frame)
            line.grid(row=index, column=0, sticky="ew", pady=3)
            line.grid_columnconfigure(1, weight=1)
            self.field_rows[f] = line
            ttk.Label(line, text=tr(label) + ("*" if f in ("name", "code") else ""),
                      width=12, anchor="e").grid(row=0, column=0, sticky="e")

            if f == "remark":
                self.remark = tk.Text(line, width=30, height=3, wrap="word")
                self.remark.grid(row=0, column=1, padx=(8, 4), sticky="ew")
                if row and row["remark"]:
                    self.remark.insert("1.0", row["remark"])
                continue
            if f == "image":
                box = ttk.Frame(line)
                box.grid(row=0, column=1, padx=(8, 4), sticky="w")
                self.imglbl = ttk.Label(box, text=tr(self.image or "（无）"), width=23)
                self.imglbl.pack(side="left")
                ttk.Button(box, text=tr("选择"), width=6, command=self.pick).pack(side="left", padx=3)
                ttk.Button(box, text=tr("清除"), width=6, command=self.clear).pack(side="left")
                continue

            if row and row[f] is not None and not (copy and f == "code"):
                value = fmt_num(row[f]) if f in ("qty", "value", "warranty_months") else str(row[f])
            elif f == "qty":
                value = "1"
            elif f == "purchase_date":
                value = date.today().isoformat()
            elif f == "company":
                value = master.default_company_value()
            elif f == "status":
                value = "在用"
            else:
                value = ""
            if f == "status":
                value = tr(value)
            var = tk.StringVar(value=value)
            self.vars[f] = var

            if f == "status":
                inp = ttk.Combobox(line, textvariable=var, values=[tr(s) for s in STATUSES], width=28)
            elif f in ("company", "site", "owner", "location"):
                options = master.db.distinct(f)
                if f == "company" and value and value not in options:
                    options.insert(0, value)
                inp = ttk.Combobox(line, textvariable=var, values=options, width=28)
            elif f == "purchase_date":
                inp = ttk.Frame(line)
                ttk.Entry(inp, textvariable=var, width=19).pack(side="left", fill="x", expand=True)
                ttk.Button(inp, text=tr("日历选择"), width=9,
                           command=lambda v=var: DatePicker(self, v)).pack(side="left", padx=(6, 0))
            else:
                inp = ttk.Entry(line, textvariable=var, width=33)
            inp.grid(row=0, column=1, padx=(8, 4), sticky="ew")
            if f == "name":
                self.name_input = inp

            if copy and f == "code":
                copies = ttk.Frame(line)
                copies.grid(row=1, column=1, padx=8, sticky="w", pady=(6, 0))
                ttk.Label(copies, text=tr("复制份数")).pack(side="left")
                ttk.Spinbox(copies, from_=1, to=500, textvariable=self.count, width=6).pack(side="left", padx=6)
                ttk.Label(copies, text=tr("多份时编码附加 -001 等序号")).pack(side="left")

        ttk.Label(self, text=tr("采购发票、合同等单据请先保存设备，再从右侧「单据」页面添加。"),
                  padding=(12, 4), foreground="gray").pack(anchor="w")
        actions = ttk.Frame(self, padding=(12, 7, 12, 12))
        actions.pack(fill="x")
        ttk.Button(actions, text=tr("保存"), command=self.ok).pack(side="left", expand=True, padx=8)
        ttk.Button(actions, text=tr("取消"), command=self.destroy).pack(side="left", expand=True, padx=8)
        self.name_input.focus_set()
        self.grab_set()
        self.wait_window()

    def _scroll_form(self, e):
        if e.delta:
            self.canvas.yview_scroll(-int(e.delta / 120) or (-1 if e.delta > 0 else 1), "units")

    def pick(self):
        path = filedialog.askopenfilename(
            parent=self, filetypes=[(tr("图片"), "*.png *.jpg *.jpeg *.gif *.bmp"), (tr("所有文件"), "*.*")])
        if path:
            self.pending_image = path
            self.imglbl.config(text=tr(os.path.basename(path)))

    def clear(self):
        self.pending_image = None
        self.image = ""
        self.imglbl.config(text=tr("（无）"))

    def ok(self):
        # Merge edited fields with the original row to preserve legacy data
        # (for example, company). The status date is managed by DB.save().
        d = dict(self.original)
        d.pop("status_date", None)
        d.update({key: var.get().strip() for key, var in self.vars.items()})
        d.setdefault("company", self.master.default_company_value())
        d["status"] = ui_status_to_db(d.get("status", ""))
        if not d["name"] or not d["code"]:
            messagebox.showwarning(tr("提示"), tr("设备名称和设备编码为必填项"), parent=self)
            return
        purchased = parse_date(d["purchase_date"])
        if not purchased:
            messagebox.showwarning(tr("提示"), tr("购买日期必须填写且格式为 YYYY-MM-DD"), parent=self)
            return
        if purchased > date.today():
            messagebox.showwarning(tr("提示"), tr("购买日期不能晚于今天"), parent=self)
            return
        d["purchase_date"] = purchased.isoformat()

        for field, label in (("qty", "数量"), ("value", "单价"), ("warranty_months", "保修期")):
            raw = d[field]
            try:
                value = float(raw.replace(",", "")) if raw else (1.0 if field == "qty" else 0.0)
                if not math.isfinite(value) or value < 0 or (field == "qty" and value <= 0):
                    raise ValueError
            except ValueError:
                messagebox.showwarning(
                    tr("提示"), tr(f"{label}必须是有效的{'正数' if field == 'qty' else '非负数'}"), parent=self)
                return
            d[field] = value
        if self.copy:
            try:
                copies = int(self.count.get())
                if not 1 <= copies <= 500:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(tr("提示"), tr("复制份数需为 1 到 500"), parent=self)
                return
            d["copies"] = copies
        d["remark"] = self.remark.get("1.0", "end").strip()
        if self.pending_image:
            try:
                filename = uuid.uuid4().hex + os.path.splitext(self.pending_image)[1].lower()
                shutil.copy2(self.pending_image, os.path.join(IMG_DIR, filename))
                self.image = filename
            except OSError as exc:
                messagebox.showerror(tr("图片保存失败"), tr(str(exc)), parent=self)
                return
        d["image"] = self.image
        self.result = d
        self.destroy()


# ---------------- 主界面 ----------------
class App(tk.Tk):
    def __init__(self):
        set_windows_app_id()
        super().__init__()
        init_scale(self)  # DPI 已在启动前开启，这里读取缩放系数
        import tkinter.font as tkfont
        for n in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            f = tkfont.nametofont(n)
            f.configure(size=11)
            if sys.platform == "win32":
                f.configure(family="Microsoft YaHei UI")
        st = ttk.Style()
        st.configure("Treeview", rowheight=S(30))
        self._heading_font = tkfont.Font(
            root=self, family="Microsoft YaHei UI" if sys.platform == "win32" else "TkDefaultFont",
            size=10, weight="bold")
        self._cell_font = tkfont.nametofont("TkDefaultFont")
        st.configure("Treeview.Heading", font=self._heading_font,
                     background="#dce9f7", foreground="#1f3a5f", padding=(S(5), S(7)), relief="flat")
        st.map("Treeview.Heading", background=[("active", "#c5dbf2")])
        self.title(tr(f"Device Registry · 设备资产管理 v{APP_VERSION}"))
        self.geometry(f"{S(1400)}x{S(800)}")
        try:  # 启动自动最大化
            self.state("zoomed")
        except tk.TclError:
            try: self.attributes("-zoomed", True)
            except tk.TclError: pass
        self.db = DB(DB_PATH)
        self.language = self.db._get(UI_LANGUAGE_KEY) or "zh"
        set_ui_language(self.language)
        self.title(tr(f"Device Registry · 设备资产管理 v{APP_VERSION}"))
        set_window_icon(self)
        self.cfg = self.db.load_cfg()
        self._load_preferences()
        self.protection_enabled = self.db.edit_protection_enabled()
        self.unlocked = not self.protection_enabled
        self._photo = None
        self.build()
        self.refresh()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._startup_alert_job = self.after(400, lambda: self.show_alerts(startup=True))

    def _load_preferences(self):
        self.default_company = self.db._get(UI_COMPANY_KEY) or ""
        self.warranty_alerts_on = self.db._get("ui_warranty_alerts") != "0"
        self.repair_alerts_on = self.db._get("ui_repair_alerts") != "0"

    def change_language(self, code):
        """Switch UI in place while retaining filters, selection and stored data."""
        code = "en" if code == "en" else "zh"
        if code == self.language:
            return
        filters = {key: (ui_status_to_db(v.get()) if key == "status" else v.get())
                   for key, v in self.fv.items()}
        keyword = self.kw.get()
        selected = tuple(self.tree.selection())
        self._save_split()
        self.language = code
        self.db.set_setting(UI_LANGUAGE_KEY, code)
        set_ui_language(code)
        self.title(tr(f"Device Registry · 设备资产管理 v{APP_VERSION}"))
        for child in self.winfo_children():
            child.destroy()
        self.build()
        self.kw.set(keyword)
        for key, value in filters.items():
            self.fv[key].set(tr(value) if key == "status" else value)
        self.refresh()
        valid = [item for item in selected if self.tree.exists(item)]
        if valid:
            self.tree.selection_set(valid)
            self.show_detail()

    def default_company_value(self):
        if self.default_company:
            return self.default_company
        companies = self.db.distinct("company")
        return companies[0] if len(companies) == 1 else ""

    def _rebuild_detail_fields(self):
        for child in self.detail.winfo_children():
            child.destroy()
        self.dv = {}
        for index, field in enumerate(LIST_COLUMNS):
            label = HEAD[field]
            ttk.Label(self.detail, text=tr(label) + (": " if self.language == "en" else "："),
                      foreground="gray").grid(row=index, column=0, sticky="ne", pady=2)
            value = ttk.Label(self.detail, text=tr(""), wraplength=S(245), justify="left")
            value.grid(row=index, column=1, sticky="nw", padx=(4, 0), pady=2)
            self.dv[field] = value
        self.detail.grid_columnconfigure(1, weight=1)

    def _resize_detail_labels(self, width):
        max_width = max(S(125), width - S(175))
        for label in self.dv.values():
            label.configure(wraplength=max_width)

    def _schedule_split_restore(self, _event=None):
        # Reapply the saved proportion after Tk finishes arranging the panes.
        pending = getattr(self, "_restore_split_job", None)
        if pending is not None:
            self.after_cancel(pending)
        self._restore_split_job = self.after_idle(self._restore_split)

    def _restore_split(self):
        self._restore_split_job = None
        width = self.split.winfo_width()
        if width <= S(500):
            return
        saved = self.db._get(UI_SPLIT_KEY)
        try:
            ratio = float(saved) if saved else 0.68
        except (TypeError, ValueError):
            ratio = 0.68
        # Native sash works on the entire separator (visible handle included).
        ratio = max(0.37, min(0.82, ratio))
        try:
            self.split.sash_place(0, int(width * ratio), 0)
        except tk.TclError:
            pass

    def _save_split(self):
        try:
            width = self.split.winfo_width()
            if width > S(500):
                ratio = self.split.sash_coord(0)[0] / width
                if 0.30 <= ratio <= 0.92:
                    self.db.set_setting(UI_SPLIT_KEY, f"{ratio:.5f}")
        except (tk.TclError, AttributeError):
            pass

    def active_alerts(self, row):
        alerts = device_alerts(row, self.cfg)
        return [
            (level, text) for level, text in alerts
            if not (level == "warn" and not self.warranty_alerts_on)
            and not (text.startswith("维修超期") and not self.repair_alerts_on)
        ]

    def _schedule_column_layout(self, _event=None):
        """Debounce automatic resizing on pane/window changes."""
        if getattr(self, "_column_resize_job", None) is not None:
            try:
                self.after_cancel(self._column_resize_job)
            except tk.TclError:
                pass
        self._column_resize_job = self.after(70, self._autofit_columns)

    def _autofit_columns(self):
        """Fit headers + representative data, never truncate a complete title.

        If all column minima don't fit, their minimum widths are retained and
        the existing horizontal scrollbar handles the overflow. Otherwise,
        available width is shared proportionally across columns.
        """
        self._column_resize_job = None
        if not hasattr(self, "tree") or not self.tree.winfo_exists():
            return
        keys = list(self.tree["columns"])
        if not keys:
            return
        available = self.tree.winfo_width() - S(4)
        if available < S(80):
            return
        sample_iids = self.tree.get_children()[:60]
        mins, prefs = [], []
        for key in keys:
            title_px = self._heading_font.measure(tr(HEAD[key])) + S(28)
            minimum = max(S(58), title_px)
            max_data = 0
            for iid in sample_iids:
                text = str(self.tree.set(iid, key) or "")
                # Very long notes and alerts scroll within a sensible column size.
                max_data = max(max_data, self._cell_font.measure(text[:64]))
            target = max(minimum, S(int(COLUMN_BASE_WIDTHS.get(key, 120) * 0.82)),
                         min(S(290), max_data + S(20)))
            mins.append(minimum)
            prefs.append(min(S(310), target))
        min_total = sum(mins)
        pref_total = sum(prefs)
        if min_total >= available:
            widths = mins
        elif pref_total >= available:
            capacity = pref_total - min_total
            fraction = (available - min_total) / capacity if capacity else 0
            widths = [int(m + (p - m) * fraction)
                      for m, p in zip(mins, prefs)]
        else:
            spare = available - pref_total
            widths = [int(p + spare * p / pref_total) for p in prefs]
        # Correct integer rounding without ever dropping below header minima.
        if sum(widths) < available and min_total < available:
            widths[-1] += available - sum(widths)
        for key, minimum, width in zip(keys, mins, widths):
            self.tree.column(key, minwidth=minimum, width=width,
                             stretch=False, anchor="w")

    def open_options(self):
        """Edit application defaults without changing the visible asset fields."""
        win = tk.Toplevel(self)
        set_window_icon(win)
        win.title(tr("默认公司与提醒"))
        win.transient(self)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=tr("默认所属公司：")).pack(anchor="w")
        company = tk.StringVar(master=win, value=self.default_company)
        ttk.Entry(frame, textvariable=company, width=40).pack(anchor="w", pady=(5, 7))
        ttk.Label(
            frame,
            text=tr("留空时，如果历史记录只有一个公司，新增设备也会自动填写该公司。"),
            wraplength=S(450), foreground="gray",
        ).pack(anchor="w", pady=(0, 15))

        alert_frame = ttk.LabelFrame(frame, text=tr("提醒开关"), padding=10)
        alert_frame.pack(fill="x")
        warranty = tk.BooleanVar(master=win, value=self.warranty_alerts_on)
        repair = tk.BooleanVar(master=win, value=self.repair_alerts_on)
        ttk.Checkbutton(
            alert_frame, text=tr("启用保修到期提醒"), variable=warranty,
        ).pack(anchor="w")
        ttk.Checkbutton(
            alert_frame, text=tr("启用维修超期提醒"), variable=repair,
        ).pack(anchor="w", pady=(6, 0))

        actions = ttk.Frame(frame)
        actions.pack(anchor="e", fill="x", pady=(16, 0))

        def save():
            self.db.set_setting(UI_COMPANY_KEY, company.get().strip())
            self.db.set_setting("ui_warranty_alerts", "1" if warranty.get() else "0")
            self.db.set_setting("ui_repair_alerts", "1" if repair.get() else "0")
            self._load_preferences()
            self.refresh()
            win.destroy()

        ttk.Button(actions, text=tr("保存设置"), command=save).pack(side="right", padx=(7, 0))
        ttk.Button(actions, text=tr("取消"), command=win.destroy).pack(side="right")
        return win

    def build_menu(self):
        # 自绘菜单栏: Windows 原生菜单栏字体由系统决定，无法放大
        mfont = ("Microsoft YaHei UI" if sys.platform == "win32" else "", 12)
        mb = tk.Frame(self, bg="#f0f0f0"); mb.pack(side="top", fill="x")
        self.edit_items = []  # (menu, label): 需解锁才可用
        def add(menu, label, cmd, edit=False):
            menu.add_command(label=tr(label), command=cmd)
            if edit: self.edit_items.append((menu, tr(label)))
        def new_menu(title):
            mbtn = tk.Menubutton(mb, text=tr(title), font=mfont, relief="flat", padx=S(14), pady=S(5),
                                 bg="#f0f0f0", activebackground="#d9e8fb")
            mbtn.pack(side="left")
            m = tk.Menu(mbtn, tearoff=0, font=mfont); mbtn["menu"] = m
            return m
        m = new_menu("文件")
        add(m, "导入表单", self.import_file, True)
        add(m, "导出表单", self.export_file)
        m.add_separator()
        add(m, "导出备份", self.backup, True)
        add(m, "恢复备份", self.restore, True)
        m.add_separator()
        add(m, "退出", self.on_close)
        m = new_menu("设备")
        add(m, "新增设备", self.add, True)
        add(m, "编辑选中", self.edit, True)
        add(m, "复制设备", self.copy_device, True)
        add(m, "删除选中", self.remove, True)
        m = new_menu("统计")
        add(m, "统计报表", self.show_stats)
        add(m, "提醒中心", self.show_alerts)
        add(m, "操作日志", self.show_log)
        m = new_menu("设置")
        self.lock_menu = m
        add(m, "解锁编辑", self.toggle_lock)
        add(m, "编辑保护设置", self.configure_edit_protection)
        add(m, "修改密码", self.change_pw, True)
        add(m, "预警与折旧设置", self.edit_cfg, True)
        add(m, "默认公司与提醒", self.open_options)
        m = new_menu("语言 / Language")
        self.language_choice = tk.StringVar(value=self.language)
        m.add_radiobutton(label="简体中文", variable=self.language_choice, value="zh",
                          command=lambda: self.change_language("zh"))
        m.add_radiobutton(label="English", variable=self.language_choice, value="en",
                          command=lambda: self.change_language("en"))
        m = new_menu("帮助")
        add(m, "关于", lambda: messagebox.showinfo(
            tr("关于"), tr(f"Device Registry v{APP_VERSION}\n\n免费、离线使用的设备资产管理工具。")))
        ttk.Separator(self, orient="horizontal").pack(side="top", fill="x")

    def build(self):
        self.build_menu()
        # 菜单栏下方的功能按键条
        tb = ttk.Frame(self, padding=(8, 6)); tb.pack(side="top", fill="x")
        self.edit_btns = []
        def btn(text, cmd, edit=False):
            b = ttk.Button(tb, text=tr(text), command=cmd); b.pack(side="left", padx=3, ipady=3)
            if edit: self.edit_btns.append(b)
            return b
        def sep():
            ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        self.lock_btn = btn("解锁编辑", self.toggle_lock)
        sep()
        btn("新增设备", self.add, True)
        btn("编辑选中", self.edit, True)
        btn("复制设备", self.copy_device, True)
        btn("删除选中", self.remove, True)
        sep()
        btn("导入表单", self.import_file, True)
        btn("导出表单", self.export_file)
        btn("统计报表", self.show_stats)
        self.alert_btn = btn("提醒中心", self.show_alerts)
        sep()
        btn("导出备份", self.backup, True)  # 恢复备份、修改密码在菜单里，避免窄屏时按键条放不下
        self.mode_lbl = ttk.Label(tb, text=tr(""), foreground="gray"); self.mode_lbl.pack(side="right", padx=8)
        ttk.Separator(self, orient="horizontal").pack(side="top", fill="x")

        # 最下方状态栏: 版权信息靠右
        foot = ttk.Frame(self, padding=(8, 3)); foot.pack(side="bottom", fill="x")
        ttk.Label(foot, text=tr(f"Device Registry v{APP_VERSION}"), foreground="gray").pack(side="right")
        ttk.Label(foot, text=tr("颜色说明：")).pack(side="left")
        for tag, text in LEGEND:
            tk.Label(foot, text=f" {tr(text)} ", bg=TAG_COLORS[tag][0], fg=TAG_COLORS[tag][1],
                     font=("", 10)).pack(side="left", padx=3)

        # 左侧清单、右侧详情均由 Panedwindow 管理，可拖动中间分隔条。
        self.split = tk.PanedWindow(
            self, orient=tk.HORIZONTAL, showhandle=True,
            sashwidth=S(11), sashpad=S(3), sashrelief=tk.RAISED,
            handlesize=S(13), handlepad=S(26),
            opaqueresize=True, borderwidth=0, relief=tk.FLAT,
            background="#b6c9db", cursor="sb_h_double_arrow")
        self.split.pack(side="top", fill="both", expand=True)
        main = ttk.Frame(self.split, padding=8)
        right = ttk.Frame(self.split, padding=8, width=S(430))
        self.split.add(main, minsize=S(320), stretch="always")
        self.split.add(right, minsize=S(295), stretch="always")
        self.split.bind("<ButtonRelease-1>", lambda _e: self._save_split())
        self.split.bind("<Configure>", self._schedule_split_restore)
        ttk.Label(right, text=tr("设备详情"), font=("", 14, "bold")).pack(pady=(0, 3))
        ttk.Label(right, text=tr("拖动中间把手调整清单和详情的宽度"),
                  font=("", 9), foreground="#62768d").pack(pady=(0, 5))
        pf = ttk.Frame(right, height=S(190)); pf.pack(fill="x"); pf.pack_propagate(False)
        self.preview = ttk.Label(pf, text=tr("选中设备可预览图片"), anchor="center", relief="groove")
        self.preview.pack(fill="both", expand=True)
        self.nb = nb = ttk.Notebook(right); nb.pack(fill="both", expand=True, pady=(8, 0))
        bg = ttk.Style().lookup("TFrame", "background") or "#f0f0f0"

        # 页1: 详情(含备注)，内容多时可滚动
        t1 = ttk.Frame(nb); nb.add(t1, text=tr("详情"))
        cv = tk.Canvas(t1, highlightthickness=0, bg=bg); vsb = ttk.Scrollbar(t1, command=cv.yview)
        cv.configure(yscrollcommand=vsb.set); vsb.pack(side="right", fill="y"); cv.pack(side="left", fill="both", expand=True)
        self.detail = ttk.Frame(cv); wid = cv.create_window((0, 0), window=self.detail, anchor="nw")
        self.detail.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfigure(wid, width=e.width))
        wheel = lambda e: cv.yview_scroll(-1 if e.delta > 0 else 1, "units")
        cv.bind("<Enter>", lambda e: cv.bind_all("<MouseWheel>", wheel))
        cv.bind("<Leave>", lambda e: cv.unbind_all("<MouseWheel>"))
        self.dv = {}
        self._rebuild_detail_fields()
        right.bind("<Configure>", lambda e: self._resize_detail_labels(e.width))

        # 页2: 采购单据
        t2 = ttk.Frame(nb, padding=4); nb.add(t2, text=tr("单据"))
        self.att_tree = ttk.Treeview(t2, columns=("n", "d"), show="headings", height=6)
        self.att_tree.heading("n", text=tr("文件名")); self.att_tree.heading("d", text=tr("添加日期"))
        self.att_tree.column("n", width=S(250)); self.att_tree.column("d", width=S(90))
        self.att_tree.pack(fill="both", expand=True)
        self.att_tree.bind("<Double-1>", lambda _: self.att_open())
        ab = ttk.Frame(t2); ab.pack(fill="x", pady=(6, 0))
        ttk.Button(ab, text=tr("打开"), command=self.att_open).pack(side="left")
        for text, cmd in (("添加单据", self.att_add), ("删除单据", self.att_remove)):
            b = ttk.Button(ab, text=tr(text), command=cmd); b.pack(side="left", padx=(6, 0)); self.edit_btns.append(b)
        ttk.Label(t2, text=tr("可添加采购发票、合同、验收单等任意文件。批量“复制设备”时，单据会一并带上。"),
                  foreground="gray", wraplength=S(390)).pack(anchor="w", pady=(6, 0))

        # 页3: 变更记录
        t3 = ttk.Frame(nb, padding=4); nb.add(t3, text=tr("变更记录"))
        self.log_box = tk.Text(t3, wrap="word", state="disabled", relief="flat", bg=bg, height=8)
        ls = ttk.Scrollbar(t3, command=self.log_box.yview); self.log_box.configure(yscrollcommand=ls.set)
        ls.pack(side="right", fill="y"); self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_configure("h", font=("", 10, "bold"), foreground="#1f3a5f", spacing1=8)

        # 搜索区域只有一行；宽度不足时使用下方横向滚动条，而不是换行。
        bar_outer = ttk.Frame(main)
        bar_outer.pack(fill="x")
        self.search_canvas = tk.Canvas(bar_outer, highlightthickness=0, height=S(42),
                                      background=bg)
        self.search_canvas.pack(fill="x")
        bar = ttk.Frame(self.search_canvas)
        self.search_window = self.search_canvas.create_window((0, 0), window=bar, anchor="nw")
        bar.bind("<Configure>", lambda e: self.search_canvas.configure(
            scrollregion=(0, 0, e.width, e.height)))
        search_scroll = ttk.Scrollbar(bar_outer, orient="horizontal",
                                     command=self.search_canvas.xview)
        search_scroll.pack(fill="x")
        self.search_canvas.configure(xscrollcommand=search_scroll.set)
        self.search_canvas.bind("<Shift-MouseWheel>",
                                lambda e: self.search_canvas.xview_scroll(
                                    -1 if e.delta > 0 else 1, "units"))
        self.kw = tk.StringVar()
        self.fv = {key: tk.StringVar() for key in
                   ("company", "site", "owner", "location", "status")}
        ttk.Label(bar, text=tr("搜索")).pack(side="left", padx=(2, 0))
        self.kw_entry = ttk.Entry(bar, textvariable=self.kw, width=16)
        self.kw_entry.pack(side="left", padx=(4, 6), pady=4)
        self.kw_entry.bind("<Return>", lambda _: self.refresh())
        self.kw_entry.bind("<KeyRelease>", lambda _: self.refresh())
        self.cb = {}

        def combo(label, key, width):
            ttk.Label(bar, text=tr(label)).pack(side="left", padx=(5, 0))
            widget = ttk.Combobox(bar, textvariable=self.fv[key],
                                  width=width, state="readonly")
            widget.pack(side="left", padx=(4, 5), pady=4)
            widget.bind("<<ComboboxSelected>>", lambda _: self.after_idle(self.refresh))
            self.cb[key] = widget

        combo("公司", "company", 10)
        combo("办公地点", "site", 10)
        combo("产权人", "owner", 9)
        combo("存放位置", "location", 10)
        combo("状态", "status", 8)
        ttk.Button(bar, text=tr("查询"), command=self.refresh).pack(side="left", padx=(8, 4))
        ttk.Button(bar, text=tr("重置"), command=self.reset).pack(side="left", padx=(0, 6))
        summary_bar = ttk.Frame(main)
        summary_bar.pack(fill="x", pady=(3, 0))
        self.stat = ttk.Label(summary_bar, text=tr(""), foreground="#34516b")
        self.stat.pack(side="right")

        body = ttk.Frame(main); body.pack(fill="both", expand=True, pady=6)
        cols = LIST_COLUMNS
        self.tree = ttk.Treeview(body, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            head_min = self._heading_font.measure(tr(HEAD[c])) + S(28)
            self.tree.heading(c, text=tr(HEAD[c]))
            self.tree.column(c, width=max(head_min, S(COLUMN_BASE_WIDTHS.get(c, 110))),
                             minwidth=head_min, stretch=False, anchor="w")
        sb = ttk.Scrollbar(body, command=self.tree.yview)
        xsb = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=xsb.set)
        body.grid_rowconfigure(0, weight=1); body.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew"); sb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")  # 列多时可左右滚动
        for tag, (bg, fg) in TAG_COLORS.items():
            self.tree.tag_configure(tag, background=bg, foreground=fg)
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)
        self.tree.bind("<Double-1>", lambda _: self.edit())
        self.tree.bind("<Configure>", self._schedule_column_layout)
        self._schedule_column_layout()
        self.update_mode()
        self._schedule_split_restore()

    # --- 状态 ---
    def update_mode(self):
        """Editing is allowed by default unless the user enabled protection."""
        can_edit = not self.protection_enabled or self.unlocked
        if self.protection_enabled:
            command_text = "锁定(退出编辑)" if self.unlocked else "解锁编辑"
            mode_text = "模式: 可编辑" if self.unlocked else "模式: 只读"
        else:
            command_text = "编辑保护：已关闭"
            mode_text = "模式: 可编辑（无需密码）"
        self.mode_lbl.config(text=tr(mode_text))
        self.lock_btn.config(text=tr(command_text))
        self.lock_menu.entryconfig(0, label=tr(command_text))
        for button in self.edit_btns:
            button.state(["!disabled"] if can_edit else ["disabled"])
        for menu, label in self.edit_items:
            menu.entryconfig(label, state="normal" if can_edit else "disabled")
        # The change-password command has meaning only when protection is on.
        self.lock_menu.entryconfig(tr("修改密码"), state=(
            "normal" if self.protection_enabled and self.unlocked else "disabled"))

    def toggle_lock(self):
        if not self.protection_enabled:
            self.configure_edit_protection()
            return
        if self.unlocked:
            self.unlocked = False
        else:
            entered = simpledialog.askstring(
                tr("解锁"), tr("请输入编辑密码:"), show="*", parent=self)
            if entered is None:
                return
            if not self.db.check_password(entered):
                messagebox.showerror(tr("错误"), tr("密码错误"), parent=self)
                return
            self.unlocked = True
        self.update_mode()

    def need_unlock(self):
        if self.protection_enabled and not self.unlocked:
            messagebox.showinfo(tr("只读模式"),
                                tr("请先解锁编辑后重试"), parent=self)
            return False
        return True

    def configure_edit_protection(self):
        """Password protection is opt-in. Old password hashes remain protected."""
        window = tk.Toplevel(self)
        set_window_icon(window)
        window.title(tr("编辑密码保护"))
        window.transient(self)
        window.resizable(False, False)
        content = ttk.Frame(window, padding=20)
        content.pack(fill="both", expand=True)
        desired = tk.BooleanVar(master=window, value=self.protection_enabled)
        window.protection_choice = desired
        ttk.Checkbutton(content, text=tr("启用编辑密码保护"),
                        variable=desired).pack(anchor="w", pady=(0, 12))
        ttk.Label(content, text=tr("未启用密码保护，可直接新增、编辑、删除设备。"),
                  justify="left", wraplength=S(470)).pack(anchor="w", pady=(0, 9))
        ttk.Label(content, text=tr("启用后可使用密码锁定编辑；仅控制软件内操作，不会加密数据库。"),
                  justify="left", wraplength=S(470), foreground="#5d6878").pack(anchor="w", pady=(0, 9))
        ttk.Label(content, text=tr("关闭密码保护需要验证现有密码。已保存的设备数据不会受到影响。"),
                  justify="left", wraplength=S(470), foreground="#5d6878").pack(anchor="w", pady=(0, 16))
        buttons = ttk.Frame(content)
        buttons.pack(anchor="e", fill="x")

        def save_protection():
            enable = desired.get()
            if enable == self.protection_enabled:
                window.destroy()
                return
            if enable:
                new_pwd = simpledialog.askstring(
                    tr("设置密码"), tr("首次使用，请设置编辑密码:"), show="*", parent=window)
                if not new_pwd:
                    return
                confirmed = simpledialog.askstring(
                    tr("设置密码"), tr("请再次输入密码:"), show="*", parent=window)
                if confirmed is None:
                    return
                if confirmed != new_pwd:
                    messagebox.showerror(tr("错误"), tr("两次输入不一致"), parent=window)
                    return
                self.db.set_password(new_pwd)
                self.db.set_setting(UI_EDIT_PROTECT_KEY, "1")
                self.protection_enabled = True
                self.unlocked = True
                self.db.log("启用编辑密码保护")
                feedback = "编辑密码保护已启用"
            else:
                entered = simpledialog.askstring(
                    tr("编辑密码保护"), tr("输入当前密码以关闭编辑保护："),
                    show="*", parent=window)
                if entered is None:
                    return
                if not self.db.check_password(entered):
                    messagebox.showerror(tr("错误"), tr("密码错误"), parent=window)
                    return
                if not messagebox.askyesno(
                    tr("确认"), tr("确认关闭密码保护？关闭后任何能访问此程序的人都能在软件中修改记录。"),
                    parent=window):
                    return
                self.db.disable_edit_protection()
                self.protection_enabled = False
                self.unlocked = True
                self.db.log("关闭编辑密码保护")
                feedback = "编辑密码保护已关闭"
            self.update_mode()
            window.destroy()
            messagebox.showinfo(tr("完成"), tr(feedback), parent=self)

        window.save_protection = save_protection  # UI test handle; also used by Save button
        ttk.Button(buttons, text=tr("保存设置"),
                   command=save_protection).pack(side="right", padx=(7, 0))
        ttk.Button(buttons, text=tr("取消"),
                   command=window.destroy).pack(side="right")
        window.grab_set()
        return window

    def change_pw(self):
        if not self.protection_enabled:
            self.configure_edit_protection()
            return
        if not self.need_unlock():
            return
        p1 = simpledialog.askstring(tr("修改密码"), tr("新密码:"), show="*", parent=self)
        if not p1:
            return
        if simpledialog.askstring(tr("修改密码"), tr("再次输入新密码:"), show="*", parent=self) != p1:
            messagebox.showerror(tr("错误"), tr("两次输入不一致"), parent=self)
            return
        self.db.set_password(p1)
        self.db.log("修改密码")
        messagebox.showinfo(tr("完成"), tr("密码已修改"), parent=self)

    # --- 列表 ---
    def reset(self):
        self.kw.set("")
        for v in self.fv.values(): v.set("")
        self.refresh()

    def rows(self):
        g = lambda k: self.fv[k].get().strip()
        return self.db.query(self.kw_entry.get().strip(), owner=g("owner"), location=g("location"),
                             status=ui_status_to_db(g("status")), company=g("company"), site=g("site"))

    def cell(self, r, c):
        if c == "warranty": return tr(warranty_info(r, self.cfg["warn_days"])[0])
        if c == "net":
            v = net_value(r, self.cfg); return "" if v is None else f"{v:,.2f}"
        if c == "alert": return "; ".join(tr(t) for _, t in self.active_alerts(r))
        if c in ("qty", "value", "warranty_months"): return fmt_num(r[c]) if r[c] is not None else ""
        if c == "status": return tr(r[c] or "")
        return r[c] or ""

    def refresh(self):
        for key in self.cb:
            vals = [""] + [tr(x) if key == "status" else x for x in self.db.distinct(key)]
            if list(self.cb[key].cget("values")) != vals:
                self.cb[key]["values"] = vals
        self.cfg = self.db.load_cfg()
        keep = self.tree.selection()  # 刷新后保持选中，右侧详情不丢失
        self.tree.delete(*self.tree.get_children())
        rows = self.rows(); tq = tv = tn = 0
        cols = self.tree["columns"]
        for r in rows:
            vals = [self.cell(r, c) for c in cols]
            al = self.active_alerts(r)
            tag = "high" if any(l == "high" for l, _ in al) else (
                warranty_info(r, self.cfg["warn_days"])[2] if self.warranty_alerts_on else ""
            )
            if tag == "soon" and not al: tag = "in"  # 已报废/已处置设备的保修到期不再标黄
            self.tree.insert("", "end", iid=str(r["id"]), values=vals, tags=(tag,))
            q = r["qty"] or 0
            tq += q; tv += (r["value"] or 0) * q
            nv = net_value(r, self.cfg); tn += (nv if nv is not None else (r["value"] or 0)) * q
        summary = (f"Records: {len(rows)} | Quantity: {fmt_num(tq)} | "
                   f"Total Value: {tv:,.2f} | Net Value: {tn:,.2f}") if self.language == "en" else (
                   f"共 {len(rows)} 条 | 数量 {fmt_num(tq)} | 价值合计 {tv:,.2f} | 净值合计 {tn:,.2f}")
        self.stat.config(text=summary)
        n = sum(1 for r in self.db.query() if self.active_alerts(r))
        self.alert_btn.config(text=tr("提醒中心") + (f" ({n})" if n else ""))
        keep = [k for k in keep if self.tree.exists(k)]
        if keep: self.tree.selection_set(keep)
        self.show_detail()
        self._schedule_column_layout()

    def show_detail(self, _=None):
        sel = self.tree.selection()
        self._photo = None
        r = self.db.get(int(sel[0])) if len(sel) == 1 else None
        for f, lbl in self.dv.items():
            lbl.config(text=tr("" if not r else self.cell(r, f)))
        if r and "warranty" in self.dv and warranty_info(r, self.cfg["warn_days"])[1]:
            w = warranty_info(r, self.cfg["warn_days"])
            self.dv["warranty"].config(text=tr(f"{tr(w[0])}, {tr('到期 ')}{w[1]}"))
        # 单据页
        self.att_tree.delete(*self.att_tree.get_children())
        atts = self.db.attachments(r["id"]) if r else []
        for a in atts:
            self.att_tree.insert("", "end", iid=str(a["id"]), values=(a["orig_name"], (a["added"] or "")[:10]))
        self.nb.tab(1, text=tr("单据") + (f" ({len(atts)})" if atts else ""))
        # 变更记录页
        lb = self.log_box; lb.config(state="normal"); lb.delete("1.0", "end")
        logs = self.db.logs_for(r["id"]) if r else []
        for lg in logs:
            lb.insert("end", f"{lg['ts']}  {tr(lg['action'])}  ({lg['operator'] or '-'})\n", "h")
            if lg["detail"]:
                lb.insert("end", "    " + lg["detail"].replace("\n", "\n    ") + "\n")
        if not logs:
            lb.insert("end", "(暂无记录)" if r else "选中单个设备可查看变更记录")
        lb.config(state="disabled")
        # 图片
        if not r:
            self.preview.config(image="", text=tr("选中单个设备可查看详情和图片")); return
        try:
            p = stored_file_path(IMG_DIR, r["image"]) if r["image"] else ""
        except ValueError:
            self.preview.config(image="", text=tr("图片路径不安全"))
            return
        if not p or not os.path.isfile(p):
            self.preview.config(image="", text=tr("(无图片)")); return
        try:
            pil = pillow_modules()
            if pil:
                Image, ImageTk = pil
                im = Image.open(p); im.thumbnail((S(400), S(180))); self._photo = ImageTk.PhotoImage(im)
            else:
                self._photo = tk.PhotoImage(file=p)
            self.preview.config(image=self._photo, text=tr(""))
        except Exception:
            self.preview.config(image="", text=tr("无法预览(安装 Pillow 可支持更多格式)"))

    # --- 采购单据 ---
    def _cur_id(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showinfo(tr("提示"), tr("请先在列表中选中一条设备")); return None
        return int(sel[0])

    def att_add(self):
        if not self.need_unlock(): return
        i = self._cur_id()
        if i is None: return
        files = filedialog.askopenfilenames(title=tr("选择采购单据(可多选)"), parent=self)
        n = 0
        for f in files:
            if os.path.getsize(f) > MAX_ATT_MB * 1024 * 1024:
                messagebox.showwarning(tr("提示"), tr(f"{os.path.basename(f)} 超过 {MAX_ATT_MB}MB，已跳过")); continue
            try:
                self.db.add_attachment(i, f); n += 1
            except OSError as e:
                messagebox.showerror(tr("错误"), tr(f"添加失败: {e}"))
        if n: self.show_detail(); self.nb.select(1)

    def att_open(self):
        sel = self.att_tree.selection()
        if not sel:
            messagebox.showinfo(tr("提示"), tr("请先选中一个单据")); return
        r = self.db.conn.execute("SELECT stored FROM attachments WHERE id=?", (int(sel[0]),)).fetchone()
        try:
            p = stored_file_path(ATT_DIR, r["stored"]) if r else ""
        except ValueError:
            p = ""
        if not os.path.isfile(p):
            messagebox.showerror(tr("错误"), tr("单据文件不存在(可能被移动或删除)")); return
        try:
            open_path(p)
        except Exception as e:
            messagebox.showerror(tr("错误"), tr(f"无法打开: {e}"))

    def att_remove(self):
        if not self.need_unlock(): return
        sel = self.att_tree.selection()
        if sel and messagebox.askyesno(tr("确认"), tr("删除选中的单据？(若其他设备共用同一文件，文件会保留)")):
            for x in sel: self.db.remove_attachment(int(x))
            self.show_detail()

    # --- 预警与折旧设置 ---
    def edit_cfg(self):
        if not self.need_unlock(): return
        win = tk.Toplevel(self); set_window_icon(win); win.title(tr("预警与折旧设置")); win.transient(self); win.resizable(False, False)
        frm = ttk.Frame(win, padding=14); frm.pack()
        vs = {}
        for i, (k, label) in enumerate(CFG_LABELS):
            ttk.Label(frm, text=tr(label)).grid(row=i, column=0, sticky="e", pady=4)
            vs[k] = tk.StringVar(value=fmt_num(self.cfg[k]))
            ttk.Entry(frm, textvariable=vs[k], width=10).grid(row=i, column=1, padx=8)
        ttk.Label(frm, foreground="gray", justify="left",
                  text=tr("折旧: 按购买时间起直线折旧，到期后保留残值。\n报废设备一律提醒“待处理”，处理完改为“已处置”。")
                  ).grid(row=len(CFG_LABELS), column=0, columnspan=2, pady=6)
        def ok():
            try:
                cfg = {k: float(v.get()) for k, v in vs.items()}
                if cfg["life_years"] <= 0 or not 0 <= cfg["residual_pct"] <= 100 or min(cfg["warn_days"], cfg["repair_days"]) < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(tr("提示"), tr("请输入有效数字(年限>0，残值率0-100，天数>=0)"), parent=win); return
            ch = [f"{lb}: {fmt_num(self.cfg[k])} → {fmt_num(cfg[k])}" for k, lb in CFG_LABELS if self.cfg[k] != cfg[k]]
            self.db.save_cfg(cfg)
            if ch: self.db.log("修改预警与折旧设置", "\n".join(ch))
            win.destroy(); self.refresh()
        bb = ttk.Frame(frm); bb.grid(row=len(CFG_LABELS) + 1, column=0, columnspan=2)
        ttk.Button(bb, text=tr("保存"), command=ok).pack(side="left", padx=8)
        ttk.Button(bb, text=tr("取消"), command=win.destroy).pack(side="left", padx=8)

    # --- 提醒中心 ---
    def collect_alerts(self):
        items = []
        for r in self.db.query():
            for lvl, txt in self.active_alerts(r):
                items.append((0 if lvl == "high" else 1, r, txt))
        items.sort(key=lambda x: (x[0], x[2]))
        return items

    def show_alerts(self, startup=False):
        items = self.collect_alerts()
        if startup:
            if not items: return
            hi = sum(1 for i in items if i[0] == 0)
            messagebox.showwarning(tr("设备提醒"), tr(f"有 {len(items)} 条设备提醒(其中紧急 {hi} 条)，\n可在“提醒中心”查看详情。"))
            return
        win = tk.Toplevel(self); set_window_icon(win); win.title(tr("提醒中心")); win.geometry(f"{S(900)}x{S(480)}"); win.transient(self)
        ttk.Label(win, padding=8, font=("", 12, "bold"),
                  text=tr(f"共 {len(items)} 条提醒 (紫色=维修超期/报废待处理，橙黄=保修即将到期)。双击跳转到该设备。")).pack(anchor="w")
        cols = ("t", "name", "code", "owner", "loc")
        t = ttk.Treeview(win, columns=cols, show="headings")
        for c, h, w in (("t", "提醒内容", 220), ("name", "设备名称", 180), ("code", "设备编码", 130),
                        ("owner", "产权人", 100), ("loc", "存放位置", 140)):
            t.heading(c, text=tr(h)); t.column(c, width=S(w))
        t.tag_configure("high", background=TAG_COLORS["high"][0], foreground=TAG_COLORS["high"][1])
        t.tag_configure("warn", background=TAG_COLORS["soon"][0], foreground=TAG_COLORS["soon"][1])
        for i, (lvl, r, txt) in enumerate(items):
            t.insert("", "end", iid=f"{i}:{r['id']}", tags=("high" if lvl == 0 else "warn",),
                     values=(tr(txt), r["name"], r["code"], r["owner"] or "", r["location"] or ""))
        t.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        def jump(_):
            sel = t.selection()
            if not sel: return
            self.reset(); iid = sel[0].split(":")[1]
            if self.tree.exists(iid):
                self.tree.selection_set(iid); self.tree.see(iid)
            win.destroy()
        t.bind("<Double-1>", jump)

    # --- 统计报表 ---
    def show_stats(self):
        rows = self.rows()  # 与当前筛选条件一致
        win = tk.Toplevel(self); set_window_icon(win); win.title(tr("统计报表")); win.geometry(f"{S(760)}x{S(520)}"); win.transient(self)
        total_q = sum(r["qty"] or 0 for r in rows)
        total_v = sum((r["qty"] or 0) * (r["value"] or 0) for r in rows)
        total_n = sum((r["qty"] or 0) * (net_value(r, self.cfg) if net_value(r, self.cfg) is not None else (r["value"] or 0)) for r in rows)
        flt = [f"{tr(n)}={v}" for n, v in (("公司", self.fv["company"].get()), ("办公地点", self.fv["site"].get()),
               ("产权人", self.fv["owner"].get()), ("位置", self.fv["location"].get()),
               ("状态", self.fv["status"].get()), ("关键字", self.kw_entry.get().strip())) if v]
        ttk.Label(win, padding=8, font=("", 12, "bold"),
                  text=tr(f"设备 {len(rows)} 条 | 数量合计 {fmt_num(total_q)} | 资产总价值 {total_v:,.2f} | 折旧后净值 {total_n:,.2f}"
                       + (f"    (筛选: {'，'.join(flt)})" if flt else "    (全部)"))).pack(anchor="w")
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        for key, title in (("company", "按公司"), ("site", "按办公地点"), ("status", "按状态"),
                           ("owner", "按产权人"), ("location", "按存放位置")):
            agg = {}
            for r in rows:
                g = agg.setdefault(r[key] or "(未填写)", [0, 0.0, 0.0])
                g[0] += 1; g[1] += r["qty"] or 0; g[2] += (r["qty"] or 0) * (r["value"] or 0)
            fr = ttk.Frame(nb); nb.add(fr, text=tr(title))
            t = ttk.Treeview(fr, columns=("g", "n", "q", "v", "p"), show="headings")
            for c, h, w in (("g", title[1:], 220), ("n", "设备条数", 90), ("q", "数量合计", 100),
                            ("v", "价值合计", 130), ("p", "价值占比", 90)):
                t.heading(c, text=tr(h)); t.column(c, width=S(w), anchor="w" if c == "g" else "e")
            for g, (n, q, v) in sorted(agg.items(), key=lambda x: -x[1][2]):
                t.insert("", "end", values=(tr(g) if key == "status" else g, n, fmt_num(q), f"{v:,.2f}", f"{v / total_v:.1%}" if total_v else "-"))
            t.pack(fill="both", expand=True)

    # --- 操作日志(全部) ---
    def show_log(self):
        win = tk.Toplevel(self); set_window_icon(win); win.title(tr("操作日志")); win.geometry(f"{S(1100)}x{S(560)}"); win.transient(self)
        top = ttk.Frame(win, padding=8); top.pack(fill="x")
        kw = tk.StringVar(); ttk.Label(top, text=tr("搜索(编码/名称/动作/内容/操作人)")).pack(side="left")
        e = ttk.Entry(top, textvariable=kw, width=26); e.pack(side="left", padx=6)
        info = ttk.Label(top, text=tr("")); info.pack(side="right")
        body = ttk.Frame(win); body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("ts", "who", "act", "code", "name", "detail")
        t = ttk.Treeview(body, columns=cols, show="headings")
        for c, h, w in (("ts", "时间", 150), ("who", "操作人", 90), ("act", "动作", 100), ("code", "设备编码", 120),
                        ("name", "设备名称", 140), ("detail", "内容", 460)):
            t.heading(c, text=tr(h)); t.column(c, width=S(w))
        sb = ttk.Scrollbar(body, command=t.yview); t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True); sb.pack(side="left", fill="y")
        def load(_=None):
            t.delete(*t.get_children())
            rows = self.db.all_logs(kw.get().strip(), 2000)
            for r in rows:
                t.insert("", "end", values=(r["ts"], r["operator"] or "", tr(r["action"]), r["code"] or "",
                                            r["name"] or "", (r["detail"] or "").replace("\n", " ； ")))
            info.config(text=tr(f"显示最近 {len(rows)} 条(最多 2000 条)"))
        e.bind("<KeyRelease>", load); load()

    # --- 增删改 ---
    def add(self):
        if not self.need_unlock(): return
        d = EditDialog(self).result
        if d:
            i = self.try_save(d)
            if i and self.tree.exists(str(i)):  # 选中新记录，方便立刻在右侧添加采购单据
                self.tree.selection_set(str(i)); self.tree.see(str(i))

    def edit(self):
        if not self.need_unlock(): return
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showinfo(tr("提示"), tr("请选中一条记录")); return
        i = int(sel[0]); d = EditDialog(self, self.db.get(i)).result
        if d: self.try_save(d, i)

    def copy_device(self):
        if not self.need_unlock(): return
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showinfo(tr("提示"), tr("请选中一条要复制的设备")); return
        src_id = int(sel[0]); src_code = self.db.get(src_id)["code"]
        d = EditDialog(self, self.db.get(src_id), copy=True).result
        if not d: return
        n = d.pop("copies")
        w = max(3, len(str(n)))
        codes = [d["code"]] if n == 1 else [f"{d['code']}-{i:0{w}d}" for i in range(1, n + 1)]
        dup = [c for c in codes if self.db.find_code(c)]
        if dup:
            messagebox.showerror(tr("错误"), tr(f"设备编码已存在，未复制：\n{'、'.join(dup[:5])}{' 等' if len(dup) > 5 else ''}")); return
        for c in codes:
            new_id = self.db.save(dict(d, code=c), src="复制", note=f"复制自 {src_code}")
            self.db.copy_attachments(src_id, new_id)
        self.refresh()
        messagebox.showinfo(tr("完成"), tr(f"已复制 {n} 条设备"))

    def try_save(self, d, id_=None):
        ex = self.db.find_code(d["code"])
        if ex and ex["id"] != id_:
            messagebox.showerror(tr("错误"), tr(f"设备编码 {d['code']} 已存在(编码不区分大小写，全公司唯一)")); return
        try:
            saved_id = self.db.save(d, id_)
        except (sqlite3.DatabaseError, ValueError, OverflowError) as exc:
            messagebox.showerror(tr("错误"), str(exc), parent=self)
            return None
        self.refresh()
        return saved_id

    def remove(self):
        if not self.need_unlock(): return
        sel = self.tree.selection()
        if sel and messagebox.askyesno(tr("确认"), tr(f"删除选中的 {len(sel)} 条记录？")):
            self.db.delete([int(s) for s in sel]); self.refresh()

    # --- 导入导出 ---
    def export_file(self):
        rows = self.rows()
        types = [("Excel", "*.xlsx"), ("CSV", "*.csv")] if has_excel() else [("CSV", "*.csv")]
        p = filedialog.asksaveasfilename(defaultextension=types[0][1], filetypes=types,
                                         initialfile=f"设备清单_{datetime.now():%Y%m%d_%H%M}")
        if not p: return
        head = [tr(h) for _, h in DISPLAY] + [tr(HEAD["image"])]
        data = [[self.cell(r, f) if f in ("warranty", "net", "alert") else r[f] for f, _ in DISPLAY] + [r["image"]] for r in rows]
        if p.lower().endswith(".xlsx"):
            openpyxl = excel_module()
            if openpyxl is None:
                messagebox.showerror(tr("错误"), tr("导入 xlsx 需要安装 openpyxl: pip install openpyxl"))
                return
            wb = openpyxl.Workbook(); ws = wb.active; ws.title = tr("设备清单")
            ws.append(head); [ws.append(x) for x in data]
            for i, w in enumerate([18, 16, 14, 16, 12, 12, 14, 8, 12, 8, 8, 10, 12, 12, 22, 12, 24, 24, 20], 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
            wb.save(p)
        else:
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f); w.writerow(head); w.writerows(data)
        messagebox.showinfo(tr("完成"), tr(f"已导出 {len(rows)} 条(当前筛选结果)"))

    def import_file(self):
        if not self.need_unlock(): return
        p = filedialog.askopenfilename(filetypes=[(tr("表单"), "*.xlsx *.csv"), (tr("所有文件"), "*.*")])
        if not p: return
        try:
            if p.lower().endswith(".xlsx"):
                openpyxl = excel_module()
                if openpyxl is None:
                    messagebox.showerror(tr("错误"), tr("导入 xlsx 需要安装 openpyxl: pip install openpyxl")); return
                ws = openpyxl.load_workbook(p, read_only=True, data_only=True).active
                table = [list(r) for r in ws.iter_rows(values_only=True)]
            else:
                table = None
                for enc in ("utf-8-sig", "gbk"):
                    try:
                        with open(p, newline="", encoding=enc) as f:
                            table = list(csv.reader(f))
                        break
                    except UnicodeDecodeError:
                        continue
                if table is None:
                    raise UnicodeError("CSV 编码无法识别（支持 UTF-8 和 GBK）")
        except Exception as e:
            messagebox.showerror(tr("错误"), tr(f"读取失败: {e}")); return
        if not table: return
        rev = {h: f for f, h in FIELDS}
        rev.update({EN_TEXT.get(h, h): f for f, h in FIELDS})
        rev.update(OLD_HEAD)
        idx = {rev[str(h).strip()]: i for i, h in enumerate(table[0]) if h and str(h).strip() in rev}
        if "name" not in idx or "code" not in idx:
            messagebox.showerror(tr("错误"), tr("表头需包含“设备名称”和“设备编码”")); return
        add = upd = skip = nodate = 0
        for row in table[1:]:
            g = lambda f: ("" if f not in idx or idx[f] >= len(row) or row[idx[f]] is None else str(row[idx[f]]).strip())
            d = {f: g(f) for f, _ in FIELDS}
            d["status"] = ui_status_to_db(d.get("status", ""))
            if not d["name"] or not d["code"]: skip += 1; continue
            d["qty"] = num(d["qty"], 1); d["value"] = num(d["value"]); d["warranty_months"] = num(d["warranty_months"])
            pd_ = parse_date(d["purchase_date"]); d["purchase_date"] = pd_.isoformat() if pd_ else ""
            if not pd_ and "purchase_date" in idx: nodate += 1
            ex = self.db.find_code(d["code"])
            if ex:
                old = self.db.get(ex["id"])
                for k, _ in FIELDS:  # 导入表没有该列时保留原值(状态日期由 save 按状态是否变化处理)
                    if k not in idx and k != "status_date": d[k] = old[k]
            try:
                self.db.save(d, ex["id"] if ex else None, src="导入")
            except (ValueError, sqlite3.IntegrityError):
                skip += 1
                continue
            if ex:
                upd += 1
            else:
                add += 1
        self.db.log("导入表单", f"{os.path.basename(p)}: 新增 {add} 条，更新 {upd} 条，跳过 {skip} 条")
        self.refresh()
        messagebox.showinfo(tr("导入完成"), tr(f"新增 {add} 条，按编码更新 {upd} 条，跳过 {skip} 条"
                            + (f"\n其中 {nodate} 条购买日期缺失或无法识别，已留空，请补充。" if nodate else "")))

    # --- 备份 ---
    def backup(self):  # 需解锁(输入密码)后才可用
        if not self.need_unlock(): return
        p = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[(tr("备份包"), "*.zip")],
                                         initialfile=f"设备登记备份_{datetime.now():%Y%m%d_%H%M%S}")
        if not p: return
        fd, tmp = tempfile.mkstemp(suffix=".db"); os.close(fd)
        try:
            dst = sqlite3.connect(tmp)
            try:
                self.db.conn.backup(dst)  # 一致性快照
            finally:
                dst.close()
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(tmp, "devices.db")
                for folder, prefix in ((IMG_DIR, "images"), (ATT_DIR, "attachments")):
                    for name in os.listdir(folder):
                        source = os.path.join(folder, name)
                        # Ignore symlinks and subdirectories: backups must not
                        # traverse outside the application-owned data folders.
                        if os.path.isfile(source) and not os.path.islink(source):
                            z.write(source, f"{prefix}/{name}")
        except Exception as e:
            messagebox.showerror(tr("备份失败"), tr(str(e))); return
        finally:
            try: os.remove(tmp)
            except OSError: pass
        messagebox.showinfo(tr("完成"), tr(f"备份已保存:\n{p}"))

    def restore(self):
        if not self.need_unlock():
            return
        path = filedialog.askopenfilename(filetypes=[(tr("备份包"), "*.zip")])
        if not path or not messagebox.askyesno(
            tr("确认"), tr("恢复将覆盖当前所有数据（包括密码），并创建当前数据库备份。继续？")
        ):
            return

        stage = None
        closed = False
        backed_up = False
        written_files = []
        backup_path = DB_PATH + ".bak"
        try:
            # 校验并完整解压到临时目录；确认有效后才覆盖当前数据。
            stage = tempfile.mkdtemp(prefix=".device-registry-restore-", dir=BASE)
            total_size = 0
            with zipfile.ZipFile(path, "r") as archive:
                members = archive.infolist()
                names = set()
                for item in members:
                    filename = item.filename
                    if item.is_dir():
                        if filename not in ("images/", "attachments/"):
                            raise ValueError(f"备份包含非法目录：{filename}")
                        continue
                    parts = filename.split("/")
                    valid = filename == "devices.db" or (
                        len(parts) == 2 and parts[0] in ("images", "attachments")
                        and parts[1] not in ("", ".", "..")
                    )
                    if not valid or chr(92) in filename or filename in names:
                        raise ValueError(f"备份包含不安全或重复的文件名：{filename}")
                    names.add(filename)
                    total_size += item.file_size
                    if item.file_size > 512 * 1024 * 1024 or total_size > 4 * 1024**3:
                        raise ValueError("备份文件体积超出安全限制")
                if "devices.db" not in names:
                    raise ValueError("缺少 devices.db，不是有效的设备备份")
                if archive.testzip() is not None:
                    raise ValueError("ZIP 校验失败，备份内容可能损坏")
                for item in members:
                    if item.is_dir():
                        continue
                    destination = os.path.join(stage, *item.filename.split("/"))
                    os.makedirs(os.path.dirname(destination), exist_ok=True)
                    with archive.open(item) as source, open(destination, "wb") as target:
                        shutil.copyfileobj(source, target)

            staged_db = os.path.join(stage, "devices.db")
            check = sqlite3.connect(staged_db)
            try:
                result = check.execute("PRAGMA quick_check").fetchone()
                if not result or result[0] != "ok":
                    raise ValueError("备份数据库完整性校验失败")
                has_devices = check.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name='devices'"
                ).fetchone()
                if not has_devices:
                    raise ValueError("备份缺少设备数据表")
            finally:
                check.close()

            self.db.close()
            closed = True
            shutil.copy2(DB_PATH, backup_path)
            backed_up = True

            # Keep replaced attachments available for rollback until the DB
            # has been installed. Existing image files may share their names.
            rollback_dir = os.path.join(stage, ".previous_files")
            for category, folder in (("images", IMG_DIR), ("attachments", ATT_DIR)):
                staged_folder = os.path.join(stage, category)
                if not os.path.isdir(staged_folder):
                    continue
                for filename in os.listdir(staged_folder):
                    destination = stored_file_path(folder, filename)
                    original = None
                    if os.path.lexists(destination):
                        if not os.path.isfile(destination) or os.path.islink(destination):
                            raise ValueError("目标附件路径不是普通文件")
                        rollback_parent = os.path.join(rollback_dir, category)
                        os.makedirs(rollback_parent, exist_ok=True)
                        original = os.path.join(rollback_parent, filename)
                        shutil.copy2(destination, original)
                    os.replace(os.path.join(staged_folder, filename), destination)
                    written_files.append((destination, original))
            # Database is swapped last so that its file references already exist.
            os.replace(staged_db, DB_PATH)
        except Exception as exc:
            for destination, previous in reversed(written_files):
                try:
                    if previous:
                        os.replace(previous, destination)
                    else:
                        os.remove(destination)
                except OSError:
                    pass
            if closed and backed_up:
                try:
                    shutil.copy2(backup_path, DB_PATH)
                except OSError:
                    pass
            if closed:
                self.db = DB(DB_PATH)
            messagebox.showerror(tr("恢复失败"), tr(f"备份未能安全恢复：{exc}"))
            return
        finally:
            if stage:
                shutil.rmtree(stage, ignore_errors=True)

        self.db = DB(DB_PATH)
        new_language = self.db._get(UI_LANGUAGE_KEY) or "zh"
        self._load_preferences()
        self.protection_enabled = self.db.edit_protection_enabled()
        self.unlocked = not self.protection_enabled
        if new_language != self.language:
            self.change_language(new_language)
        else:
            self.refresh()
        self.update_mode()
        self.db.log("恢复备份", os.path.basename(path))
        restored_msg = ("恢复完成；备份启用了密码保护，请使用备份的编辑密码解锁。"
                        if self.protection_enabled else
                        "恢复完成；未启用密码保护，可以直接编辑。")
        messagebox.showinfo(tr("完成"),
                            tr(restored_msg) + "\n" + tr("恢复前数据库已另存为 devices.db.bak。"), parent=self)

    def on_close(self):
        for callback in (getattr(self, "_restore_split_job", None),
                         getattr(self, "_column_resize_job", None),
                         getattr(self, "_startup_alert_job", None)):
            if callback is not None:
                try:
                    self.after_cancel(callback)
                except tk.TclError:
                    pass
        self._save_split()
        self.db.close()
        self.destroy()


if __name__ == "__main__":
    enable_dpi()  # 必须先于 Tk 窗口创建
    App().mainloop()
