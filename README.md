# Device Registry

**免费、离线的设备资产管理工具。** 基于 Python、Tkinter 和 SQLite，适合个人、小型企业、学校及团队。无需账户、服务器或云端服务，支持简体中文和 English。

[English README](README_EN.md) · [使用指南](docs/USER_GUIDE.md) · [版本更新](CHANGELOG.md) · [自愿赞助](DONATE.md)

![Device Registry 图标](assets/icon.png)

![Device Registry v1.3.0 界面预览（演示数据）](docs/screenshot.png)

> 当前版本：**1.3.0**，采用 [MIT License](LICENSE)。版权署名：**sunly56**。软件免费使用，打赏完全自愿。请确认你拥有发布源码、图标及其他素材的权利。

## 功能

- **固定清单布局**：默认展示设备名称、设备编码、型号、办公地点、产权人、存放位置、状态、数量、单位、单价(元)、购买日期、保修期(月)、当前净值(元)、备注，共 14 列。不再提供字段显示勾选功能。
- **简洁录入**：表单直接展示对应的可编辑字段，另提供设备图片；当前净值按折旧规则自动计算，不能手工修改。
- **设备管理**：新增、编辑、复制、删除；按设备编码（不区分大小写）查重，支持批量复制。
- **查找与筛选**：搜索栏保持单行，支持公司、地点、产权人、存放位置、状态等筛选。窗口较窄时可横向滚动。
- **可拖动布局**：资产清单与右侧详情之间有鼠标可拖动的分隔条，调整后的比例自动记忆。表头使用 10pt 字体并按空间计算列宽；列多时通过横向滚动查看。
- **双语界面**：在「语言 / Language」菜单即时切换中文和英文；保存语言偏好，不改动数据库中的中文业务状态。
- **可选密码保护**：新用户默认可以直接编辑；用户可自行启用编辑密码。已有编辑密码的旧库仍受保护。**密码不加密数据库。**
- **业务功能**：采购单据与设备照片、操作日志、折旧和保修提醒、统计报表、Excel/CSV 导入导出、ZIP 备份与恢复。

### 关于旧数据库字段

旧版「所属公司」「状态日期」以及之前保存的字段显示偏好，不会被删除。新版只是固定清单与表单的**显示方式**。原有公司会保留；新设备可在「设置 → 默认公司与提醒」指定默认公司，或在仅有一个公司时自动沿用。

CSV/Excel 导出仍保留完整设备字段，方便交换数据与迁移；因此导出文件的列可能多于主界面的 14 列。

## 运行源码

需要 **Python 3.9+** 和 Tkinter；Windows 官方 Python 安装通常包含 Tkinter。

```bash
python device_registry.py
```

Windows 可双击 `run.bat`。核心功能只依赖 Python 标准库；如需 Excel `.xlsx` 及更多图片格式，可安装可选依赖：

```bash
python -m pip install -r requirements.txt
python device_registry.py
```

程序会在其运行目录创建 `devices.db`、`images/`、`attachments/`。建议放在个人有写入权限的文件夹，**不要直接放进 Windows 的 Program Files 目录**。

## 首次使用

1. 运行程序，确认主界面顶部可见「新增设备」。新数据库默认无需设置密码。
2. 填写设备名称与设备编码，其他字段按需要填写。购买日期默认当天，数量默认 1。
3. 在右侧查看详情、图片；保存设备后可切换到「单据」页上传采购合同或发票。
4. 拖动清单和详情间的蓝色分隔条调整宽度。列表底部横向滚动条用于查看右侧尚未显示的列。
5. 可选：在「设置 → 编辑保护设置」开启密码；在「设置 → 默认公司与提醒」设置默认公司和提醒开关。
6. 定期通过「文件 → 导出备份」保存 ZIP，并在其他位置存放备份副本。

详细步骤见 [用户指南](docs/USER_GUIDE.md)。

## Windows EXE 构建

项目仅维护和发布 **Windows 标准版**，使用 `build_exe.bat` 构建 `DeviceRegistry.exe`，包含 Excel 与 Pillow 图片扩展支持。不提供轻量版。

构建脚本使用隔离的虚拟环境，并包含 `assets/icon.ico`（EXE、标题栏）和 `assets/icon.png`（窗口）。实际 EXE 体积与启动时间取决于 Windows、Python 和 PyInstaller 版本，本环境**未在 Windows 上实测**。

GitHub Actions 仅构建及发布 `DeviceRegistry.exe`。具体操作见 [构建与发布](docs/BUILD_AND_RELEASE.md)。

## 测试与兼容性

```bash
python -m pip install -r requirements-dev.txt
# Linux 桌面环境或 CI：
xvfb-run -a python -m pytest -q tests/
# 无图形桌面的核心测试（也适合 Windows CI）：
python -m pytest -q tests/test_core.py
```

本版已在 Linux 虚拟显示环境中通过 **42 项自动化测试**，覆盖固定列布局、双语、编辑保护、历史 SQLite 库、附件引用、文件路径与备份回滚。并在此前提供的历史数据库**副本**上验证了完整性及记录保留。Windows 10/11 实机、不同 DPI 缩放、多显示器以及最终 EXE 仍须手工验收；不能据此保证所有环境零缺陷。

## 数据安全

- 真实设备记录、操作日志和密码哈希保存在 `devices.db`；照片和附件分别位于 `images/`、`attachments/`。
- 程序不依赖在线服务，也没有内置遥测上传；但用户自行运行的系统或软件环境可能有其他网络行为。
- **不要把真实数据库、备份、照片、发票或合同上传到公开仓库。** `.gitignore` 可帮助避免误提交，但发布前仍须核对 `git status`。
- 编辑密码只限制软件里的编辑操作，并不加密 SQLite、附件或 ZIP 备份。处理敏感资料时请配合磁盘加密与访问权限。

详见 [PRIVACY.md](PRIVACY.md) 和 [SECURITY.md](SECURITY.md)。

## GitHub 发布与反馈

仓库已附带自动测试、Windows 打包工作流、问题反馈模板、开发贡献说明和发布检查清单。请先按 [发布检查清单](docs/PUBLISH_CHECKLIST.md) 核对**许可证、真实数据和 Windows 实机测试**，再创建公开仓库。

欢迎通过 GitHub Issues 反馈问题，详见 [SUPPORT.md](SUPPORT.md)。

**隐私提醒**：如果你自己维护仓库，还应在 GitHub 的 `Settings → Emails` 打开邮箱隐私设置，避免 Git 提交记录公开个人邮箱。MIT [LICENSE](LICENSE) 只保留版权署名，不公开开发者个人邮箱。问题反馈请通过 [GitHub Issues（见 SUPPORT.md）](SUPPORT.md)。项目免费使用，赞助入口见 [DONATE.md](DONATE.md)；当前仓库已包含微信赞赏二维码。
