# 构建、测试与 GitHub 发布

## 从源码运行

Windows 安装 Python 3.9 或更高版本并勾选 Tkinter 支持，然后执行：

```cmd
py -3 device_registry.py
```

更多 Excel/图片支持：`py -3 -m pip install -r requirements.txt`。

## 自动化测试

Linux / GitHub Actions 的 UI 测试：

```bash
python -m pip install -r requirements-dev.txt
xvfb-run -a python -m pytest -q tests/
```

Windows 无显示界面的核心数据层测试：

```cmd
py -3 -m pip install -r requirements-dev.txt
py -3 -m pytest -q tests/test_core.py
```

本地 Linux 虚拟显示环境通过 42 项测试；Windows EXE 尚未在本环境构建或验证。

## 打包 Windows EXE

- 运行 `build_exe.bat`：唯一的 Windows 标准版，支持 `openpyxl` 和 `Pillow`；输出 `dist/DeviceRegistry.exe`。
- 打包脚本使用隔离虚拟环境和 PyInstaller `--onefile`。ICO 图标用于 EXE，PNG/ICO 内嵌用于窗口。
- 不构建、上传或发布 Lite 版本。
- 实际文件体积由 Python、Tcl/Tk、PyInstaller 和可选依赖决定；不要在没有实测前宣传某个固定的缩减比例。
- `--onefile` 启动时需要解压。如果更看重启动速度，可自行将脚本切换到 `--onedir`，但必须一同发布整个目录。

## GitHub 仓库准备

1. 确认你有权公开分发源码、图标和依赖；本仓库使用 [MIT License](../LICENSE)。
2. 许可证署名为 `Copyright (c) 2026 sunly56`；不在任何公开文件中提供个人邮箱，用户通过 GitHub Issues 反馈问题。
3. 如需开启微信赞赏，按 `DONATE.md` 提示放入经过验证的微信赞赏二维码 `assets/wechat-donate.png`；不上传支付密钥、私人账号资料。
4. 检查 `.gitignore` 和 `git status`，不要上传真实 `devices.db`、图片、发票、合同、备份。
5. 在 Windows 10/11 的数据副本上实测新增、编辑、复制、导入、导出、备份恢复，以及高 DPI、多显示器和图标。
6. 将**解压后的项目文件**添加到 GitHub 仓库，不是把 ZIP 当成仓库内容上传。
7. 为防止提交记录暴露私人邮箱，上传前打开 GitHub `Settings → Emails`，启用 `Keep my email addresses private`；命令行提交请将 `git config user.email` 设置为 GitHub 提供的 noreply 地址，并可启用阻止泄露私人邮箱的推送设置。

示例命令（在项目文件夹中，且完成上述检查之后）：

```bash
git init
git add .
git status
git commit -m "Release v1.3.0"
git branch -M main
# 在 GitHub 建立空仓库后，添加该仓库实际的远程 URL：
git remote add origin <你的仓库地址>
git push -u origin main
```

## 发布 Release

Windows 构建工作流支持手动触发，以及推送 `v*` 标签时仅生成 `DeviceRegistry.exe` 并发布 GitHub Release。建议先手动触发构建、下载产物并在干净的 Windows 电脑上试运行，确认无误后再推送正式标签。

```bash
git tag v1.3.0
git push origin v1.3.0
```

Tag 发布前再次确认 LICENSE、打赏信息与最终 EXE 的 Windows 验收。图标缓存会影响任务栏显示效果，建议在新系统或新文件名下验收。
