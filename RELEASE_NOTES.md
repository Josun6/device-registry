# Device Registry v1.3.0 — GitHub Release 说明

本版将资产清单统一为固定 14 列，并移除了此前的字段显隐勾选配置。新增/编辑窗口也改为固定输入项；设备的净值仍由折旧规则计算，旧数据库里的公司和状态历史不会丢失。

同时修复了窗口尺寸变化时分隔条比例恢复不稳定、状态修改后日期可能保持旧值的问题。备份恢复增加附件同步回滚逻辑，并继续保留中英文菜单、可选编辑密码、图片附件与报表。

**Windows 仅发布标准版** `DeviceRegistry.exe`，包含 Excel 和更多图片格式支持；不提供轻量版。请先在真实 Windows 10/11 上构建和验收，再正式发布 Release。

**升级提示**：先备份 `devices.db`、`images/`、`attachments/`，用数据副本验证新版本。程序内的编辑密码不是数据库加密。详细说明见 `README.md`、`README_EN.md` 和 `docs/USER_GUIDE.md`。

**验证情况**：Linux 虚拟显示环境 42 项自动化测试通过；历史 SQLite 数据库副本验证通过。Windows 多显示器、高 DPI 与最终 EXE 尚需实机验收。
