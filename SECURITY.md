# 安全说明 / Security

## 数据保护范围

Device Registry 把设备、操作日志和密码哈希保存在本地 SQLite。新数据库默认无需编辑密码；如选择启用，应用会用 PBKDF2-SHA256 哈希验证密码，并兼容旧版密码格式。

**编辑密码只限制通过图形界面修改数据，绝不等于数据库加密。** 数据库、照片、合同附件、CSV/XLSX 和 ZIP 备份默认均为明文文件。管理敏感资料时请使用磁盘加密、文件权限与可靠的离线备份。

ZIP 恢复时会限制条目路径、检查压缩包和数据库完整性，并在恢复失败时尝试回滚数据库及已覆盖附件。但任何备份恢复都有系统中断风险，切勿对唯一一份重要数据直接恢复。

## 漏洞报告

建议维护者在 GitHub 仓库启用 **Private vulnerability reporting**，通过私密渠道接受安全报告；不要在公开 Issue 中直接发布可利用的敏感细节。尚未设置私密联系方式时，可通过公开 Issues 报告不涉及漏洞利用细节的一般缺陷。

This app stores data locally and does not encrypt database files or backups. Prefer GitHub private vulnerability reporting for sensitive security issues.
