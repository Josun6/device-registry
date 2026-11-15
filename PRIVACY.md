# 数据与隐私 / Privacy

Device Registry 是本地离线工具，不依赖账户或云端同步，也不包含内置遥测上传。设备登记、用户偏好、日志、编辑密码哈希保存在本地 `devices.db`，图片和单据分别保存在 `images/` 与 `attachments/`。

本地编辑密码不是文件加密；SQLite、附件、导出表单和 ZIP 备份应按敏感业务数据处理。请仅在有写权限的个人目录使用，并管理好文件权限与离线备份。

公开 GitHub 仓库不得包含真实业务数据、含有设备信息的截图或个人收款密钥。仓库演示截图只使用虚构数据。

No cloud service or telemetry is built in. Asset data and backup files remain local but are not encrypted by the editing password.
