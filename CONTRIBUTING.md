# 参与开发 / Contributing

欢迎修复问题、改进翻译、完善测试和使用文档。提交之前建议先开 GitHub Issue 说明问题，避免与已有修改冲突。

## 开发环境

Python 3.9+，Tkinter，pytest。核心逻辑只依赖标准库，`requirements.txt` 中的 Excel / 图片包是可选依赖。

```bash
python -m pip install -r requirements-dev.txt
xvfb-run -a python -m pytest -q tests/      # Linux GUI + core tests
python -m pytest -q tests/test_core.py       # headless core tests
python -m compileall -q device_registry.py
```

## 提交约定

- 优先保持现有单文件运行入口和 SQLite 数据格式兼容。
- 用户界面文案需要同步更新中文/英文；新增业务状态须保留历史数据库规范。
- 修改资产列布局请先讨论；1.3.0 起固定 14 列，不再提供自定义显隐。
- 修复缺陷时请附最小重现方法和自动化测试。
- 请勿提交真实设备、公司数据、附件、数据库、隐私信息、密钥或个人收款信息。
- 保持变量名清楚、异常处理聚焦，避免不必要的抽象和大段模板说明。

Contributions are welcome. Please preserve the local SQLite format, update both languages, add regression tests, and never include real asset records or private attachments.
