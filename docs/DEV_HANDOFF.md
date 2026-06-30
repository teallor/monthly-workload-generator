# 开发交接手册

原创作者：Rafael_Huang

## 开发前必读

1. `docs/PROJECT_MEMORY.md`
2. `docs/PROJECT_STATUS.md`
3. `docs/ACCEPTANCE_REPORT.md`
4. `docs/ACCEPTANCE_TEST.md`
5. `docs/RUNBOOK.md`
6. `docs/TROUBLESHOOTING.md`
7. `docs/CHANGELOG.md`
8. `docs/RELEASE_HISTORY.md`

## 入口与职责

- `gui_app.py`：Tkinter/ttk GUI。
- `main.py`：CLI 参数与流程。
- `workload_service.py`：preview/write 服务层。
- `workload_writer.py`：Excel 模板定位与写入。
- `rules.py`：分类、教师匹配和汇总。
- `parsers/`：各格式解析与 OCR。
- `workspace_manager.py`：GUI 工作区和设置。
- `office_bridge.ps1`：Windows Office COM。

## 本地运行与打包

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python gui_app.py
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

构建产物为 `dist/月度工作量表自动生成器.exe`。PyInstaller 使用 `gui_app.spec` 收集 OCR 模型、ONNX Runtime、OpenCV、NumPy 和 Pillow。

## 常见坑

- 旧 `.doc/.xls` 与严格保留 `.xls` 格式依赖 Microsoft Office COM。
- onefile EXE 较大且首次启动较慢。
- Tkinter 控件状态必须在主线程更新。
- OCR 修改不能引入固定月份或固定样例偏置。
- 同名最终 Excel 必须避让，禁止静默覆盖。
- `workspace/`、`output/`、真实 Office/图片材料和 `app_settings.json` 不得提交。

## 分支和发布

- 新功能：`feature/功能名`
- 修复：`fix/问题名`
- 修改前保留旧 EXE；修改后执行验收清单。
- 验收通过后合并 `main`，打语义化标签，创建 GitHub Release。
- EXE/ZIP 只作为 Release 附件；普通 Git 仅提交源码、文档和脱敏清单。

## 回退

不要移动历史标签。发生严重问题时停止发布，保留问题构建，使用 `release/backup/` 中旧版独立运行，并在 `docs/RELEASE_HISTORY.md` 记录原因。
