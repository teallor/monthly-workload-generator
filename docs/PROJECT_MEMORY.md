# 项目记忆

原创作者：Rafael_Huang。本项目由 Rafael_Huang 原创设计并主导完成，CODEX 辅助代码实现、文档整理与工程化归档。

## 项目状态

- 项目名称：月度工作量表自动生成器
- 封版日期：2026-06-29
- 当前正式版本：v1.0 正式验收版（94/100，已完成朋友端到端测试）
- 正式入口：`月度工作量表自动生成器.exe`
- 主要用户：不熟悉 Python 和命令行的 Windows 办公用户

## 程序用途

程序从 Word、PDF、Excel 和图片课表中识别指定月份、指定教师的课程，先生成可核对的 preview，再以原工作量表为模板生成月度 Excel。课程区分培训与考核，考核再区分理论、实训并按班级汇总。

## 入口与架构

- `gui_app.py`：中文 GUI 主入口；正式用户使用。正式发布入口为打包后的 EXE。
- `main.py`：CLI 主入口；维护、批处理和诊断使用。
- `workload_service.py`：GUI 共用的预览、写入服务。
- `workload_writer.py`：模板定位、版式定位和 Excel 写入。
- `office_bridge.ps1`：通过 Microsoft Office COM 读取旧 `.doc/.xls` 并保持 Excel 格式。
- `preview.py`：preview CSV、排除 CSV、日志和解析快照。
- `rules.py`：教师匹配、培训/考核分类与考核汇总规则。
- `workspace_manager.py`：GUI 工作区、材料复制和设置。
- `parsers/`：Word、PDF、Excel、图片和 OCR 解析器。
- `gui_app.spec`、`build_exe.ps1`：PyInstaller onefile 打包入口。

## 核心规则

- “黄”“黄/王”“黄、王”均可命中“黄佳豪”。
- 等级工、司机、技能竞赛、技能等级认定、应知、应会等归入考核。
- 总复习：考核-理论，一天 8 课时。
- 实训：考核-实训，一天 8 课时；依据明确时优先使用依据课时。
- 45 分钟为 1 课时，90 分钟为 2 课时；全天为 8 课时。
- 必须先 preview；同月份快照存在且无待确认项后才能 write。
- 同名最终 Excel 自动使用 `_1`、`_2` 等新文件名，禁止覆盖已有结果。

## 输入、输出与数据边界

- 模板：`.xls`、`.xlsx`。
- 依据：`.doc`、`.docx`、`.pdf`、`.xls`、`.xlsx`、`.jpg`、`.jpeg`、`.png`。
- GUI 数据位于 EXE 同目录的 `workspace/`；CLI 默认输出位于 `output/<YYYY-MM>/`。
- 真实模板、课表、照片、output、workspace 和 `app_settings.json` 不应提交 GitHub。

## 2026 年 7 月基准

- 启用 OCR：8 条；培训 3、考核 5。
- 考核汇总：2602 期实训 24、2601 期理论 8、2602 期理论 8。
- 最终 Excel 已由用户人工核对通过。

## 已知限制

- 旧 `.doc/.xls` 和严格保留格式依赖 Windows Microsoft Word/Excel；WPS 不等价。
- OCR 受模糊、反光、倾斜、遮挡和裁切影响，必须核对 preview。
- CLI 缺目录、错模板、输出占用会被阻止，但显示技术堆栈；重定向中文可能乱码。CLI仅供开发者使用，该问题不影响GUI/EXE正式用户入口。
- 朋友已完成EXE启动、说明理解、preview、Excel生成、Excel打开及同名文件避让的端到端测试，未发现阻断性问题。
- 暂无独立自动化测试套件。验收证据位于 `tmp/acceptance_2026-06-29_final_152933/records/`。

## 目录现状

根目录混有真实输入、历史输出、OCR 调试、构建缓存和多轮验收目录。它们按要求未删除。继续开发时不要把这些当源码提交；以 `.gitignore` 为边界。

## 下一次开发前必读

1. `docs/PROJECT_MEMORY.md`
2. `docs/ACCEPTANCE_REPORT.md`
3. `docs/RUNBOOK.md`
4. `docs/TROUBLESHOOTING.md`
5. `docs/CHANGELOG.md`
