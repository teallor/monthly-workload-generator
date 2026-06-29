# 月度工作量表自动生成器（GUI + CLI）

当前正式版本：**v1.0 正式验收版**。正式用户入口为中文GUI/EXE；朋友端到端测试已通过。v1.0已经冻结，后续新需求进入v1.1。

## 原创作者

Rafael_Huang

## 开发方式

本项目由 Rafael_Huang 原创设计并主导完成，包括需求定义、业务流程设计、测试验收、版本封版与发布。CODEX 辅助完成代码实现、文档整理与工程化归档。本项目属于 AI-assisted development，原创作者不是 CODEX 或 OpenAI。

## 当前版本

v1.0 正式验收版，已完成正式验收及朋友端到端测试。

本程序用于生成任意年份、任意月份的个人工作量表。它扫描课程通知、课程表、PDF、Excel 和图片依据，只提取目标教师在目标月份内的课程，区分培训与考核，并基于原 Excel 模板生成新文件。

程序提供两种使用方式：

- **中文图形界面（推荐）**：适合不熟悉命令行的 Windows 用户，点选文件即可完成预览和生成。
- **命令行 CLI**：适合批处理、自动化和技术人员。

GUI 和 CLI 共用同一套解析、分类、preview 校验和 Excel 写入逻辑，不是两套程序。

## 图形界面快速开始

### 直接运行 Python 版

```powershell
python gui_app.py
```

也可以通过原 CLI 入口启动：

```powershell
python main.py --gui
```

界面只需要四步：

1. 点击“上传工作量表模板”，选择 `.xls` 或 `.xlsx`。
2. 点击“上传课表/通知材料”多选文件，或点击“导入材料文件夹”。
3. 选择年份、月份，确认教师姓名和别名。
4. 点击“生成预览”，检查待写入、排除课程、考核汇总和解析日志；确认后点击“生成最终 Excel”。

材料列表会显示文件名、类型、大小、解析状态和本地路径。需要确认的课程用醒目颜色显示，确认完成前会阻止生成 Excel。生成成功后可直接打开 Excel 或输出文件夹。

如果照片无法自动识别，可点击“手工补录课程”。手工补录记录仅保留在本次界面会话，不写入最近路径配置。

模板文件名能识别出年月时，GUI 会建议生成模板的下一个月，也可以随时点击“按模板生成下月”。

耗时的扫描和写入会在后台线程运行，界面不会因解析 Word、PDF 或图片而长时间假死。

## Windows EXE

仓库提供 PyInstaller 单文件打包配置。首次打包时运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

脚本会安装 `requirements-build.txt` 中的依赖，并生成：

```text
dist/月度工作量表自动生成器.exe
```

双击该 EXE 即可启动。程序会自动在 EXE 所在目录旁建立工作区，无需用户手工整理：

```text
workspace/
  templates/             GUI 导入的模板副本
  materials/2026-08/     GUI 导入的本月课表材料
  sessions/2026-08/      每次解析使用的隔离副本
  output/2026-08/        preview 和最终 Excel
```

原始文件不会被覆盖；同名文件会自动变成 `课表_1.pdf`、`课表_2.pdf`。`app_settings.json` 只保存最近使用的路径和界面设置，不保存文件内容。`workspace/` 和 `app_settings.json` 已加入 `.gitignore`，不得提交真实业务材料。

旧版 `.doc`、`.xls` 的读取和最终 Excel 写入仍依赖本机安装 Microsoft Word/Excel。图片 OCR 使用本地 RapidOCR + ONNX Runtime，源码环境安装 `requirements.txt` 后即可使用，不调用付费云 API。

如果 Windows 安全提示拦截自制 EXE，可选择“更多信息 → 仍要运行”，或继续使用 `python gui_app.py`。

## 支持格式

- 工作量表模板：`.xls`、`.xlsx`
- 课程依据：`.doc`、`.docx`、`.pdf`、`.xls`、`.xlsx`
- 图片课表：`.jpg`、`.jpeg`、`.png`（可选 OCR，默认关闭）

旧版 `.doc`、`.xls` 及最终工作量表写入依赖本机 Microsoft Word/Excel。程序优先使用 Windows Office COM，以保留模板格式。

## 第一次使用

1. 安装 Microsoft Excel、Word 和 Python。
2. 在程序目录打开 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

OCR 相关依赖也可以单独安装：

```powershell
pip install rapidocr-onnxruntime opencv-python pillow numpy
```

## GUI 如何导入材料

GUI 用户不需要自己放置或整理目录：

- “上传工作量表模板”：选择一个 `.xls/.xlsx`，程序复制到 `workspace/templates/` 并检查模板结构。
- “上传课表/通知材料”：一次选择多个 Word、PDF、Excel 或图片文件。
- “导入材料文件夹”：递归导入整个文件夹；不支持的文件会在列表中标为“不支持”。
- 删除或清空材料列表只影响当前列表，不会删除原始文件。
- 下次启动会恢复上次模板、材料列表、输出目录、年月、教师、别名和 OCR 开关；路径失效时会明确提示重新选择。

## CLI 文件如何放置

最简单的方式是把以下内容放在程序目录：

- 一个文件名含“工作量表”和“黄佳豪”的 `.xls` 或 `.xlsx` 模板。
- 任意数量的 Word、PDF、Excel 或图片课表依据。
- `main.py`、`config.json` 及其他程序文件。

如果依据文件放在单独目录，使用 `--input-dir`；如果目录内有多个工作量表模板，必须使用 `--template` 明确指定。被选中的模板不会被当成课程依据重复解析。

推荐在本地使用以下目录结构（`materials/` 已被 Git 忽略，不会上传真实业务材料）：

```text
monthly-workload-generator/
  materials/
    2026-08/
      培训资源部2026年7月工作量表（黄佳豪）.xls
      课程通知.docx
      课程安排.pdf
      照片课表.jpg
  output/
  main.py
  config.json
```

对应命令：

```powershell
python main.py --input-dir ".\materials\2026-08" `
  --target-month 2026-08 --preview --output-dir output
```

以上目录放置方式只针对 CLI；GUI 会自动管理 `workspace/`。

## 日常操作

1. 把工作量表模板放入文件夹。
2. 把任意课程通知、课表、PDF 或图片放入同一文件夹；也可以单独建立依据目录。
3. 先运行 preview。
4. 检查 `<output-dir>/<目标月份>/` 中的 `待写入预览表.csv`、`被排除课程.csv` 和 `解析日志.txt`。
5. 确认无误后，另行运行 write。

### Preview 示例

```powershell
python main.py --target-month 2026-07 --preview
python main.py --target-month 2026-08 --preview
python main.py --target-month 2026-09 --preview
python main.py --target-month "2026年8月" --preview
python main.py --target-month "8月" --preview
python main.py --target-month 2026-07 --preview --enable-ocr --output-dir output_ocr
python main.py --target-month 2026-07 --preview --ocr --output-dir output_ocr
```

`--enable-ocr` 与 `--ocr` 等价。未传开关时，JPG/PNG 只导入但不自动识别；启用后会尝试 0°、90°、180°、270° 四个方向，并依据标题、月份、实训、总复习、教师等关键词及版面位置选择最佳方向。

GUI 可通过“工具 → OCR 自检”或设置区“检测 OCR”检查 RapidOCR、ONNX Runtime、OpenCV、NumPy、Pillow、模型文件及实际模型会话。正式 onefile EXE 已内置这些组件，不需要朋友另装 Python 或 Tesseract。

只输入 `8月` 时，程序优先从已选模板文件名推断年份；模板不含年份时使用当前年份，并在解析日志中明确记录推断依据。

### Write 示例

```powershell
python main.py --target-month 2026-07 --write
python main.py --target-month 2026-08 --write
python main.py --target-month 2026-09 --write
```

Write 不会覆盖模板。输出文件名和表头年份月份都会根据 `--target-month` 自动更新。
如果目标目录中已经存在同名最终 Excel，程序不会覆盖，而会自动生成 `_1`、`_2` 等新文件名，并在完成提示和日志中显示实际保存路径。
Write 必须先找到同一月份目录中的 preview 和解析快照，并会在命令行打印目标月份、模板、输出文件、课程数量与清单。只有输入 `y` 才会生成 Excel。

### 指定模板及依据目录

```powershell
python main.py --template "培训资源部2026年6月工作量表（黄佳豪）.xls" `
  --input-dir ".\课表依据" --target-month 2026-08 --preview
```

没有传 `--template` 时，程序在输入目录顶层寻找文件名同时包含配置中的 `template_keyword` 和 `teacher_name` 的 `.xls/.xlsx` 文件。找到多个时会列出文件并停止，必须使用 `--template` 选择。

## 配置文件

`config.json` 可修改：

```json
{
  "teacher_name": "黄佳豪",
  "teacher_aliases": ["黄"],
  "target_month": "",
  "template_keyword": "工作量表",
  "input_dir": ".",
  "output_dir": "output",
  "enable_ocr": false,
  "require_confirm_before_write": true
}
```

命令行参数优先于配置文件。`target_month` 可以留空并在运行时传入；`teacher_aliases` 用于配置姓氏简称。建议保留 `require_confirm_before_write: true`。

## Preview 文件说明

每个月份使用独立子目录。例如 `--output-dir output --target-month 2026-08` 会写入 `output/2026-08/`，不会覆盖其他月份。最终 Excel 也位于同一月份目录：`output/2026-08/培训资源部2026年8月工作量表（黄佳豪）.xls`。

- `待写入预览表.csv`：目标月份、目标教师且规则明确的课程。
- `被排除课程.csv`：非目标月份参考课程、其他教师课程、日期不明或低置信度课程，并注明状态和原因。
- `解析日志.txt`：模板定位、月份推断、扫描文件、数量和警告。
- `解析结果.json`：供同一月份的 `--write` 做一致性校验，不建议手工修改。

## 课时与分类

- 文件明确给出课时则优先使用。
- 否则45分钟为1课时；90分钟为2课时；仅写全天或没有具体时长时为8课时。
- 等级工、司机、技能竞赛、技能等级认定、职业技能等级、高级工、中级工、初级工、技师、高级技师、应知、应会归为考核。
- 理论、应知、理论知识、知识辅导归为考核-理论。
- 实训、实操、操作、模拟器、排故、维护保养、应会归为考核-实训。
- 无明确考核属性时默认培训。
- 教师字段支持全名、姓氏简称和多人分隔写法，例如 `黄`、`黄、王`、`黄/王` 都能命中配置教师“黄佳豪”，并在 preview 中标明命中方式。
- 考核项目中的“总复习”归为考核-理论，一天按8课时。
- 实训一天按考核-实训8课时；文件明确给出课时或具体时间段时，以文件信息为准。
- 如果无法判断实训是一整天还是半天，保留在 preview 中并标记“需确认”，确认前禁止 write。

## 图片 OCR 与调试文件

图片 OCR 使用 RapidOCR 中文模型，先做方向检测，再对最佳方向尝试灰度化、放大、对比度增强、去噪和自适应二值化。月历式课表会按左右表格、月份区域、日期单元格、课程与相邻教师坐标解析。

OCR 方向判断使用界面或命令行指定的目标月份进行评分，不固定为某个月份。启用 OCR 后会在 `<output-dir>/ocr_debug/` 生成：

- `<图片名>_最佳方向.png`
- `<图片名>_ocr文本.txt`
- `<图片名>_ocr结果.json`

OCR 原始文本和调试路径也会写入“解析日志”。低置信度候选仍显示在 preview 中并标记为“需确认”。OCR 失败不会中断 Word、PDF、Excel 的解析，可改用 GUI 的“手工补录课程”。历史 `.records.json` 侧车不再冒充 OCR 成功结果。

## 必须人工确认的情况

- 自动发现多个模板。
- 只输入月份且无法形成可靠年份推断。
- 日期、教师、培训/考核或理论/实训无法可靠判断。
- 图片 OCR 候选置信度较低。
- 教学模块行数不足。程序会停止并显示所需行数和模板容量，未经确认不会扩展。
- 模板表头、月份单元格或教学区域无法稳定定位。

## 常见问题

- **找不到模板**：确认模板是 `.xls/.xlsx`，且文件名包含配置中的 `template_keyword` 和 `teacher_name`；也可以传 `--template`。
- **找到多个模板**：程序不会猜测，请使用 `--template "文件名.xls"`。
- **无法解析 `.doc` 或写入 `.xls`**：确认 Windows 已安装 Microsoft Word 和 Excel，并关闭正在编辑的目标文件。
- **图片没有识别**：默认 `enable_ocr=false`。GUI 勾选“启用图片 OCR”，或 CLI 添加 `--enable-ocr`。照片模糊、反光、严重透视或表格裁切时可使用手工补录。
- **教师简称未命中**：检查 `config.json` 中的 `teacher_aliases`。`黄`、`黄/王`、`黄、王` 均可命中黄佳豪。
- **write 被阻止**：先运行同月份 preview；日期、分类、课时或图片 OCR 仍待确认时，程序不会写 Excel。
- **教学槽位不足**：程序会报告培训行数或考核汇总槽位不足，未经确认不会扩展模板。
- **preview 被覆盖**：不会。每个月份分别写入 `<output-dir>/<YYYY-MM>/`。
- **GUI 的“生成最终 Excel”是灰色**：必须先成功生成当前月份 preview。
- **改变月份后不能直接写入**：这是安全保护。请重新点击“扫描并生成预览”，防止误读上一次月份的结果。
- **界面点击后暂时不能再点按钮**：扫描或写入正在后台执行，完成后按钮会自动恢复。
- **EXE 双击无反应**：先确认 Microsoft Office 已安装；也可从 PowerShell 启动 Python 版查看详细错误。
- **打包时报缺少模块**：先运行 `.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt`，再运行 `build_exe.ps1`。
- **上传后原文件会被改吗**：不会。GUI 复制到内部 `workspace/`，从不覆盖原始文件。
- **同名材料怎么办**：自动增加 `_1`、`_2`，不会静默覆盖。
- **最近使用的文件不见了**：界面会显示“路径不存在，请重新选择”，重新上传即可。
- **材料显示“解析失败”**：打开“解析日志”查看对应文件原因；确认文件未损坏且 Word/Excel 已安装。

## PyInstaller 手工打包

先安装运行及打包依赖：

```powershell
pip install -r requirements.txt
pip install -r requirements-build.txt
```

基础命令为：

```powershell
pyinstaller --onefile --windowed --name 月度工作量表自动生成器 gui_app.py
```

本项目还需要打包 `config.json`、`office_bridge.ps1`、RapidOCR 的 ONNX 模型、ONNX Runtime/OpenCV DLL 及 Office/PDF 隐藏导入，因此推荐使用已经配好的 spec。`gui_app.spec` 使用 `collect_all` 收集 `rapidocr_onnxruntime`、`onnxruntime`、`cv2`、`numpy` 和 `PIL`：

```powershell
pyinstaller --noconfirm --clean gui_app.spec
```

或直接运行 `build_exe.ps1`。生成结果位于 `dist/月度工作量表自动生成器.exe`。

## 制作给非程序员的发布包

发布版只包含 EXE 和说明文件，禁止复制 `workspace/`、`app_settings.json`、真实模板、真实课表、照片或输出 Excel。发布包结构为：

```text
release/
  月度工作量表自动生成器.exe
  启动说明.txt
  使用说明.md
  示例材料放这里.txt
  月度工作量表自动生成器_发布版.zip
```

朋友解压后双击 EXE 即可，不需要安装 Python。为了读取旧版 `.doc/.xls` 并严格保持 `.xls` 格式，建议安装 Microsoft Word 和 Excel；仅 WPS 环境可能无法完整支持 Office COM。

## 界面截图

如需在仓库展示 GUI 截图，请使用脱敏材料运行程序后，将截图保存到 `docs/images/gui-main.png`。不要把真实课程名称、教师数据、工作量表或照片课表提交到 GitHub。

## 版本库建议

`output/` 中包含可重复生成的 CSV、日志、解析快照、正式 Excel 及业务数据，默认已加入 `.gitignore`，建议本地保留但不要提交。若需要在仓库中提供演示文件，建议脱敏后放入单独的 `examples/` 目录，不要直接提交真实工作量表或课程通知。
