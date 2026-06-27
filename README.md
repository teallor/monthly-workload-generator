# 任意月份工作量表自动生成程序

本程序用于生成任意年份、任意月份的个人工作量表。它扫描课程通知、课程表、PDF、Excel 和图片依据，只提取目标教师在目标月份内的课程，区分培训与考核，并基于原 Excel 模板生成新文件。程序始终先生成 preview；只有同月份 preview 已存在、校验通过且用户输入 `y` 后，才会复制模板并写入最终 Excel。

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

## 文件如何放置

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
```

只输入 `8月` 时，程序优先从已选模板文件名推断年份；模板不含年份时使用当前年份，并在解析日志中明确记录推断依据。

### Write 示例

```powershell
python main.py --target-month 2026-07 --write
python main.py --target-month 2026-08 --write
python main.py --target-month 2026-09 --write
```

Write 不会覆盖模板。输出文件名和表头年份月份都会根据 `--target-month` 自动更新。
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

## 图片人工校核侧车

当本机未启用 OCR，或复杂照片表格无法可靠 OCR 时，可在图片旁放置同名的 `<图片文件名>.records.json`。程序会优先读取这份人工校核记录，并保留原图片文件名作为来源。本项目中的照片课表已使用该机制录入经用户确认的日期、教师简称和课程内容。

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
- **图片没有识别**：默认 `enable_ocr=false`。可启用 OCR，或为复杂图片提供 `<图片名>.records.json` 人工校核侧车。
- **教师简称未命中**：检查 `config.json` 中的 `teacher_aliases`。`黄`、`黄/王`、`黄、王` 均可命中黄佳豪。
- **write 被阻止**：先运行同月份 preview；日期、分类、课时或图片 OCR 仍待确认时，程序不会写 Excel。
- **教学槽位不足**：程序会报告培训行数或考核汇总槽位不足，未经确认不会扩展模板。
- **preview 被覆盖**：不会。每个月份分别写入 `<output-dir>/<YYYY-MM>/`。

## 版本库建议

`output/` 中包含可重复生成的 CSV、日志、解析快照、正式 Excel 及业务数据，默认已加入 `.gitignore`，建议本地保留但不要提交。若需要在仓库中提供演示文件，建议脱敏后放入单独的 `examples/` 目录，不要直接提交真实工作量表或课程通知。
