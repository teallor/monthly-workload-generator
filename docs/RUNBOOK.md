# 运行手册（无需懂代码）

原创作者：Rafael_Huang；CODEX 辅助实现与文档整理。

当前版本：v1.0正式验收版。朋友端到端测试已通过。

## 推荐：双击 EXE

1. 解压发布包到普通文件夹，不要直接在压缩包内运行。
   建议先用测试材料熟悉流程，不要直接处理唯一原件；重要文件应提前备份。
2. 双击 `月度工作量表自动生成器.exe`。
3. 上传 `.xls/.xlsx` 工作量表模板。
4. 上传 Word、PDF、Excel 或照片课表。
5. 设置目标年月和教师；照片需自动识别时勾选 OCR。
6. 点击“生成预览”，核对课程、课时、分类、教师和排除原因。
7. 处理所有待确认项后，点击“生成最终 Excel”。
8. 输出位于 `workspace/output/<YYYY-MM>/`。

如果同名最终 Excel 已存在，程序自动保存为 `_1`、`_2` 等新文件，不覆盖原文件；生成完成界面会显示实际路径。当前界面提供“手工补录课程”，没有按特定工种设置的专用快速补录入口。

朋友电脑不需要 Python。处理旧 `.doc/.xls` 并严格保持格式时，建议安装 Microsoft Word 和 Excel。

## 源码启动

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python gui_app.py
```

## CLI

```powershell
python main.py --target-month 2026-08 --preview --input-dir ".\materials\2026-08" --output-dir output
python main.py --target-month 2026-08 --write --input-dir ".\materials\2026-08" --output-dir output
```

CLI 写入只有输入 `y` 才继续。

## OCR 自检

在 GUI 点击“检测 OCR”或“工具 → OCR 自检”。OCR 结果必须核对 preview。
OCR 评分使用用户选择的目标月份，不固定为 7 月。

## 打包

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

输出为 `dist/月度工作量表自动生成器.exe`。发布时不复制 workspace、设置、真实材料或 output。

## 每月检查单

- 模板和目标教师正确。
- 年月正确，所有材料已导入。
- preview 中培训、考核数量和总课时合理。
- 待确认项已处理。
- 最终 Excel 能打开、标题月份正确、其他模块未改。
