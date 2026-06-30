# 发布目录说明

- `latest/`：当前 v1.0.0 稳定版候选，供审核和后续制作 GitHub Release 附件。
- `backup/`：上一版 EXE，仅用于回退，不对外作为最新版分发。
- `历史归档/`、`v1.0_内测候选版/`、`v1.0_正式验收版/`：历史记录，保留但不再修改。
- release 根目录中早期 ZIP、EXE 和说明属于历史构建，不应与 `latest/` 混用。

公开发布时只使用 `latest/` 的干净内容制作 ZIP，不包含 `workspace/`、`app_settings.json`、真实模板、真实课表、照片、输出 Excel 或个人路径。
