# 后续维护提示词

## 开始维护前

> 请先阅读 docs/PROJECT_MEMORY.md、RUNBOOK.md、ACCEPTANCE_REPORT.md、CHANGELOG.md 和 TROUBLESHOOTING.md。检查 Git 状态，不删除真实材料、output、workspace 或历史验收记录。先说明范围；没有明确要求时不新增功能。

## 修复缺陷

> 只修复：【填写现象】。先复现并保存证据，定位根因，做最小修改。不得顺便重构或美化。回归正常 preview、中文空格路径、异常输入和 EXE 连续启动；更新 CHANGELOG 与验收报告。

## 生成新月份

> 生成【YYYY-MM】工作量表。先 preview，不直接 write。列出培训、考核、排除和待确认项；确认后再生成 Excel，不覆盖其他月份。

## 重新发布

> 使用现有 gui_app.spec 打包 onefile EXE。不得包含真实模板、课表、照片、workspace、output、app_settings.json 或个人路径。做干净启动和两次重启测试，记录 SHA256。

## 发布前审查

> 按 docs/ACCEPTANCE_REPORT.md 的 100 分表重新评分。低于 85 分不得进入候选发布；没有朋友端到端实测不得标为正式验收版。记录命令、退出码、失败原因和证据路径，禁止写“未知错误”。
