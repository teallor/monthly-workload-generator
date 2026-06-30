# 项目状态

检查日期：2026-06-30  
项目：月度工作量表自动生成器  
原创作者：Rafael_Huang  
许可证：MIT

## 当前结论

项目处于 **v1.0.0 稳定版发布准备阶段**。GUI/EXE 已完成朋友端到端测试和本机新版界面轻量冒烟；GitHub 目前仍停留在 `v1.0`，新版源码、文档和 EXE 尚未同步，因此不能声称 `v1.0.0` 已公开发布。

## 关键入口

- 正式用户入口：`dist/月度工作量表自动生成器.exe`
- GUI 源码入口：`gui_app.py`
- CLI 入口：`main.py`
- 业务服务：`workload_service.py`
- Excel 写入：`workload_writer.py`
- Office 兼容桥：`office_bridge.ps1`
- 打包：`build_exe.ps1`、`gui_app.spec`

## 文件检查

| 项目 | 状态 | 说明 |
|---|---|---|
| 新版 EXE | 存在 | `dist/月度工作量表自动生成器.exe` |
| 旧版备份 | 存在 | `dist/月度工作量表自动生成器_old.exe` |
| README | 存在 | 已标注作者、MIT、用途和使用流程 |
| LICENSE | 存在 | MIT，Copyright (c) 2026 Rafael_Huang |
| 运行依赖 | 存在 | `requirements.txt` |
| 构建依赖 | 存在 | `requirements-build.txt` |
| Git 仓库 | 已初始化 | 分支 `main`，远端 `origin` 已绑定 |
| GitHub | 已公开 | `teallor/monthly-workload-generator` |
| OCR | 已实现 | RapidOCR + ONNX Runtime；失败时提示并允许普通材料继续处理 |

## 数据与上传边界

根目录仍保留真实模板、通知、PDF、照片、输出和历史工作区，均属于本地业务资料，不得上传。`.gitignore` 已排除 Office 文件、图片、输出目录、工作区、设置、构建缓存、EXE 和历史 release 二进制包。

公开仓库只应包含源码、配置示例、依赖清单和脱敏文档。新版 EXE 应通过 GitHub Release 附件分发，不进入普通 Git 提交。

## 待完成

1. 审核本轮文档和标准发布目录。
2. 对提交候选执行隐私与大文件复核。
3. 提交并推送新版源码和文档。
4. 创建 `v1.0.0` 标签与 GitHub Release。
5. 上传新版发布 ZIP 和 SHA256，并验证下载链接。
