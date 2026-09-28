# msaccess-vcs-mcp（占位）

本目录**未随本仓库分发**，仅作占位说明。

## 内容
VCS 加载项的 MCP 封装（Python 包，版本 0.1.0，作者 Adam Waller）：把加载项能力暴露为 MCP 工具——导出 `vcs_export`、快速保存 `vcs_fast_save`、按对象导出/导入、`vcs_build` / `vcs_merge_build`、对象清单、差异比较、只读 SQL 查询、VBA 调用与构建进度回调等，共 70 个文件。

## 官方来源
- 官方仓库：<https://github.com/joyfullservice/msaccess-vcs-mcp>

## 在技能中的用途
技能版本控制三通道中的 **MCP 通道**：供 AI agent（Cursor / Claude Code 等 MCP 客户端）经 MCP 操纵 VCS 加载项，完成 Access 数据库的文本源导出与构建（与 `Microsoft Access Version Control System/` 的加载项通道配合使用）。
