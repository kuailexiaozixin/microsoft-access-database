# 上游同步日志

> 本文件以"日志流"形式累积记录上游 vendored 内容的同步 / 漂移事件，与 `manifest.json`（来源登记 + 基线指纹）和 `scripts/check_upstream_drift.py`（漂移检测）配套。
>
> 原则：**vendored 文件原文件永不手改**。任何对 `manifest.json` 中登记的 7 路上游目录（`Microsoft Access Version Control System` / `msaccess-vcs-addin` / `msaccess-vcs-mcp` / `Version_Control_v5.0.1` / `盟威Access快速开发平台V2.7.0版(64位)` / `Edonsoft Development Framework_x64` / `examples`）的主动改动，都必须走"整体替换 + 重新 `--register` 基线"流程，并在本文件追加一条记录。
>
> 列含义：`日期 | 事件`（`事件`内用次级条目列明对象 / 来源 / 动作 / 影响）。

| 日期 | 事件 |
| --- | --- |
| 2026-09-11 | **机制建立 + 基线登记**：参照 `xlwings` / `fastapi` 的 manifest + SYNCLOG 机制，为 7 路上游建立漂移跟踪。落盘 `manifest.json`（含每路来源、kind、repo/ref/version 及基线指纹）与 `scripts/check_upstream_drift.py`（本地 struct/content 双哈希离线漂移检测 + `git ls-remote` 在线上游更新检测 + `--register` 写基线）。基线指纹：MAVCS 分发物 5 文件 / 7MB；msaccess-vcs-addin 552 文件 / 8MB；msaccess-vcs-mcp 70 文件；Version_Control_v5.0.1 1 文件 / 13MB；盟威平台 563 文件 / 18MB；Edonsoft 框架 122 文件 / 14MB；examples 95 文件 / 15MB。初始检测全部 UNCHANGED。 |
