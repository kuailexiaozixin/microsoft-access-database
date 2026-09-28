#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""上游漂移检测器（对齐 xlwings / fastapi 的 manifest + SYNCLOG 机制）。

本技能 vendored（引入并锁定）了多路开源 / 商业上游内容。本脚本检测两类漂移：

  1. 本地漂移（LOCAL_DRIFT）：vendored 目录被动过——新增 / 删除 / 改大小 / 改内容。
     依据 manifest.json 里每条来源注册时计算的基线指纹（baseline_* 字段）与当前实时状态比对。
     这是**离线可用**的核心能力：vendored 内容"原文件永不手改"，一旦本地副本偏离基线即报警。
  2. 上游更新（UPSTREAM_AHEAD）：对登记了 repo + ref 的上游，经 `git ls-remote` 查询该 ref 的
     最新 commit，与 pinned_sha 比对（pinned_sha 为空则只报告最新 commit）。需要网络；不可达
     时降级跳过并提示，不影响本地漂移判断。

用法:
    python scripts/check_upstream_drift.py                 # 检测 + 打印报告；本地漂移→退出码 1
    python scripts/check_upstream_drift.py --register      # 重算基线并写回 manifest.json（合法更新后调用）
    python scripts/check_upstream_drift.py --quiet         # 仅打印摘要 + 退出码（适合 CI）
    python scripts/check_upstream_drift.py --no-upstream   # 跳过在线上游检查（纯离线）

基线指纹算法（对齐"原文件永不手改"原则，离线可复现）：
    struct_hash  = sha256( 按相对路径排序后的 "相对路径|字节大小\\n" )
    content_hash = sha256( 按相对路径排序后的 每文件 sha256 拼接 )
  检测时先比 struct_hash（仅 stat，极快）：相同 → UNCHANGED；不同 → 再比 content_hash 并列出变动文件，
  以区分"真正内容改动"与"云盘同步导致的同名同大小副本文"（后者内容哈希一致，只提示复核）。

约束:
    - 本脚本只读写 manifest.json / SYNCLOG.md，绝不改动任何 local_dir 内容。
    - 本地漂移默认导致非零退出码（可被门禁拦截）；上游更新默认只提示、不报错。
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
MANIFEST = SKILL_ROOT / "manifest.json"
SYNCLOG = SKILL_ROOT / "SYNCLOG.md"
GIT_TIMEOUT = 20  # 单次 git ls-remote 超时（秒）


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def walk_files(local_dir: Path, only_relpaths=None):
    """返回 [(相对路径, 字节大小)]，按相对路径排序；含二进制（.accdb 等）。

    only_relpaths 不为 None 时，只遍历 local_dir 下这些子路径（相对 local_dir），
    用于 examples 这类"混合容器"——只跟踪其中的开源积木，排除技能自有件。
    相对路径始终相对于 local_dir 计算。
    """
    rows = []
    if only_relpaths:
        roots = [os.path.join(local_dir, r) for r in only_relpaths]
    else:
        roots = [local_dir]
    for root in roots:
        if not os.path.isdir(root):
            continue
        for dirpath, _dirnames, filenames in os.walk(root):
            for fn in filenames:
                fp = os.path.join(dirpath, fn)
                try:
                    sz = os.path.getsize(fp)
                except OSError:
                    continue
                rel = os.path.relpath(fp, local_dir).replace(os.sep, "/")
                rows.append((rel, sz))
    rows.sort()
    return rows


def is_placeholder(local_dir):
    """判断目录是否为"占位目录"——仅含 README.md（发布态用占位链接替代真实 vendored 二进制）。

    本地完整 checkout 的 vendored 目录含真实内容（.accdb/.dll 等），不算占位；
    发布态把重二进制目录压成仅 README.md 占位，漂移门禁应跳过这类目录，避免误报。
    接受 str 或 Path。
    """
    local_dir = Path(local_dir)
    if not local_dir.is_dir():
        return False
    for dirpath, _dirnames, filenames in os.walk(local_dir):
        for fn in filenames:
            if fn != "README.md":
                return False
    return True


def source_is_placeholder(local_dir, only_relpaths=None):
    """整个来源在当前 checkout 里是否被压成占位（发布态）。

    - 无 track_relpaths：等价于 is_placeholder(local_dir)（目录仅含 README.md）。
    - 有 track_relpaths（如 examples 只跟踪其中若干开源积木）：
      当**所有**被跟踪子路径都存在且都是"仅 README.md"的占位目录时，整个来源视为占位，
      漂移门禁应跳过——发布态把开源积木目录压成占位，避免误报。
      只要有任一被跟踪子路径缺失或含真实内容，就按正常漂移逻辑处理。
    接受 str 或 Path。
    """
    local_dir = Path(local_dir)
    if not only_relpaths:
        return is_placeholder(local_dir)
    seen = 0
    for r in only_relpaths:
        p = local_dir / r
        if not p.is_dir():
            return False
        if not is_placeholder(p):
            return False
        seen += 1
    return seen > 0


def compute_struct(local_dir: Path, only_relpaths=None):
    rows = walk_files(local_dir, only_relpaths)
    payload = "".join("%s|%d\n" % (r, s) for r, s in rows).encode("utf-8")
    total = sum(s for _, s in rows)
    return sha256_bytes(payload), len(rows), total


def compute_content(local_dir: Path, only_relpaths=None):
    parts = []
    if only_relpaths:
        roots = [os.path.join(local_dir, r) for r in only_relpaths]
    else:
        roots = [local_dir]
    for root in roots:
        if not os.path.isdir(root):
            continue
        for dirpath, _dirnames, filenames in os.walk(root):
            for fn in filenames:
                fp = os.path.join(dirpath, fn)
                try:
                    with open(fp, "rb") as f:
                        data = f.read()
                except OSError:
                    continue
                parts.append(sha256_bytes(data))
    parts.sort()
    return sha256_bytes("".join(parts).encode("utf-8"))


def changed_files(local_dir: Path, baseline_map):
    """baseline_map: {相对路径: 大小}。返回 (added, removed, resized) 三个列表。"""
    cur = {r: s for r, s in walk_files(local_dir)}
    added = [r for r in cur if r not in baseline_map]
    removed = [r for r in baseline_map if r not in cur]
    resized = [r for r in cur if r in baseline_map and cur[r] != baseline_map[r]]
    return added, removed, resized


def check_upstream(repo: str | None, ref: str | None):
    """返回 (latest_sha, error_or_None)。需要网络；失败返回 (None, 原因)。"""
    if not repo or not ref:
        return None, "未登记 repo/ref，跳过"
    url = repo if repo.startswith("http") else "https://github.com/%s" % repo
    try:
        out = subprocess.run(
            ["git", "ls-remote", "--refs", url, ref],
            capture_output=True, text=True, timeout=GIT_TIMEOUT,
        )
    except Exception as e:  # 超时 / git 不存在 / 网络不可达
        return None, "git ls-remote 失败: %s" % e
    if out.returncode != 0:
        return None, "git ls-remote 不可达（可能离线/需代理）"
    line = out.stdout.strip().splitlines()
    if not line:
        return None, "ref 在远端不存在"
    return line[0].split()[0], None


def load_manifest():
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def save_manifest(data):
    MANIFEST.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def register():
    data = load_manifest()
    today = date.today().isoformat()
    for s in data["sources"]:
        ld = SKILL_ROOT / s["local_dir"]
        if not ld.is_dir():
            print("  !! 跳过（local_dir 不存在）: %s" % s["local_dir"])
            continue
        if source_is_placeholder(ld, s.get("track_relpaths")):
            # 占位态（发布仓库仅有 README）没有真实内容，若照算会用占位数据覆盖完整基线
            print("  .. 跳过（占位态，不覆盖基线）: %s" % s["local_dir"])
            continue
        struct_h, count, total = compute_struct(ld, s.get("track_relpaths"))
        content_h = compute_content(ld, s.get("track_relpaths"))
        s["registered_date"] = today
        s["baseline_file_count"] = count
        s["baseline_total_bytes"] = total
        s["baseline_struct_hash"] = struct_h
        s["baseline_content_hash"] = content_h
        print("  已登记基线: %-50s 文件=%d 大小=%dMB" % (s["local_dir"], count, total // 1024 // 1024))
    save_manifest(data)
    print("已写回 %s" % MANIFEST.name)


def append_synclog(heading, rows):
    stamp = date.today().isoformat()
    block = "\n| %s | %s |\n" % (stamp, heading)
    for r in rows:
        block += "| %s |\n" % " | ".join(str(x) for x in r)
    # 追加到文件末尾（SYNCLOG 以日志流形式累积）
    with SYNCLOG.open("a", encoding="utf-8") as f:
        f.write("\n" + block)


def main():
    args = set(sys.argv[1:])
    do_register = "--register" in args
    quiet = "--quiet" in args
    no_upstream = "--no-upstream" in args

    if not MANIFEST.is_file():
        print("错误：找不到 %s" % MANIFEST)
        return 2
    if not SYNCLOG.is_file():
        SYNCLOG.write_text("# 上游同步日志\n\n", encoding="utf-8")

    data = load_manifest()
    sources = data["sources"]

    if do_register:
        print("=== 重算并登记基线 ===")
        register()
        return 0

    print("=" * 78)
    print("上游漂移检测（扫描根: %s）" % SKILL_ROOT)
    print("=" * 78)
    local_drift = False
    upstream_info = []
    for s in sources:
        ld = SKILL_ROOT / s["local_dir"]
        if not ld.is_dir():
            print("  [缺失] %-50s local_dir 不存在" % s["local_dir"])
            local_drift = True
            continue
        if source_is_placeholder(ld, s.get("track_relpaths")):
            print("  [%-26s] %-22s %-14s" % (s["local_dir"][:26], s.get("kind", ""), "PLACEHOLDER(跳过)"))
            continue
        cur_struct, cur_count, cur_total = compute_struct(ld, s.get("track_relpaths"))
        base_struct = s.get("baseline_struct_hash")
        base_content = s.get("baseline_content_hash")
        base_count = s.get("baseline_file_count")
        base_bytes = s.get("baseline_total_bytes")

        if base_struct is None:
            status = "未登记基线（运行 --register）"
        elif cur_struct == base_struct:
            status = "UNCHANGED"
        else:
            # 结构变了：算内容哈希确认是否真漂移
            added, removed, resized = changed_files(ld, {})
            cur_content = compute_content(ld, s.get("track_relpaths"))
            if base_content and cur_content == base_content:
                status = "STRUCT_CHANGED_CONTENT_SAME（疑似云盘同步副本文，请复核）"
            else:
                local_drift = True
                status = "LOCAL_DRIFT"
                # 列出变动文件（重新读基线结构需要原 map；此处以 size 差速判，简单列 added/removed/renamed）
                print("       变动: +%d 文件 / -%d 文件 / ~%d 改大小（详见下方）" % (len(added), len(removed), len(resized)))
                for r in added[:10]:
                    print("         + %s" % r)
                for r in removed[:10]:
                    print("         - %s" % r)
                for r in resized[:10]:
                    print("         ~ %s" % r)
        detail_count = "%d→%d 文件" % (base_count or 0, cur_count) if base_count is not None else "-"
        detail_bytes = ("%d→%d MB" % (base_bytes // 1024 // 1024, cur_total // 1024 // 1024)) if base_bytes else "-"
        print("  [%-26s] %-22s %-14s %s / %s" % (s["local_dir"][:26], s.get("kind", ""), status, detail_count, detail_bytes))

        # 上游更新检查
        if not no_upstream and s.get("repo") and s.get("ref"):
            latest, err = check_upstream(s["repo"], s["ref"])
            if err:
                upstream_info.append((s["local_dir"], "上游检查跳过: %s" % err))
            elif latest and latest != s.get("pinned_sha"):
                upstream_info.append((s["local_dir"], "上游最新 %s（pinned=%s）" % (latest[:10], (s.get("pinned_sha") or "空")[:10])))

    print("-" * 78)
    if upstream_info:
        print("上游更新提示：")
        for name, msg in upstream_info:
            print("  · %s: %s" % (name, msg))
    print("本地漂移：%s" % ("有（见上）" if local_drift else "无"))
    print("=" * 78)
    return 1 if local_drift else 0


if __name__ == "__main__":
    sys.exit(main())
