#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""上游自动拉取 + 增量 / 整体刷新（对齐 xlwings 的 sync_upstream.py）。

本技能 vendored（引入并锁定）了多路上游。原有 check_upstream_drift.py 只做"离线本地漂移 +
在线上游 sha 提示"，本脚本把能力补齐为"检测 + 自动拉取 + 刷新"：

  git 源（manifest 里 sync=="git"，登记了 repo + ref）：
      经 codeload tarball 拉取最新 ref，解包后**原子替换** local_dir（整体刷新；
      增量语义由 git 历史保证，无需逐文件 diff）。随后把最新 commit 写回 pinned_sha。

  下载源（sync=="download"，如 Edonsoft x64 rar）：
      1. 抓取 DownloadCounter.ashx?action=download&edition=x64（返回 .rar）；
      2. 用 tools/unrar/UnRAR.exe 解包（该包是 RAR 1.5，7za 解不了，必须用 UnRAR）；
      3. 本地化补丁（localize）：改写 appsettings.json 的 DbConnect 与前端 25 张链接表的
         Connect，重指到同目录 Data.accdb（同 v202610.01 的有意改动，使框架离线可运行，
         并阻止框架自身启动时的自动重链改动基线）；
      4. **原子替换 vendored 子集**——只刷新包内件（appsettings.json / Data.accdb /
         EdonSoft Development Framework.accdb / edonsoftCore.accde / assets/），
         保留 edonsoft-docs/（该目录由本技能从帮助站抓取，不在 rar 内）；
      5. 更新 manifest 的 version / pinned_version，并调用 check_upstream_drift.py --register 重算基线。

增量更新的语义：git 源走 tarball 整体刷新（等价增量，历史即增量）；下载源（二进制包）按"包"
整体刷新 vendored 子集，edonsoft-docs 等本技能自有件不被覆盖——这就是下载型上游的自然粒度。

漂移检测交给 track_upstream.py（在线上游版本 / sha）+ check_upstream_drift.py（离线本地结构/内容哈希）。
本脚本只负责"拉取"，不负责报警。

用法:
    python scripts/sync_upstream.py                 # 同 --check
    python scripts/sync_upstream.py --check         # 逐源报告是否需要更新（不改动任何文件）
    python scripts/sync_upstream.py --ensure        # 把所有"落后"的源拉取到最新
    python scripts/sync_upstream.py --force <id>    # 强制拉取指定源（忽略版本是否变化）
    python scripts/sync_upstream.py --no-network    # 仅离线自检（跳过下载 / 版本探测）

注意: Edonsoft 本地化补丁需要本机安装 Microsoft Access 且 Python 带 pywin32（EnsureDispatch）。
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.request
from datetime import date
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
MANIFEST = SKILL_ROOT / "manifest.json"
SYNCLOG = SKILL_ROOT / "SYNCLOG.md"
UNRAR = SKILL_ROOT / "tools" / "unrar" / "UnRAR.exe"

HTTP_TIMEOUT = 120
GIT_TIMEOUT = 25

# 下载源配置
EDONSOFT_PAGE = "https://www.edonsoft.com/access-framework"
EDONSOFT_DL = "https://www.edonsoft.com/DownloadCounter.ashx?action=download&edition=x64"

_VER_BLOCK = re.compile(r'<span class="version-number">v(\d{6})\.(\d{2})</span>')
_DATE_BLOCK = re.compile(r'<span class="version-date">(\d{4}-\d{2}-\d{2})</span>')


def log(*a):
    print("[sync]", *a)


def load_manifest():
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def save_manifest(d):
    MANIFEST.write_text(json.dumps(d, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------
# 网络
# --------------------------------------------------------------------------
def http_get(url, binary=True):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (sync_upstream)"})
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as r:
        data = r.read()
    return data


# --------------------------------------------------------------------------
# 版本探测（Edonsoft 下载源）
# --------------------------------------------------------------------------
def probe_edonsoft_version():
    """抓取官网，返回 (latest_version_str, release_date_str)。无网络 / 未找到抛异常。

    官网版本历史区结构：
        <span class="version-number">v202611.02</span>
        <span class="version-tag tag-latest">当前版本</span>
        <span class="version-date">2026-11-01</span>
    取数值最大的 version-number 为最新版；日期取其后紧邻的 version-date。
    """
    html = http_get(EDONSOFT_PAGE, binary=False).decode("utf-8", "ignore")
    nums = _VER_BLOCK.findall(html)
    if not nums:
        raise RuntimeError("官网未找到 '<span class=\"version-number\">vYYYYMM.DD</span>'")
    best = max(nums, key=lambda t: (int(t[0]), int(t[1])))
    latest = "v%s.%s" % (best[0], best[1])
    idx = html.find('<span class="version-number">%s</span>' % latest)
    rel = None
    if idx >= 0:
        dm = _DATE_BLOCK.search(html, idx, idx + 400)
        if dm:
            rel = dm.group(1)
    return latest, rel


# --------------------------------------------------------------------------
# git 源：最新 commit
# --------------------------------------------------------------------------
def git_latest_sha(repo, ref):
    url = repo if repo.startswith("http") else "https://github.com/%s" % repo
    try:
        out = subprocess.run(
            ["git", "ls-remote", "--refs", url, ref],
            capture_output=True, text=True, timeout=GIT_TIMEOUT,
        )
    except Exception as e:  # 超时 / git 不存在 / 网络不可达
        return None, "git ls-remote 失败: %s" % e
    if out.returncode != 0:
        return None, "git ls-remote 不可达（可能离线 / 需代理）"
    lines = out.stdout.strip().splitlines()
    if not lines:
        return None, "ref 在远端不存在"
    return lines[0].split()[0], None


# --------------------------------------------------------------------------
# Edonsoft 解包 + 本地化
# --------------------------------------------------------------------------
def run_unrar(out_dir: Path, rar_path: Path):
    if not UNRAR.is_file():
        raise RuntimeError("缺少 tools/unrar/UnRAR.exe（Edonsoft 包是 RAR 1.5，只能用 UnRAR 解）")
    out_dir.mkdir(parents=True, exist_ok=True)
    r = subprocess.run(
        [str(UNRAR), "x", "-o+", "-y", str(rar_path), str(out_dir) + "\\"],
        capture_output=True, text=False,
    )
    if r.returncode != 0:
        err = (r.stderr or r.stdout or b"").decode("utf-8", "ignore")
        raise RuntimeError("UnRAR 解包失败: " + err[:400])


def localize_edonsoft(work_dir: Path):
    """本地化补丁（在最终落盘位置执行）：appsettings.json 的 DbConnect + 前端链接表 Connect
    重指到同目录 Data.accdb。

    注意: Edonsoft 前端在 Access 中打开时会按"前端所在目录"自动重链链接表与后端（框架设计），
    因此本函数在**最终落盘目录**执行——无论是框架自动重链还是下方显式 RefreshLink，目标都是
    同目录 Data.accdb，二者一致，基线稳定。绝不要在临时解包目录里重链，否则框架会把链接写回临时目录。
    """
    backend_win = str(work_dir / "Data.accdb")  # 原生反斜杠绝对路径

    # 1) appsettings.json（纯文本，直接改）
    asp = work_dir / "appsettings.json"
    if asp.is_file():
        cfg = json.loads(asp.read_text(encoding="utf-8"))
        for b in cfg.get("Backends", []):
            if b.get("Type") == "Access":
                b["DbConnect"] = "MS Access;DATABASE=%s" % backend_win
                b["AccessPath"] = ".\\Data.accdb"
        asp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        log("  localize: appsettings.json DbConnect ->", backend_win)

    # 2) 前端链接表（DAO RefreshLink；框架自动重链也会落到同目录，二者一致）
    fe = work_dir / "EdonSoft Development Framework.accdb"
    if fe.is_file():
        relink_access_tables(fe, work_dir / "Data.accdb")


def relink_access_tables(frontend: Path, backend: Path):
    import sys as _sys
    import win32com.client

    _sys.coinit_flags = 0
    backend_str = str(backend)
    app = win32com.client.gencache.EnsureDispatch("Access.Application")
    app.Visible = False
    app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable：阻止框架启动宏改写基线
    try:
        app.OpenCurrentDatabase(str(frontend), True)  # exclusive=True
        db = app.CurrentDb()
        n = 0
        for t in db.TableDefs:
            conn = t.Connect
            if conn and "Data.accdb" in conn:
                t.Connect = ";DATABASE=%s" % backend_str
                t.RefreshLink()
                n += 1
        app.CloseCurrentDatabase()
        log("  localize: 重链 %d 张链接表 -> %s" % (n, backend_str))
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def replace_vendored_subset(target_dir: Path, extracted_dir: Path, keep=None):
    """原子替换：把 extracted_dir 的包内件搬进 target_dir，保留 keep 列表（如 edonsoft-docs）。

    步骤：建 .stage -> 拷贝包内件 + 拷贝 keep 件 -> 校验关键文件 -> 交换（target->.old, stage->target）
    -> 删除 .old（被替换的 vendored 二进制，可再生，非用户个人文件）。
    """
    keep = keep or []
    pkg_items = [p.name for p in extracted_dir.iterdir()]
    for must in ("appsettings.json", "Data.accdb", "EdonSoft Development Framework.accdb", "edonsoftCore.accde"):
        if must not in pkg_items:
            raise RuntimeError("包内缺少预期文件: %s（实际: %s）" % (must, pkg_items))

    stage = target_dir.with_name(target_dir.name + ".stage")
    old = target_dir.with_name(target_dir.name + ".old")
    if stage.exists():
        shutil.rmtree(stage)
    if old.exists():
        shutil.rmtree(old)
    stage.mkdir(parents=True)

    for item in extracted_dir.iterdir():
        dest = stage / item.name
        if item.is_dir():
            shutil.copytree(item, dest)
        else:
            shutil.copy2(item, dest)
    for k in keep:
        src = target_dir / k
        if src.exists():
            dest = stage / k
            if src.is_dir():
                shutil.copytree(src, dest)
            else:
                shutil.copy2(src, dest)

    target_dir.rename(old)
    stage.rename(target_dir)
    shutil.rmtree(old)  # 被替换的 vendored 二进制，可再生，非用户个人文件
    log("  原子替换完成；保留项:", keep)


# --------------------------------------------------------------------------
# 拉取：Edonsoft 下载源
# --------------------------------------------------------------------------
def pull_edonsoft(source):
    target = SKILL_ROOT / source["local_dir"]
    latest, rel = probe_edonsoft_version()
    log("Edonsoft 上游版本: %s (%s)" % (latest, rel))

    tmp = Path(tempfile.mkdtemp(prefix="edonsoft_"))
    rar = tmp / "edonsoft_x64.rar"
    log("下载 rar ...")
    rar.write_bytes(http_get(EDONSOFT_DL))
    ext = tmp / "extracted"
    log("解包 ...")
    run_unrar(ext, rar)
    log("原子替换 vendored 子集（保留 edonsoft-docs）...")
    replace_vendored_subset(target, ext, keep=["edonsoft-docs"])
    log("本地化补丁（在落盘位置重链到同目录 Data.accdb）...")
    localize_edonsoft(target)

    update_manifest_version(source["id"], latest, rel)
    log("重算基线 ...")
    subprocess.run([sys.executable, str(SKILL_ROOT / "scripts" / "check_upstream_drift.py"), "--register"], check=True)
    append_synclog(source["id"], latest, rel)
    log("Edonsoft 拉取完成 ->", latest)


# --------------------------------------------------------------------------
# 拉取：git 源
# --------------------------------------------------------------------------
def pull_git(source):
    repo = source["repo"]
    ref = source.get("ref") or "main"
    target = SKILL_ROOT / source["local_dir"]
    url = "https://codeload.github.com/%s/tarball/%s" % (repo, ref)
    log("git 源 %s 拉取 tarball (ref=%s) ..." % (repo, ref))

    tmp = Path(tempfile.mkdtemp(prefix="git_"))
    tgz = tmp / "archive.tgz"
    tgz.write_bytes(http_get(url))
    ext = tmp / "extracted"
    ext.mkdir()
    with tarfile.open(tgz) as tf:
        tf.extractall(ext)
    top = next(ext.iterdir())  # 单一顶层目录 repo-ref-hash/

    stage = target.with_name(target.name + ".stage")
    old = target.with_name(target.name + ".old")
    if stage.exists():
        shutil.rmtree(stage)
    if old.exists():
        shutil.rmtree(old)
    stage.mkdir(parents=True)
    for item in top.iterdir():
        dest = stage / item.name
        if item.is_dir():
            shutil.copytree(item, dest)
        else:
            shutil.copy2(item, dest)
    target.rename(old)
    stage.rename(target)
    shutil.rmtree(old)

    sha, err = git_latest_sha(repo, ref)
    if sha:
        update_manifest_sha(source["id"], sha)
        log("git 源 %s 刷新完成，pinned_sha=%s" % (repo, sha[:10]))
    else:
        log("git 源 %s 刷新完成（pinned_sha 未更新: %s）" % (repo, err))
    subprocess.run([sys.executable, str(SKILL_ROOT / "scripts" / "check_upstream_drift.py"), "--register"], check=True)
    append_synclog(source["id"], "git:%s" % (sha[:10] if sha else "n/a"), None)


# --------------------------------------------------------------------------
# manifest / SYNCLOG 维护
# --------------------------------------------------------------------------
def update_manifest_version(id_, version, rel):
    d = load_manifest()
    for s in d["sources"]:
        if s["id"] == id_:
            s["version"] = "%s (%s)" % (version, rel) if rel else version
            s["pinned_version"] = version
            s["registered_date"] = date.today().isoformat()
            break
    save_manifest(d)


def update_manifest_sha(id_, sha):
    d = load_manifest()
    for s in d["sources"]:
        if s["id"] == id_:
            s["pinned_sha"] = sha
            s["registered_date"] = date.today().isoformat()
            break
    save_manifest(d)


def append_synclog(id_, version, rel):
    stamp = date.today().isoformat()
    if not SYNCLOG.is_file():
        SYNCLOG.write_text("# 上游同步日志\n\n", encoding="utf-8")
    line = "| %s | %s | %s | auto-pull via sync_upstream.py |\n" % (stamp, id_, version)
    with SYNCLOG.open("a", encoding="utf-8") as f:
        f.write("\n" + line)


# --------------------------------------------------------------------------
# 落后判定 + 主流程
# --------------------------------------------------------------------------
def source_status(source, no_network):
    sync = source.get("sync")
    if sync == "git":
        if no_network:
            return False, "（--no-network 跳过）"
        sha, err = git_latest_sha(source["repo"], source.get("ref") or "main")
        if err or not sha:
            return False, "git 检查跳过: %s" % (err or "无 sha")
        behind = sha != source.get("pinned_sha")
        return behind, "上游 %s（pinned=%s）" % (sha[:10], (source.get("pinned_sha") or "空")[:10])
    if sync == "download" and source.get("version_probe") == "edonsoft":
        if no_network:
            return False, "（--no-network 跳过）"
        try:
            latest, rel = probe_edonsoft_version()
        except Exception as e:
            return False, "版本探测失败: %s" % e
        behind = latest != source.get("pinned_version")
        return behind, "上游 %s（本地 %s）" % (latest, source.get("pinned_version"))
    return False, "未配置自动拉取（仅漂移检测）"


def pull_source(source):
    sync = source.get("sync")
    if sync == "git":
        pull_git(source)
    elif sync == "download":
        pull_edonsoft(source)
    else:
        log("源 %s 未配置自动拉取，跳过" % source["id"])


def main():
    args = sys.argv[1:]
    mode = "check"
    if "--ensure" in args:
        mode = "ensure"
    no_network = "--no-network" in args

    force_ids = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--force":
            i += 1
            if i < len(args):
                force_ids.append(args[i])
        elif a.startswith("--force="):
            force_ids.append(a.split("=", 1)[1])
        i += 1

    if not MANIFEST.is_file():
        print("错误：找不到 %s" % MANIFEST)
        return 2
    data = load_manifest()
    sources = data["sources"]

    if force_ids:
        for fid in force_ids:
            s = next((x for x in sources if x["id"] == fid), None)
            if not s:
                log("未知源 id: %s（可用: %s）" % (fid, ", ".join(x["id"] for x in sources)))
                return 2
            log("强制拉取: %s" % fid)
            pull_source(s)
        return 0

    if mode == "ensure":
        did = False
        for s in sources:
            behind, msg = source_status(s, no_network)
            if behind:
                log("拉取落后源: %s (%s)" % (s["id"], msg))
                pull_source(s)
                did = True
        if not did:
            log("没有落后的自动拉取源，无需操作。")
        return 0

    # 默认 --check
    print("=" * 78)
    print("上游自动拉取自检（需要更新？）  扫描根: %s" % SKILL_ROOT)
    print("=" * 78)
    for s in sources:
        behind, msg = source_status(s, no_network)
        flag = "需更新" if behind else "已最新"
        print("  [%-48s] %-10s %s | %s" % (s["local_dir"][:48], s.get("sync", "-"), flag, msg))
    print("=" * 78)
    print("运行 `python scripts/sync_upstream.py --ensure` 拉取所有落后源；")
    print("或 `python scripts/sync_upstream.py --force <id>` 强制拉取指定源。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
