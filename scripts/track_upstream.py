#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Microsoft-Access-DataBase 技能上游漂移跟踪（三线检测，对齐 xlwings 的 track_upstream.py）。

本技能 vendored（引入并锁定）了多路上游。"检测漂移"由 check_upstream_drift.py 负责（离线本地哈希 +
在线上游 sha 提示）；本脚本在其之上补齐**事实锚点线**，形成与 xlwings 一致的"三线"：

  ① 上游版本/sha 线（提示级）—— manifest 各路 pinned_version / pinned_sha 与远端最新比对。
     · git 源（sync=="git"）：git ls-remote 最新 commit vs pinned_sha；pinned_sha 为空→"未锁定"
       （非漂移，跑 sync_upstream.py --ensure 即锁定）；不同→DRIFT（提示，跑 sync 同步）。
     · 下载源（sync=="download", version_probe=="edonsoft"）：官网探测最新版 vs pinned_version。
     · --strict 时，线① DRIFT 也计为失败。
  ② 事实锚点线（硬漂移，STALE 必修）—— references/upstream-facts.json 逐条对照本地快照实测
     （文件存在 / 字节大小 / 文件数 / Access 对象数）。STALE 说明自研文档对上游的断言已与快照不符，
     必须先修文档再引用（带病作业 = 断言失真）。需要 Access 的窗体/模块/宏/报表计数默认 SKIP，
     加 --deep 才核验（会改写前端字节，跑后须 --register）。
  ③ 版本一致性线（硬）—— manifest 的 version / pinned_version 必须与 upstream-facts.json 的
     snapshot.edonsoft_version 自洽；git 源 pinned_sha 缺失给出"未锁定"提示。

分工纪律:
    - sync_upstream.py 拉取上游后：必须跑本脚本 → 修 STALE 锚点指向的自研文档 → 重跑至 ②③ 全绿。
    - 引用上游对象数 / 文件大小前，可先跑本脚本确认锚点未破。
    - 新增上游事实断言时：登记进 upstream-facts.json，否则漂移无人知晓。

用法:
    python scripts/track_upstream.py              # 三线全查（窗体/模块等对象数 SKIP）
    python scripts/track_upstream.py --deep       # 线② 额外核验窗体/模块/宏/报表（需 Access）
    python scripts/track_upstream.py --quick      # 线① 只查 Edonsoft 下载源（核心）
    python scripts/track_upstream.py --line 2     # 只跑第②线（锚点）
    python scripts/track_upstream.py --strict     # 线① DRIFT 也计为失败
    python scripts/track_upstream.py --no-network # 线① 跳过在线上游检查（纯离线）

退出码:
    0 = 全绿（或仅线①提示级漂移且未加 --strict）
    1 = 线②锚点 STALE / 线③版本不一致 / --strict 下线①有漂移 / 检测本身失败
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

import sync_upstream as sync  # noqa: E402  复用 probe_edonsoft_version / git_latest_sha / MANIFEST

FACTS = SKILL_ROOT / "references" / "upstream-facts.json"


# ---------------------------------------------------------------- 线① 上游版本/sha
def line1_upstream(quick: bool, no_network: bool) -> tuple[list[str], int]:
    """返回 (报告行列表, 漂移路数)。"""
    manifest = json.loads(sync.MANIFEST.read_text(encoding="utf-8-sig"))
    lines: list[str] = []
    drifted = 0
    for src in manifest["sources"]:
        sid = src["id"]
        if quick and sid != "edonsoft-framework":
            continue
        if src.get("sync") == "git":
            repo, ref = src.get("repo"), src.get("ref") or "main"
            if no_network:
                lines.append("    [%s] SKIP（--no-network）" % sid)
                continue
            head, err = sync.git_latest_sha(repo, ref)
            pinned = src.get("pinned_sha")
            if head is None:
                lines.append("    [%s] UNREACHABLE（git 不可达，稍后重试）" % sid)
            elif not pinned:
                lines.append("    [%s] 未锁定（pinned_sha 空，跑 sync_upstream.py --ensure 锁定）" % sid)
            elif head != pinned:
                drifted += 1
                lines.append("    [%s] DRIFT %s → %s（跑 sync_upstream.py --ensure 同步）" % (sid, (pinned or "")[:10], head[:10]))
            else:
                lines.append("    [%s] OK %s" % (sid, head[:10]))
        elif src.get("sync") == "download" and src.get("version_probe") == "edonsoft":
            if no_network:
                lines.append("    [%s] SKIP（--no-network）" % sid)
                continue
            try:
                latest, rel = sync.probe_edonsoft_version()
            except Exception as e:
                lines.append("    [%s] UNREACHABLE（版本探测失败: %s）" % (sid, e))
                continue
            pinned = src.get("pinned_version")
            if latest != pinned:
                drifted += 1
                lines.append("    [%s] DRIFT %s → %s（跑 sync_upstream.py --force %s 同步）" % (sid, pinned or "?", latest, sid))
            else:
                lines.append("    [%s] OK %s（本地与上游一致）" % (sid, latest))
        else:
            lines.append("    [%s] 未配置自动拉取（仅基线漂移，见 check_upstream_drift.py）" % sid)
    return lines, drifted


# ---------------------------------------------------------------- 线② 事实锚点
_ENUM_CACHE: dict = {}


def _enum_dao(frontend: Path, backend: Path) -> dict:
    """DAO.DBEngine 只读枚举（不改写字节）：返回 (fe_*, be_*) 对象数 + 逐名清单。"""
    key = ("dao", str(frontend), str(backend))
    if key in _ENUM_CACHE:
        return _ENUM_CACHE[key]
    import win32com.client
    sys.coinit_flags = 0
    dao = win32com.client.gencache.EnsureDispatch("DAO.DBEngine.120")
    fe = dao.OpenDatabase(str(frontend), False, True)  # readonly
    local_names, linked_names = [], []
    for t in fe.TableDefs:
        n = t.Name
        if n.startswith("MSys") or n.startswith("~"):
            continue
        if t.Connect and "Data.accdb" in t.Connect:
            linked_names.append(n)
        else:
            local_names.append(n)
    qnames = [q.Name for q in fe.QueryDefs if not q.Name.startswith("~")]
    out = {"fe": {"tables_local": len(local_names), "tables_linked": len(linked_names),
                  "tables_total": len(local_names) + len(linked_names), "queries": len(qnames),
                  "tables_local_names": sorted(local_names),
                  "tables_linked_names": sorted(linked_names), "queries_names": sorted(qnames)}}
    fe.Close()
    be = dao.OpenDatabase(str(backend), False, True)
    be_tables = [t.Name for t in be.TableDefs if not t.Name.startswith(("MSys", "~"))]
    be_queries = [q.Name for q in be.QueryDefs if not q.Name.startswith("~")]
    out["be"] = {"tables": len(be_tables), "queries": len(be_queries),
                 "tables_names": sorted(be_tables), "queries_names": sorted(be_queries)}
    be.Close()
    _ENUM_CACHE[key] = out
    return out


def _enum_app(frontend: Path) -> dict:
    """Access.Application 枚举窗体/模块/宏/报表（会改写前端字节）。"""
    key = ("app", str(frontend))
    if key in _ENUM_CACHE:
        return _ENUM_CACHE[key]
    import win32com.client
    sys.coinit_flags = 0
    app = win32com.client.gencache.EnsureDispatch("Access.Application")
    app.Visible = False
    app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    app.OpenCurrentDatabase(str(frontend), False)
    cp = app.CurrentProject
    out = {"forms": cp.AllForms.Count, "modules": cp.AllModules.Count,
           "macros": cp.AllMacros.Count, "reports": cp.AllReports.Count,
           "forms_names": sorted(x.Name for x in cp.AllForms),
           "modules_names": sorted(x.Name for x in cp.AllModules)}
    app.CloseCurrentDatabase()
    app.Quit()
    _ENUM_CACHE[key] = out
    return out


def _check_fact(fact: dict, deep: bool) -> str | None:
    """返回 None=OK；"SKIP:..."=跳过（不计失败）；否则 STALE 错误描述。"""
    src_dir = SKILL_ROOT / fact["source"]
    target = src_dir / fact["file"]
    kind = fact["kind"]
    if kind == "exists":
        return None if target.exists() else "文件缺失 %s" % fact["file"]
    if kind == "size":
        if not target.is_file():
            return "文件缺失 %s" % fact["file"]
        sz = target.stat().st_size
        return None if sz == fact["value"] else "大小 %d ≠ 台账 %d（%s）" % (sz, fact["value"], fact["file"])
    if kind == "filecount":
        if not target.is_dir():
            return "目录缺失 %s" % fact["file"]
        pat = fact.get("pattern", "*")
        n = sum(1 for x in target.rglob(pat) if x.is_file())
        return None if n == fact["value"] else "文件数 %d ≠ 台账 %d（%s/%s）" % (n, fact["value"], fact["file"], pat)
    if kind in ("accdb_dao_frontend", "accdb_dao_backend"):
        try:
            fe = src_dir / "EdonSoft Development Framework.accdb"
            be = src_dir / "Data.accdb"
            data = _enum_dao(fe, be)
        except Exception as e:
            return "SKIP: DAO 枚举失败（%s）" % e
        got = data["fe"] if kind == "accdb_dao_frontend" else data["be"]
        for k, v in fact["value"].items():
            if got.get(k) != v:
                return "%s %s=%s ≠ 台账 %s" % (fact["file"], k, got.get(k), v)
        return None
    if kind == "accdb_app":
        if not deep:
            return "SKIP: 需 --deep（会改写前端字节，跑后须 --register）"
        try:
            fe = src_dir / "EdonSoft Development Framework.accdb"
            got = _enum_app(fe)
        except Exception as e:
            return "SKIP: Access 枚举失败（本机无 Access/pywin32？%s）" % e
        for k, v in fact["value"].items():
            if got.get(k) != v:
                return "%s %s=%s ≠ 台账 %s" % (fact["file"], k, got.get(k), v)
        return None
    if kind == "accdb_name_set":
        if not deep:
            return "SKIP: 需 --deep（会改写前端字节，跑后须 --register）"
        try:
            fe = src_dir / "EdonSoft Development Framework.accdb"
            be = src_dir / "Data.accdb"
            dao = _enum_dao(fe, be)
            app = _enum_app(fe)
            got = {
                "forms": app["forms_names"], "modules": app["modules_names"],
                "tables_local": dao["fe"]["tables_local_names"],
                "tables_linked": dao["fe"]["tables_linked_names"],
                "queries": dao["fe"]["queries_names"],
                "backend_tables": dao["be"]["tables_names"],
            }
        except Exception as e:
            return "SKIP: 对象名集合枚举失败（%s）" % e
        for k, want in fact["value"].items():
            cur = sorted(got.get(k, []))
            if cur != sorted(want):
                added = sorted(set(cur) - set(want))
                removed = sorted(set(want) - set(cur))
                return "对象名集合漂移 %s：新增 %s / 消失 %s" % (
                    k, (added[:5] or "无"), (removed[:5] or "无"))
        return None
    return "未知 kind: %s" % kind


def line2_facts(deep: bool) -> tuple[list[str], int, int]:
    """返回 (报告行, STALE 数, SKIP 数)。"""
    lines: list[str] = []
    stale = 0
    skipped = 0
    data = json.loads(FACTS.read_text(encoding="utf-8-sig"))
    for fact in data["facts"]:
        res = _check_fact(fact, deep)
        if res is None:
            lines.append("    [%s] OK" % fact["id"])
        elif res.startswith("SKIP:"):
            skipped += 1
            lines.append("    [%s] %s" % (fact["id"], res))
        else:
            stale += 1
            docs = ", ".join(fact.get("docs", [])) or "（无文档指针）"
            lines.append("    [%s] STALE %s → 修: %s" % (fact["id"], res, docs))
    return lines, stale, skipped


# ---------------------------------------------------------------- 线③ 版本一致性
def line3_consistency() -> tuple[list[str], int]:
    lines: list[str] = []
    broken = 0
    manifest = json.loads(sync.MANIFEST.read_text(encoding="utf-8-sig"))
    facts = json.loads(FACTS.read_text(encoding="utf-8-sig"))
    snap_ver = facts["snapshot"]["edonsoft_version"]
    entry = next((s for s in manifest["sources"] if s["id"] == "edonsoft-framework"), None)
    if entry is None:
        lines.append("    [edonsoft] STALE manifest 缺少 edonsoft-framework 条目")
        broken += 1
    else:
        pinned = entry.get("pinned_version")
        ver = entry.get("version", "")
        if pinned != snap_ver:
            broken += 1
            lines.append("    [edonsoft] STALE pinned_version=%s ≠ facts %s" % (pinned, snap_ver))
        elif snap_ver not in ver:
            broken += 1
            lines.append("    [edonsoft] STALE version='%s' 未含 pinned_version=%s" % (ver, snap_ver))
        else:
            lines.append("    [edonsoft] OK version/pinned_version 与 facts 自洽 (%s)" % snap_ver)
    # git 源锁定提示
    for s in manifest["sources"]:
        if s.get("sync") == "git" and not s.get("pinned_sha"):
            lines.append("    [%s] 提示：pinned_sha 未锁定（跑 sync_upstream.py --ensure 锁定）" % s["id"])
    return lines, broken


# ---------------------------------------------------------------- main
def main() -> int:
    args = sys.argv[1:]
    quick = "--quick" in args
    strict = "--strict" in args
    deep = "--deep" in args
    no_network = "--no-network" in args
    only_line = None
    if "--line" in args:
        try:
            only_line = args[args.index("--line") + 1]
        except IndexError:
            only_line = None

    hard_fail = 0
    if deep:
        print("（--deep 会打开前端枚举窗体/模块/宏/报表，将改写前端字节；跑完请执行 "
              "check_upstream_drift.py --register 重算基线）")

    print("== 线① 上游版本/sha（提示级" + ("，--strict 计失败" if strict else "") + "）==")
    if only_line in (None, "1"):
        l1, drift = line1_upstream(quick, no_network)
        print("\n".join(l1) or "    （无跟踪路）")
        if drift and strict:
            hard_fail += 1
        if drift:
            print("    ↳ 提示：上游有更新，同步后需跑线②复核锚点")
    else:
        print("    （跳过）")

    print("== 线② 事实锚点（硬漂移，STALE 必修）==")
    if only_line in (None, "2"):
        l2, stale, skipped = line2_facts(deep)
        print("\n".join(l2))
        hard_fail += stale
        print("    ↳ %d 条 STALE / %d 条 SKIP（%s）" % (stale, skipped, "加 --deep 核验窗体/模块/宏/报表" if skipped else "全部核验"))
    else:
        print("    （跳过）")

    print("== 线③ 版本一致性（硬）==")
    if only_line in (None, "3"):
        l3, broken = line3_consistency()
        print("\n".join(l3))
        hard_fail += broken
    else:
        print("    （跳过）")

    print()
    print("结论: " + ("FAIL（硬漂移，先修再引用）" if hard_fail else "PASS（可安全引用自研文档中的上游断言）"))
    return 1 if hard_fail else 0


if __name__ == "__main__":
    sys.exit(main())
