# -*- coding: utf-8 -*-
"""VCS 一致性核验：实时库对象 vs .src 真相源，找出任何"一边有一边无"的漂移。

用法:
    python vcs_consistency_check.py <主库.accdb> [<对应 .src 目录>]

输出四张清单 + 漂移报告。退出码 0=无漂移，3=有漂移。
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

# vbext_ct_StdModule=1, vbext_ct_ClassModule=2, vbext_ct_MSForm=3, vbext_ct_Document=100
TYPE_NAME = {1: "标准模块", 2: "类模块", 3: "窗体类", 100: "文档类(窗体/报表代码)"}


def live_objects(app, db):
    proj = fwguard.get_proj(app, db)
    if proj is None:
        return None
    comps = [(c.Name, c.Type, c.CodeModule.CountOfLines) for c in proj.VBComponents]
    cdb = app.CurrentDb()
    tables = sorted(t.Name for t in cdb.TableDefs if not t.Name.startswith("MSys"))
    queries = sorted(q.Name for q in cdb.QueryDefs if not q.Name.startswith("~"))
    forms = sorted(app.CurrentProject.AllForms(i).Name for i in range(app.CurrentProject.AllForms.Count))
    reports = sorted(app.CurrentProject.AllReports(i).Name for i in range(app.CurrentProject.AllReports.Count))
    macros = sorted(app.CurrentProject.AllMacros(i).Name for i in range(app.CurrentProject.AllMacros.Count))
    return proj.Name, comps, tables, queries, forms, reports, macros


def src_objects(src):
    def ls(sub, exts):
        d = os.path.join(src, sub)
        if not os.path.isdir(d):
            return []
        return sorted(os.path.splitext(f)[0] for f in os.listdir(d)
                      if os.path.splitext(f)[1].lower() in exts and not f.startswith("~"))
    return {
        "std_modules": ls("modules", {".bas"}),
        "cls_modules": ls("modules", {".cls"}),
        "forms": ls("forms", {".form"}),
        "macros": ls("macros", {".macro"}),
        "queries": ls("queries", {".sql"}),
        "tables": sorted(set(ls("tbldefs", {".json", ".sql"}))),
    }


def diff(label, live, srcv):
    only_live = sorted(set(live) - set(srcv))
    only_src = sorted(set(srcv) - set(live))
    print("\n[%s] 实时库 %d / .src %d" % (label, len(live), len(srcv)))
    if only_live:
        print("  ! 只在实时库：" + ", ".join(only_live))
    if only_src:
        print("  ! 只在 .src ：" + ", ".join(only_src))
    if not only_live and not only_src:
        print("  ok 完全一致")
    return only_live, only_src


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    db = os.path.abspath(sys.argv[1])
    src = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else (db + ".src")
    base = os.path.dirname(db)
    if not os.path.isdir(src):
        print("[致命] 找不到 .src 目录: %s" % src)
        return 2

    fwguard.start_guard()
    fwguard.kill_access(base)
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    app.OpenCurrentDatabase(db)
    try:
        got = live_objects(app, db)
        if got is None:
            print("[致命] 按 FileName 定位不到目标工程")
            return 2
        pname, comps, tables, queries, forms, reports, macros = got
        print("[工程] %s" % pname)
        print("[实时库] 组件 %d / 表 %d / 查询 %d / 窗体 %d / 报表 %d / 宏 %d"
              % (len(comps), len(tables), len(queries), len(forms), len(reports), len(macros)))

        std = [n for n, t, _ in comps if t == 1]
        cls = [n for n, t, _ in comps if t == 2]
        doc = [n for n, t, _ in comps if t in (3, 100)]
        print("[类型明细] 标准模块 %d / 类模块 %d / 文档类 %d" % (len(std), len(cls), len(doc)))
        for n, t, ln in comps:
            if t == 2:
                print("    类模块 %-16s type=%d 行数=%d" % (n, t, ln))

        s = src_objects(src)
        drift = 0
        pairs = (
            ("标准模块", std, s["std_modules"]),
            ("类模块", cls, s["cls_modules"]),
            ("窗体", forms, s["forms"]),
            ("宏", macros, s["macros"]),
            ("查询", queries, s["queries"]),
            ("表", tables, s["tables"]),
        )
        for label, live_list, src_list in pairs:
            ol, only_src = diff(label, live_list, src_list)
            drift += len(ol) + len(only_src)

        print("\n===== 漂移合计 %d 项 =====" % drift)
        print("结论：%s" % ("无漂移，实时库与 .src 对齐" if drift == 0 else "存在漂移，需处理"))
        return 0 if drift == 0 else 3
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        fwguard.kill_access(base)
        fwguard.stop_guard()


if __name__ == "__main__":
    raise SystemExit(main())
