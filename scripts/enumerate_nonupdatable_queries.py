# -*- coding: utf-8 -*-
"""enumerate_nonupdatable_queries.py —— 定位错误 #3027（记录集只读）的根因。

错误 #3027「不能更新。数据库或对象为只读。」常见三类根因：
  1) 文件/链接层：.accdb 只读、.laccdb 残留锁、后端只读。  -> 本脚本只做逻辑层静态扫描。
  2) 逻辑非可更新记录集（最常见）：RecordSource 是含 DISTINCT/GROUP BY/UNION/聚合、
     或多表 JOIN 且未含各基表主键的查询。
  3) 设计层：窗体 RecordsetType=Snapshot，或控件绑表达式。

本脚本静态扫描：
  - 每个数据库的 QueryDefs：标出疑似非可更新查询（DISTINCT/GROUP BY/UNION/聚合/JOIN）。
  - 每个窗体的 RecordSource + RecordsetType，标注绑定到疑似只读查询的窗体。

用法：python scripts/enumerate_nonupdatable_queries.py [--db 库路径 ...]
      --db 可多次指定；缺省扫描本技能自带的 Edonsoft 后端库（Edonsoft Development Framework_x64/Data.accdb）。
"""
import os, sys, time, subprocess, re, argparse
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)
# 缺省扫描目标：本技能 12 个子项目中 Edonsoft 框架的后端数据文件（通用、可替换为任意 .accdb/.mdb）
DEFAULT_DBS = [os.path.join(SKILL_ROOT, "Edonsoft Development Framework_x64", "Data.accdb")]

AGGS = ("SUM(", "COUNT(", "AVG(", "MIN(", "MAX(", "FIRST(", "LAST(", "STDEV(", "VAR(")
JOIN_RE = re.compile(r"\bJOIN\b", re.I)
FROM_RE = re.compile(r"\bFROM\b\s+([`\"\[\w.]+)", re.I)


def scan_sql(name, sql):
    flags = []
    s = sql or ""
    up = s.upper()
    if "DISTINCT" in up: flags.append("DISTINCT")
    if "GROUP BY" in up: flags.append("GROUP BY")
    if " UNION " in (" " + up + " "): flags.append("UNION")
    if any(a in up for a in AGGS): flags.append("聚合函数")
    joins = len(JOIN_RE.findall(s))
    if joins:
        # 多表 JOIN：统计 FROM 后的表数量
        froms = FROM_RE.findall(s)
        tables = set()
        for f in froms:
            tables.add(f.split(".")[-1].strip("[]`\""))
        # JOIN 关键字数 + 主 FROM 表
        n_tables = len(tables) + joins
        if n_tables >= 2:
            flags.append("多表JOIN(%d)" % n_tables)
    return flags


def main():
    ap = argparse.ArgumentParser(
        description="静态扫描 .accdb/.mdb 中疑似非可更新的查询与窗体（定位错误 #3027）")
    ap.add_argument("--db", action="append", default=[], metavar="PATH",
                    help="要扫描的数据库文件路径，可多次指定；缺省扫描 Edonsoft 后端库")
    args = ap.parse_args()
    dbs = args.db if args.db else DEFAULT_DBS
    for db_path in dbs:
        if not os.path.exists(db_path):
            print("跳过（文件不存在）: %s" % db_path)
            continue
        print("\n========== %s ==========" % os.path.basename(db_path))
        app = win32.gencache.EnsureDispatch("Access.Application")
        app.AutomationSecurity = 3; app.Visible = False
        app.OpenCurrentDatabase(db_path); time.sleep(3)
        try:
            db = app.CurrentDb()
            # 查询
            susp = []
            qd = db.QueryDefs
            nq = qd.Count
            for i in range(nq):
                q = qd(i)
                if q.Name.startswith("~"):  # 自动生成行源，跳过
                    continue
                try:
                    sql = q.SQL
                except Exception:
                    sql = ""
                fl = scan_sql(q.Name, sql)
                if fl:
                    susp.append((q.Name, fl, sql[:120]))
            print("查询总数(剔除~):", nq - sum(1 for i in range(nq) if qd(i).Name.startswith("~")))
            print("疑似非可更新查询 %d 条：" % len(susp))
            for nm, fl, head in susp:
                print("  - %s  [%s]" % (nm, ", ".join(fl)))
                print("      SQL头: %s" % head.replace("\n", " "))
            # 窗体（仅主库有窗体）
            try:
                forms = app.CurrentProject.AllForms
                nf = forms.Count
                print("窗体总数:", nf)
                for i in range(nf):
                    fn = forms(i).Name
                    try:
                        frm = app.Forms(fn)  # 仅已打开
                        rs = frm.RecordSource; rt = frm.RecordsetType
                    except Exception:
                        # 未打开则跳过（需设计视图，静态扫描从 .src 读更稳）
                        continue
                    fl = scan_sql(fn, rs)
                    print("  - 窗体 %s RecordSource=%r RecordsetType=%s %s" % (
                        fn, rs, rt, ("[疑似只读]" if fl else "")))
            except Exception as e:
                print("窗体扫描跳过:", e)
        finally:
            try: app.CloseCurrentDatabase()
            except Exception: pass
            try: app.Quit(2)
            except Exception: pass
            time.sleep(1)
            subprocess.run("taskkill /F /IM MSACCESS.EXE", shell=True, capture_output=True)
            time.sleep(1)
            lf = db_path.replace(".accdb", ".laccdb")
            try: os.remove(lf)
            except Exception: pass


if __name__ == "__main__":
    main()
