# -*- coding: utf-8 -*-
"""probe_vba_project.py —— 只读探针：报告目标库 VBA 工程真实状态。

输出：组件总数、类模块(type=2)数与名称、标准模块数、是否有模块含未定义的条件编译常量
（含 #If 引用未定义常量的模块会在关库时被静默丢弃，是本技能重点排查对象）。

用法：python scripts/probe_vba_project.py <目标库.accdb>
"""
import os, sys, time, subprocess
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
if len(sys.argv) > 1:
    MAIN = os.path.abspath(sys.argv[1])
else:
    print("用法: python scripts/probe_vba_project.py <目标库.accdb>")
    raise SystemExit(2)


def main():
    app = win32.gencache.EnsureDispatch("Access.Application")
    app.AutomationSecurity = 3
    app.Visible = False
    app.OpenCurrentDatabase(MAIN)
    time.sleep(4)
    try:
        # 铁律：ActiveVBProject 常指向 VCS 加载项，按 FileName 定位目标工程
        proj = None
        for p in app.VBE.VBProjects:
            try:
                fn = p.FileName
            except Exception:
                fn = ""
            if fn and os.path.normcase(fn) == os.path.normcase(MAIN):
                proj = p
                break
        if proj is None:
            proj = app.VBE.ActiveVBProject
        print("工程名:", proj.Name)
        comps = list(proj.VBComponents)
        print("组件总数:", len(comps))
        class_mods = [c.Name for c in comps if c.Type == 2]
        print("类模块(type=2)数:", len(class_mods), class_mods)
        print("标准模块(type=1)数:", sum(1 for c in comps if c.Type == 1))
        # 重点：含 #If 条件编译的模块（引用未定义常量会静默丢弃）
        risky = []
        for c in comps:
            if c.Type in (1, 2):
                try:
                    txt = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
                except Exception:
                    continue
                ifs = [l for l in txt.split('\r\n') if l.strip().startswith('#If')]
                if ifs:
                    risky.append((c.Name, ifs))
        print("含 #If 条件编译的模块数:", len(risky))
        for name, ifs in risky:
            print("  - %s : %s" % (name, ifs))
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        time.sleep(1)
        subprocess.run("taskkill /F /IM MSACCESS.EXE", shell=True, capture_output=True)
        time.sleep(1)
        lf = MAIN.replace(".accdb", ".laccdb")
        try:
            os.remove(lf)
        except Exception:
            pass


if __name__ == "__main__":
    main()
