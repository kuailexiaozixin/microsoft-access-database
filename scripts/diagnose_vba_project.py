# -*- coding: utf-8 -*-
"""diagnose_vba_project.py —— 开工第一步：探测 VBA 工程「能力矩阵」。

每次动手改目标库之前先跑它，避免在一条死路上耗掉几十次调用。

报告内容：
    1. 组件概况（总数 / 标准模块 / 类模块 / 窗体）
    2. 引用清单（名称 + GUID）
    3. 能力矩阵：哪些操作可用、哪些被阻塞
       - 读代码 / ReplaceLine（改已有代码）
       - Add(1) / Add(2) / Remove / Import（结构性变更）
       - References 增删是否真的持久化
    4. 含 #If 条件编译的模块（静默丢弃风险）

用法：
    python scripts/diagnose_vba_project.py <目标库.accdb>

退出码始终为 0；结果看输出。
"""
import os, sys, time, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

OK, FAIL = "OK  ", "FAIL"

if len(sys.argv) > 1:
    MAIN = os.path.abspath(sys.argv[1])
else:
    print("用法: python scripts/diagnose_vba_project.py <目标库.accdb>")
    raise SystemExit(2)


def main():
    fwguard.start_guard()
    fwguard.kill_access(os.path.dirname(MAIN))
    app = EnsureDispatch("Access.Application")
    app.AutomationSecurity = 3
    app.Visible = False
    app.OpenCurrentDatabase(MAIN)
    try:
        # 铁律：ActiveVBProject 常指向 VCS 加载项(MSAccessVCS)，必须按 FileName 定位目标库
        proj = fwguard.get_proj(app, MAIN)
        if proj is None:
            print("!! 未找到目标工程（按 FileName 匹配失败），改用 ActiveVBProject 仅作参考")
            proj = app.VBE.ActiveVBProject
        print("目标工程:", proj.Name, "|", getattr(proj, "FileName", ""))

        comps = list(proj.VBComponents)
        n_std = sum(1 for c in comps if c.Type == 1)
        n_cls = sum(1 for c in comps if c.Type == 2)
        n_form = sum(1 for c in comps if c.Type == 100)
        print("=" * 60)
        print("工程:", proj.Name, " 组件总数:", len(comps))
        print(f"  标准模块(type=1): {n_std}   类模块(type=2): {n_cls}   窗体(type=100): {n_form}")

        print("\n--- 引用 ---")
        for r in proj.References:
            print(f"  {r.Name:<12} {r.GUID}")

        print("\n--- 能力矩阵 ---")
        # 1) 读代码
        try:
            c0 = [c for c in comps if c.Type == 1][0]
            _ = c0.CodeModule.Lines(1, 1)
            print(f"[{OK}] 读 CodeModule.Lines")
        except Exception as e:
            print(f"[{FAIL}] 读 CodeModule.Lines: {e}")

        # 2) 改已有代码（ReplaceLine）
        try:
            c0 = [c for c in comps if c.Type == 1][0]
            ln = 1
            while ln <= c0.CodeModule.CountOfLines and not c0.CodeModule.Lines(ln, 1).strip():
                ln += 1
            orig = c0.CodeModule.Lines(ln, 1)
            c0.CodeModule.ReplaceLine(ln, orig)  # 写回原值，无副作用
            print(f"[{OK}] ReplaceLine/改已有代码   <-- 阻塞时的唯一通道")
        except Exception as e:
            print(f"[{FAIL}] ReplaceLine: {e}")

        # 3) Add(1) / Add(2)
        for kind, label in ((1, "Add(1) 标准模块"), (2, "Add(2) 类模块")):
            try:
                c = proj.VBComponents.Add(kind)
                print(f"[{OK}] {label}")
                proj.VBComponents.Remove(c)
            except Exception as e:
                code = getattr(e, "args", ("",))[-1] if hasattr(e, "args") else ""
                print(f"[{FAIL}] {label}: {e}  (设备 I/O 错误 = 结构性变更被阻塞)")

        # 4) Import
        try:
            tmp = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "__diag_probe.bas")
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write('Attribute VB_Name = "__diag_probe"\r\nPublic Function D() As Long\r\n D=1\r\nEnd Function\r\n')
            c = proj.VBComponents.Import(tmp)
            print(f"[{OK}] Import  name={c.Name} type={c.Type}")
            proj.VBComponents.Remove(c)
            os.remove(tmp)
        except Exception as e:
            print(f"[{FAIL}] Import: {e}")

        # 5) 引用持久化
        try:
            before = sorted(r.Name for r in proj.References)
            proj.References.AddFromGuid("{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", 1, 0)
            after = sorted(r.Name for r in proj.References)
            if after == before:
                print(f"[{FAIL}] References.AddFromGuid 假成功：调用无异常但引用未持久化")
            else:
                print(f"[{OK}] References.AddFromGuid 已持久化 {set(after) - set(before)}")
        except Exception as e:
            print(f"[{FAIL}] References.AddFromGuid: {e}")

        # 6) 含 #If 条件编译的模块（静默丢弃风险）
        risky = []
        for c in comps:
            if c.Type in (1, 2):
                try:
                    txt = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
                except Exception:
                    continue
                ifs = [l for l in txt.split('\r\n') if l.strip().startswith('#If')]
                if ifs:
                    risky.append((c.Name, c.Type, ifs))
        print("\n--- 条件编译风险 ---")
        print(f"  含 #If 的模块: {len(risky)}")
        for name, t, ifs in risky:
            print(f"    {name} (type={t}): {ifs}")

        print("=" * 60)
        print("提示：若 Add/Remove/Import 全 FAIL 而 ReplaceLine OK，")
        print("      多半是定位错工程（ActiveVBProject 指向了加载项），详见 references/vba-project-operations.md")
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        fwguard.kill_access(os.path.dirname(MAIN))
        fwguard.stop_guard()


if __name__ == "__main__":
    main()
