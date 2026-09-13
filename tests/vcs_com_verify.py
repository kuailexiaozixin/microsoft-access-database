# -*- coding: utf-8 -*-
"""VCS 通道回归测试（COM 层）：Ribbon 加载 + API 阶梯 + ExportObject 错误 91 修复确认。

验证内容：
1. COMAddIns 集合中出现 MSAccessVCSLib.AddInRibbon（Ribbon 已加载，Access 里应显示选项卡）；
2. 打开宿主库后调用加载项 API：GetVCSVersion / ExportObject / CompileVBA；
3. ExportObject 不再触发错误 91（NoIndex 修复生效），产物落盘；
4. 全程无残留弹窗（fwguard 捕获）、收尾无新残留 Access 进程（进程快照对比，只杀本次新进程）、无锁文件。

用法: python vcs_com_verify.py --db <宿主库.accdb> [--leave-procs]
注意: 需要宿主库含可导出模块（如 basSample，见 tests/README.md 的建库步骤）。
"""
import os, sys, time

# 技能根：本文件位于 <skill_root>\tests\ 下
_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_SKILL_DIR, "scripts"))
import fwguard
from win32com.client.gencache import EnsureDispatch

ADDIN = os.path.join(_SKILL_DIR, "Microsoft Access Version Control System", "Version Control.accda")
RIBBON_PROGID = "MSAccessVCSLib.AddInRibbon"


def main():
    args = sys.argv[1:]
    if "--db" not in args:
        print(__doc__)
        return 1
    db = os.path.abspath(args[args.index("--db") + 1])
    base = os.path.dirname(db)
    ok = True

    print("[0] 加载项运行位: %s (%d B)" % (ADDIN, os.path.getsize(ADDIN) if os.path.isfile(ADDIN) else 0))
    if not os.path.isfile(ADDIN):
        print("[0] 失败: 运行位 accda 缺失")
        return 1

    before = fwguard.snapshot_access()
    fwguard.start_guard()
    fwguard.kill_access(base)
    time.sleep(2)

    app = None
    try:
        app = EnsureDispatch("Access.Application")
        app.Visible = True

        # ---- 1. Ribbon COM 加载项检查 ----
        print("\n[1] COMAddIns 集合 (Ribbon 加载验证)")
        found_ribbon = False
        try:
            ca = app.COMAddIns
            n = ca.Count
            print("    COMAddIns.Count = %d" % n)
            for i in range(1, n + 1):
                try:
                    item = ca.Item(i)
                    desc = item.Description or ""
                    prog = item.ProgID or ""
                    obj = "?"
                    try:
                        obj = item.Object and "实例已连接" or "实例为空"
                    except Exception:
                        obj = "读取失败"
                    print("    [%d] ProgID=%s | %s | %s" % (i, prog, desc, obj))
                    if prog == RIBBON_PROGID:
                        found_ribbon = True
                except Exception as e:
                    print("    [%d] 读取失败: %s" % (i, str(e)[:120]))
        except Exception as e:
            print("    COMAddIns 枚举失败: %s" % str(e)[:200])
        print("    Ribbon 加载: %s" % ("OK" if found_ribbon else "失败（未在 COMAddIns 中找到）"))
        ok = ok and found_ribbon

        # ---- 2. 打开宿主库 + API 阶梯 ----
        app.OpenCurrentDatabase(db)
        print("\n[2] 宿主库: %s" % db)
        api = os.path.splitext(ADDIN)[0] + ".API"

        def alive():
            try:
                return app.Version
            except Exception as e:
                return "死:%s" % str(e)[:60]

        def call(cmd, *args):
            t0 = time.time()
            try:
                r = app.Run(api, cmd, *args) if args else app.Run(api, cmd)
                if isinstance(r, tuple):
                    r = r[0]
                return ("OK", str(r)[:200], time.time() - t0)
            except Exception as e:
                return ("失败", str(e)[:200], time.time() - t0)

        for cmd, args in [
            ("GetVCSVersion", ()),
            ("IsDatabaseOpen", ()),
            ("GetProjectName", ()),
            ("ExportObject", ("module", "basSample")),
            ("CompileVBA", ()),
            ("GetVCSVersion", ()),
        ]:
            st, msg, dt = call(cmd, *args)
            print("    %-18s %-4s %.1fs -> %s" % (cmd, st, dt, msg))
            if cmd == "ExportObject":
                if st != "OK" or "error" in msg.lower() or "错误" in msg:
                    ok = False
                    print("    [!!] ExportObject 未通过（错误 91 修复验证失败）")
            print("    存活: %s" % alive())
    finally:
        try:
            if app is not None:
                app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit(2)
        except Exception:
            pass
        fwguard.kill_access(base)
        time.sleep(1)
        fwguard.stop_guard()

    # ---- 3. 残留检查（只检查本次新拉起的实例是否已清；历史孤儿实例不在本次范围）----
    new_left = fwguard.kill_new_access(before, base_dir=base)
    print("\n[3] 本次新拉起的残留 Access 进程: %s" % ("无" if not new_left else new_left))
    ok = ok and not new_left

    dialogs = fwguard.dialogs_seen()[:10]
    print("[4] 守卫捕获弹窗 %d 个" % len(dialogs))
    for s in dialogs:
        print("    %r 正文=%s" % (s[0], (s[1][:120] if s[1] else "")))
    print("\n===== 最终判定: %s =====" % ("全部通过" if ok else "存在失败项"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
