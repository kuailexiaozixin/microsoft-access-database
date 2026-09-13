# -*- coding: utf-8 -*-
"""阶梯探针：在主线程逐个调用 VCS 加载项的 API，定位「哪一类调用会失败」。

背景
----
`Application.Run("<全路径>.API", 命令, 参数...)` 的加载项来源：
1. 技能目录运行位 `Microsoft Access Version Control System\Version Control.accda`（安装位=运行位，自包含）；
2. （历史 `%AppData%\MSAccessVCS` 已回收清理，不再使用。）

本探针会：
- 打开宿主库（加载项必须有库打开才可解析）；
- 依次调用 0 参 / 1 参 / 2 参 的命令；
- 每次调用后用 `app.Version` 检查 Access 是否还活着；
- 打印每个命令的结果或异常。

用法: python probe_vcs_api.py <宿主库.accdb> [--noguard]
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADDIN_SKILL = os.path.join(_SKILL_DIR, "Microsoft Access Version Control System", "Version Control.accda")


def ladder():
    # (命令, 参数元组)  —— 从无副作用到有副作用，逐步加压
    return [
        ("GetVCSVersion", ()),
        ("IsDatabaseOpen", ()),
        ("GetProjectName", ()),
        ("GetExportFolder", ()),
        ("IsVBACompiled", ()),
        ("GetOption", ("ExportFolder",)),
        ("SetOption", ("ShowDebug", False)),
        ("GetLogContent", ("Build",)),
        ("ExportObject", ("modules", "basTestSuite")),
        ("CompileVBA", ()),
        ("GetVCSVersion", ()),                 # 复查存活
    ]


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    db = os.path.abspath(sys.argv[1])
    addin = ADDIN_SKILL
    noguard = "--noguard" in sys.argv
    base = os.path.dirname(db)
    print("[加载项] %s" % addin)
    print("[存在] %s" % os.path.isfile(addin))
    print("[守卫] %s" % ("关闭" if noguard else "开启"))

    if not noguard:
        fwguard.start_guard()
    fwguard.kill_access(base)
    time.sleep(2)
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    try:
        app.OpenCurrentDatabase(db)
        print("[库] %s" % db)
        api = os.path.splitext(addin)[0] + ".API"

        def alive():
            try:
                return app.Version
            except Exception as e:
                return "死:%s" % str(e)[:60]

        for cmd, args in ladder():
            print("--- %s(%s) ---" % (cmd, ", ".join(map(str, args))))
            t0 = time.time()
            try:
                r = app.Run(api, cmd, *args) if args else app.Run(api, cmd)
                if isinstance(r, tuple):
                    r = r[0]
                s = str(r)
                print("    OK %.1fs  ->  %s" % (time.time() - t0, s[:200]))
            except Exception as e:
                print("    失败 %.1fs  ->  %s" % (time.time() - t0, str(e)[:200]))
            print("    存活: %s" % alive())
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
        if not noguard:
            fwguard.stop_guard()
            for s in fwguard.dialogs_seen()[:8]:
                print("[弹窗] %r 正文=%s" % (s[0], (s[1][:120] if s[1] else "")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
