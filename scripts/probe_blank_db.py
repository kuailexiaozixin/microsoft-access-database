# -*- coding: utf-8 -*-
"""对照实验：在一个全新空白库里调用加载项 API，判断崩溃是「环境/加载项」问题还是「特定库」问题。

用法: python probe_blank_db.py <新建空白库路径>
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

# 默认用【运行位=安装位】的那份：技能目录 Microsoft Access Version Control System\。
# （历史 %AppData%\MSAccessVCS 已回收清理，不再使用。）
_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADDIN = os.path.join(_SKILL_DIR, "Microsoft Access Version Control System", "Version Control.accda")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    db = os.path.abspath(sys.argv[1])
    addin = ADDIN
    base = os.path.dirname(db)
    fwguard.start_guard()
    fwguard.kill_access(base)
    time.sleep(2)
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    print("[加载项] %s" % addin)
    try:
        # 用 Application.NewCurrentDatabase 建空白库（与加载项自己的建库方式一致）。
        # 不要用 DBEngine.CreateDatabase：COM 晚绑定下会报「找不到可安装的 ISAM」。
        if os.path.exists(db):
            os.remove(db)
        app.NewCurrentDatabase(db)
        print("[空白库] %s" % db)
        api = os.path.splitext(addin)[0] + ".API"

        def alive():
            try:
                return app.Version
            except Exception as e:
                return "死:%s" % str(e)[:50]

        for cmd, args in [
            ("GetVCSVersion", ()),
            ("IsDatabaseOpen", ()),
            ("GetExportFolder", ()),
            ("GetVCSVersion", ()),
        ]:
            t0 = time.time()
            try:
                r = app.Run(api, cmd, *args) if args else app.Run(api, cmd)
                if isinstance(r, tuple):
                    r = r[0]
                print("  %-18s OK   %.1fs -> %s" % (cmd, time.time() - t0, str(r)[:80]))
            except Exception as e:
                print("  %-18s 失败 %.1fs -> %s" % (cmd, time.time() - t0, str(e)[:110]))
            print("    存活: %s" % alive())
            if str(alive()).startswith("死"):
                break
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
