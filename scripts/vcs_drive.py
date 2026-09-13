# -*- coding: utf-8 -*-
"""在主线程上直接驱动 Microsoft Access Version Control System 加载项的 API。

为什么需要它
------------
`msaccess-vcs-mcp` 的 CLI 走的是「异步回调 + 后台工作线程」通道：
- 导出：add-in 回「已启动」但完成回调没回来 → 600 秒超时（实测）；
- 构建：`load_addin(db_path=None)` 报 `RPC_E_WRONG_THREAD`
  （-2147417842「应用程序调用一个已为另一线程整理的接口」）。
这两个都是**工具调用层的线程/回调问题**，不是被操纵的 Access 库本身的问题。
要验证"库能不能被完全操纵"，就在主线程上用同一个 add-in 引擎直接调它的 API。

用法
----
    python vcs_drive.py GetVCSVersion
    python vcs_drive.py Build "<源码目录>"
    python vcs_drive.py MergeBuild --db "<库.accdb>"
    python vcs_drive.py ExportByType "modules" True --db "<库.accdb>"
    python vcs_drive.py CompileVBA --db "<库.accdb>"

约定
----
- 默认使用技能目录运行位 `Microsoft Access Version Control System\Version Control.accda`
  （安装位=运行位，自包含；历史 `%AppData%\MSAccessVCS` 已回收清理，不再使用）。
- `--db <路径>` 表示先打开该库，把 add-in 加载起来；此时会先用该库调
  `GetVCSVersion` 做探针。
- `Build` 额外要求「当前库与源目录推导出的目标库同名同路径」，否则 add-in
  会弹 MsgBox（模态阻塞 → RPC 中断）；脚本会提前校验并拒绝。
- 传入 `True`/`False` 会转成布尔值；纯数字转 int；其余按字符串。
- `--then-run <过程名>`：命令跑完在宿主库里调用该过程并打印结果（用于复核回写）。
- `--settle <秒>`：命令返回后先等这么久再收尾（Build/MergeBuild 的重活可能异步派发）。
- 全程装弹窗守卫，收尾清进程与锁。
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 运行位=安装位：技能目录 Microsoft Access Version Control System\Version Control.accda
ADDIN = os.path.join(_SKILL_DIR, "Microsoft Access Version Control System", "Version Control.accda")
# 兼容：--skill-addin 曾用于切换技能副本；现运行位即技能目录，该参数保留仅作兼容（等价默认）。


def coerce(tok):
    if tok == "True":
        return True
    if tok == "False":
        return False
    if tok.lstrip("-").isdigit():
        return int(tok)
    return tok


def main():
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        return 1
    addin = ADDIN
    if "--skill-addin" in argv:
        # 兼容旧用法：运行位即技能目录，无需切换
        argv.remove("--skill-addin")
    then_run = None
    if "--then-run" in argv:
        i = argv.index("--then-run")
        then_run = argv[i + 1]
        del argv[i:i + 2]
    settle = 0
    if "--settle" in argv:
        i = argv.index("--settle")
        settle = int(argv[i + 1])
        del argv[i:i + 2]
    db = None
    if "--db" in argv:
        i = argv.index("--db")
        db = os.path.abspath(argv[i + 1])
        del argv[i:i + 2]
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 1
    command = argv[0]
    args = [coerce(a) for a in argv[1:]]

    base = os.path.dirname(db) if db else os.path.dirname(os.path.abspath(args[0])) if args and os.path.isdir(str(args[0])) else os.getcwd()
    fwguard.start_guard()
    fwguard.kill_access(base)
    time.sleep(2)                       # 等上一批 MSACCESS 真正退干净
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    print("[加载项] %s" % addin)
    try:
        # 就绪轮询：Access 刚起来时立刻调用会报「远程过程调用失败 / RPC 服务器不可用」
        ready = False
        for i in range(30):
            try:
                _ = app.Version
                ready = True
                break
            except Exception as e:
                last = str(e)[:80]
                time.sleep(1)
        print("[就绪] %s（Access %s）" % ("是" if ready else "否（超时）", getattr(app, "Version", "?")))
        if not ready:
            print("[致命] Access 未就绪: %s" % last)
            return 3

        api = os.path.splitext(os.path.abspath(addin))[0] + ".API"
        ver = None
        if db:
            app.OpenCurrentDatabase(db)
            print("[库] %s" % db)
            for attempt in range(1, 4):
                try:
                    ver = app.Run(api, "GetVCSVersion")
                    if isinstance(ver, tuple):
                        ver = ver[0]
                    break
                except Exception as e:
                    print("[加载 add-in 第 %d 次失败] %s" % (attempt, str(e)[:160]))
                    time.sleep(3)
        print("[add-in] 版本=%s" % ver)
        if ver is None:
            print("[致命] add-in 未加载成功")
            return 2

        # 完全构建（Build）时，add-in 会校验「当前打开的库」与「源目录推导出的
        # 目标库」是否是同一个文件名：
        #   - 文件名一致 → 直接进入完全构建（它会自己关库、重建、再导入）；
        #   - 文件名不一致 → 弹 MsgBox 询问，模态阻塞 → 调用被 RPC 中断；
        #   - 当前没有库打开 → Access 解析不了 `<路径>.API` 库引用，同样 RPC 失败。
        # 所以 Build 的宿主库必须与目标库同名同路径。
        if command == "Build" and db and args and os.path.isdir(str(args[0])):
            import json
            pj = os.path.join(str(args[0]), "vbe-project.json")
            try:
                with open(pj, encoding="utf-8-sig") as f:
                    tgt = json.load(f)["Items"]["FileName"]
                want = os.path.abspath(os.path.join(os.path.dirname(str(args[0])), tgt))
                if os.path.normcase(os.path.normpath(want)) != os.path.normcase(os.path.normpath(db)):
                    print("[致命] 宿主库与目标库不同名，add-in 会弹模态对话框：")
                    print("        期望宿主库 = %s" % want)
                    print("        实际宿主库 = %s" % db)
                    return 4
                print("[校验] 宿主库=目标库=%s" % want)
            except Exception as e:
                print("[警告] 无法校验宿主库名: %s" % str(e)[:120])

        t0 = time.time()
        print("[调用] %s(%s)" % (command, ", ".join(repr(a) for a in args)))
        r = app.Run(api, command, *args) if args else app.Run(api, command)
        if isinstance(r, tuple):
            r = r[0]
        print("[用时] %.1f 秒" % (time.time() - t0))
        print("[结果] %s" % (r,))

        # 某些命令（Build/MergeBuild）会经窗体把重活派发出去，外层调用很快返回。
        # 用 --settle N 在收尾（关库/杀进程）前多等一会儿，避免把没干完的活掐死。
        if settle > 0:
            print("[等待] %d 秒，让加载项把活干完…" % settle)
            deadline = time.time() + settle
            while time.time() < deadline:
                time.sleep(3)
                print("   …剩余 %ds" % int(deadline - time.time()))

        # 可选：命令跑完后，在宿主库里调用一个过程验证效果
        if then_run:
            try:
                v = app.Run(then_run)
                if isinstance(v, tuple):
                    v = v[0]
                print("[复核] %s() = %s" % (then_run, v))
            except Exception as e:
                print("[复核] %s() 失败: %s" % (then_run, str(e)[:200]))
        return 0
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
        seen = fwguard.dialogs_seen()
        if seen:
            print("[守卫记录到 %d 个弹窗]" % len(seen))
            for s in seen[:5]:
                print("   %r 正文=%s 按钮=%s" % (s[0], s[1][:100] if s[1] else "", s[1][1:]))


if __name__ == "__main__":
    raise SystemExit(main())
