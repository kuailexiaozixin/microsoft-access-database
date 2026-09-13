# -*- coding: utf-8 -*-
r"""构建 → 修复门禁 → 安装 → 探针，一键重建并安装 VCS 加载项到技能目录运行位。

流程（每步失败即退出并给出原因）：
1. 【修复门禁】校验源码中错误 91 的 NoIndex 保护仍在（4 个组件含 `If Not cIdx Is Nothing`），
   防止未来从上游重克隆/重构建时修复回退；
2. 【构建】`vcs_drive.py Build` 从源码构建 accda（Build 要求宿主库与源目录同名同路径，
   脚本自动准备同名空宿主库）；
3. 【安装】`install_vcs_addin.py` 代码化安装（/cmd INSTALL + Ribbon 部署 + 六项自检）；
4. 【探针】`probe_vcs_api.py` 阶梯探针 + ExportObject 复核（错误 91 回归门禁）。

用法:
    python rebuild_and_install.py [--src-dir <msaccess-vcs-addin\Version Control.accda.src>]
                                  [--out <msaccess-vcs-addin\Version Control.accda>]
                                  [--folder <Microsoft Access Version Control System>]
                                  [--ribbon-dll <msaccess-vcs-addin\Ribbon\Build\MSAccessVCSLib_win64.dll>]
                                  [--ribbon-xml <msaccess-vcs-addin\Ribbon\Ribbon.xml>]
                                  [--ribbon-json <运行位 Ribbon.json>]
                                  [--workdir <临时工作目录>]

说明: 构建用宿主库在 --workdir 下自动创建（<out 文件名>.host.accdb），构建完成后删除；
      运行位与 MCP .env 指向不变，本次仅重建并覆盖运行位 accda。
"""
import os, shutil, subprocess, sys, time

_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(_SKILL_DIR, "scripts")
PY = sys.executable

# 错误 91 NoIndex 修复门禁：4 个组件必须包含保护行
FIX_FILES = [
    "clsDbForm.cls",
    "clsDbModule.cls",
    "clsDbReport.cls",
    "clsDbVbeForm.cls",
]
FIX_MARK = "If Not cIdx Is Nothing Then"
FIX_COMPONENTS_DIR = os.path.join("modules", "Components")


def run(cmd, **kw):
    print("\n>>> %s" % " ".join(cmd))
    r = subprocess.run(cmd, **kw)
    if r.returncode != 0:
        print("!!! 命令失败，退出码 %d" % r.returncode)
        raise SystemExit(r.returncode)
    return r


def check_fix_gate(src_dir):
    print("\n== [1] 修复门禁：NoIndex 保护（错误 91）==")
    missing = []
    for fn in FIX_FILES:
        p = os.path.join(src_dir, FIX_COMPONENTS_DIR, fn)
        if not os.path.isfile(p):
            missing.append(fn + "(文件缺失)")
            continue
        text = open(p, "r", encoding="utf-8-sig", errors="replace").read()
        if FIX_MARK not in text:
            missing.append(fn + "(无保护行)")
        else:
            print("    OK  %s 含 %s" % (fn, FIX_MARK))
    if missing:
        print("!!! 修复门禁未通过: %s" % "; ".join(missing))
        print("    说明: 错误 91 保护缺失，请先检查源码（修复方法见 references/vcs-three-channels.md 第十三节）")
        raise SystemExit(1)


def main():
    argv = sys.argv[1:]
    def get(name, default):
        if name in argv:
            return os.path.abspath(argv[argv.index(name) + 1])
        return default

    src_dir = get("--src-dir", os.path.join(_SKILL_DIR, "msaccess-vcs-addin", "Version Control.accda.src"))
    out_db = get("--out", os.path.join(_SKILL_DIR, "msaccess-vcs-addin", "Version Control.accda"))
    folder = get("--folder", os.path.join(_SKILL_DIR, "Microsoft Access Version Control System"))
    ribbon_dll = get("--ribbon-dll", os.path.join(_SKILL_DIR, "msaccess-vcs-addin", "Ribbon", "Build", "MSAccessVCSLib_win64.dll"))
    ribbon_xml = get("--ribbon-xml", os.path.join(_SKILL_DIR, "msaccess-vcs-addin", "Ribbon", "Ribbon.xml"))
    ribbon_json = get("--ribbon-json", os.path.join(folder, "Ribbon.json"))
    workdir = get("--workdir", os.path.join(_SKILL_DIR, "_build_tmp"))

    for name, p in [("源码目录", src_dir), ("Ribbon DLL", ribbon_dll), ("Ribbon XML", ribbon_xml)]:
        if not os.path.exists(p):
            print("!!! 缺失: %s=%s" % (name, p))
            return 1

    # 1. 修复门禁
    check_fix_gate(src_dir)

    # 2. 构建（vcs_drive Build 要求宿主库与目标同名同路径）
    print("\n== [2] 构建 accda ==")
    os.makedirs(workdir, exist_ok=True)
    host_db = os.path.join(workdir, os.path.basename(out_db) + ".host.accdb")
    if os.path.exists(host_db):
        os.remove(host_db)
    try:
        # 先建同名空宿主库（Build 会自建备份/重建，但要求当前有库打开且同名）
        import pythoncom
        from win32com.client.gencache import EnsureDispatch
        pythoncom.CoInitialize()
        app = EnsureDispatch("Access.Application")
        app.Visible = False
        app.NewCurrentDatabase(host_db)
        app.Quit(2)
        time.sleep(2)
    except Exception as e:
        print("!!! 创建宿主库失败: %s" % str(e)[:200])
        return 1
    run([PY, os.path.join(SCRIPTS, "vcs_drive.py"), "Build", src_dir, "--db", host_db, "--settle", "60"],
        cwd=SCRIPTS)
    # 构建产物为源目录推导出的目标（vbe-project.json FileName 决定）；此处要求即 out_db
    if not os.path.isfile(out_db) or os.path.getsize(out_db) < 1_000_000:
        print("!!! 构建产物异常: %s (%d B)" % (out_db, os.path.getsize(out_db) if os.path.exists(out_db) else 0))
        return 1
    print("    构建产物: %s (%d B)" % (out_db, os.path.getsize(out_db)))

    # 3. 代码化安装（/cmd INSTALL + Ribbon 部署 + 六项自检）
    print("\n== [3] 代码化安装到运行位 ==")
    run([PY, os.path.join(SCRIPTS, "install_vcs_addin.py"),
         "--source", out_db,
         "--folder", folder,
         "--ribbon-dll", ribbon_dll,
         "--ribbon-xml", ribbon_xml,
         "--ribbon-json", ribbon_json],
        cwd=SCRIPTS)

    # 4. 探针复核（错误 91 回归门禁）：建含 basSample 夹具的探针库后跑 COM 回归
    print("\n== [4] 探针复核（GetVCSVersion + ExportObject）==")
    probe_db = os.path.join(workdir, "_probe_host.accdb")
    if os.path.exists(probe_db):
        os.remove(probe_db)
    try:
        import pythoncom
        from win32com.client.gencache import EnsureDispatch
        pythoncom.CoInitialize()
        app = EnsureDispatch("Access.Application")
        app.Visible = False
        app.NewCurrentDatabase(probe_db)
        app.Quit(2)
        time.sleep(2)
    except Exception as e:
        print("!!! 创建探针库失败: %s" % str(e)[:200])
        return 1
    fixture = os.path.join(_SKILL_DIR, "tests", "fixtures", "basSample.bas")
    run([PY, os.path.join(SCRIPTS, "rebuild_module_from_src.py"), probe_db, fixture])
    run([PY, os.path.join(_SKILL_DIR, "tests", "vcs_com_verify.py"), "--db", probe_db])

    # 清理临时宿主库与探针库
    for p in (host_db, probe_db):
        try:
            os.remove(p)
        except Exception:
            pass
    print("\n===== 重建并安装完成: 全部通过 =====")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
