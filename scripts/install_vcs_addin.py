# -*- coding: utf-8 -*-
"""代码化安装 VCS 加载项（还原官方安装程序的执行代码，全程无手动操作）。

原理
----
官方安装程序（Version Control.accda）的 AutoRun 支持命令行自动化：
    If UCase$(Trim$(Command$)) = "INSTALL" Then
        GetInstallSettings
        InstallVCSAddin this.blnTrustAddInFolder, this.blnUseRibbonAddIn, _
            False, this.strInstallFolder, this.blnUseCompiledAddIn
安装设置（Install Folder / Trust Folder / Use Ribbon / Compile accde / Open File）
保存在注册表 `HKCU\\Software\\VB and VBA Program Settings\\MSAccessVCS\\Install\\`。
InstallVCSAddin 会完成：复制 accda -> 目标文件夹、受信任位置、Menu Add-Ins 注册、
版本号写入，然后弹确认框并自动退出。本脚本：
1. 用 winreg 写入安装设置（Install Folder = 技能目录，Use Ribbon = 0）；
2. 启动弹窗守卫 fwguard（捕获消息框 + 自动点击按钮）；
3. 以 `/cmd INSTALL` 启动 MSACCESS.EXE 打开安装源，触发官方安装代码；
4. 等待安装进程退出；
5. 逐项验证安装结果并报告。

用法
----
    python install_vcs_addin.py --source "<安装源 accda>" --folder "<安装文件夹>"

示例（自包含安装到技能目录内）：
    python install_vcs_addin.py ^
        --source "...\\msaccess-vcs-addin\\Version Control.accda" ^
        --folder "...\\Microsoft Access Version Control System"
"""
import os, sys, time, subprocess, winreg
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard

PROJECT_NAME = "MSAccessVCS"
ADDIN_BASENAME = "Version Control"
INSTALL_KEY = r"Software\VB and VBA Program Settings\MSAccessVCS\Install"


def find_msaccess() -> str:
    """定位 MSACCESS.EXE（32/64 位 Office 常见安装路径 + 注册表兜底）。"""
    candidates = [
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\root\Office16\MSACCESS.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\MSACCESS.EXE"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\Office16\MSACCESS.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\Office16\MSACCESS.EXE"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    # 注册表兜底（Access.Application 的 Path）
    try:
        import win32com.client
        from win32com.client.gencache import EnsureDispatch
        app = EnsureDispatch("Access.Application")
        p = os.path.join(app.Path, "MSACCESS.EXE")
        app.Quit()
        if os.path.isfile(p):
            return p
    except Exception:
        pass
    raise SystemExit("[致命] 找不到 MSACCESS.EXE，请检查 Office 安装。")


def set_install_settings(install_folder: str):
    """写安装设置到注册表（对应官方 GetInstallSettings 读取的键）。"""
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, INSTALL_KEY) as k:
        winreg.SetValueEx(k, "Trust Folder", 0, winreg.REG_SZ, "1")
        winreg.SetValueEx(k, "Use Ribbon", 0, winreg.REG_SZ, "0")    # 仓库无编译 Ribbon DLL
        winreg.SetValueEx(k, "Compile accde", 0, winreg.REG_SZ, "0")
        winreg.SetValueEx(k, "Open File", 0, winreg.REG_SZ, "0")
        winreg.SetValueEx(k, "Install Folder", 0, winreg.REG_SZ, install_folder)
    print("[设置] Install Folder = %s" % install_folder)
    print("[设置] Use Ribbon = 0（无编译 Ribbon DLL，Menu Add-Ins + API 通道不受影响）")


def verify_install(install_folder: str) -> bool:
    ok = True
    # 1. 目标文件
    dest = os.path.join(install_folder, ADDIN_BASENAME + ".accda")
    exists = os.path.isfile(dest)
    print("[验证] 目标文件 %s：%s (%d B)" % (dest, "存在" if exists else "缺失",
                                             os.path.getsize(dest) if exists else 0))
    ok = ok and exists
    # 2. Menu Add-Ins 三个键（Expression / Library / Version）
    menu_root = r"Software\Microsoft\Office\16.0\Access\Menu Add-Ins"
    for item in ("&VCS Open", "&VCS Options", "&VCS Export All Source"):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, menu_root + "\\" + item) as k:
                lib, _ = winreg.QueryValueEx(k, "Library")
                expr, _ = winreg.QueryValueEx(k, "Expression")
                ver, _ = winreg.QueryValueEx(k, "Version")
            good = (lib.lower() == dest.lower()) and bool(expr)
            print("[验证] %s：Library=%s Expression=%s Version=%s -> %s" %
                  (item, lib, expr, ver, "OK" if good else "异常"))
            ok = ok and good
        except OSError as e:
            print("[验证] %s：注册表缺失 (%s)" % (item, e))
            ok = False
    # 3. 受信任位置
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\MSAccessVCS Version Control") as k:
            path, _ = winreg.QueryValueEx(k, "Path")
        print("[验证] 受信任位置：%s" % path)
        ok = ok and path.lower().rstrip("\\") == install_folder.lower().rstrip("\\")
    except OSError as e:
        print("[验证] 受信任位置：缺失 (%s)" % e)
        ok = ok and False
    # 4. Ribbon COM 加载项（DLL 文件 + COM Add-Ins 注册 + CLSID/ProgID）
    dll = os.path.join(install_folder, "MSAccessVCSLib_win64.dll")
    if os.path.isfile(dll):
        print("[验证] Ribbon DLL：%s (%d B)" % (dll, os.path.getsize(dll)))
        ok = ok and True
    else:
        print("[验证] Ribbon DLL：缺失 %s" % dll)
        ok = ok and False
    addins_key = r"Software\Microsoft\Office\Access\Addins\MSAccessVCSLib.AddInRibbon"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, addins_key) as k:
            lb, _ = winreg.QueryValueEx(k, "LoadBehavior")
            fn, _ = winreg.QueryValueEx(k, "FriendlyName")
        good = (int(lb) == 3) and bool(fn)
        print("[验证] COM Add-Ins：LoadBehavior=%s FriendlyName=%s -> %s" % (lb, fn, "OK" if good else "异常"))
        ok = ok and good
    except OSError as e:
        print("[验证] COM Add-Ins：注册缺失 (%s)" % e)
        ok = ok and False
    # 5. 技能目录自包含清单（accda + DLL + Ribbon.xml + Ribbon.json + Worker.vbs）
    need = ["Version Control.accda", "MSAccessVCSLib_win64.dll", "Ribbon.xml", "Ribbon.json", "Worker.vbs"]
    for n in need:
        p = os.path.join(install_folder, n)
        if not os.path.isfile(p):
            print("[验证] 自包含组件缺失：%s" % p)
            ok = ok and False
    print("[验证] 自包含组件：%s" % ("全部就位" if ok else "有缺失"))
    return ok


def main():
    argv = sys.argv[1:]
    if not argv or "-h" in argv or "--help" in argv:
        print(__doc__)
        return 1
    source = None
    folder = None
    ribbon_dll = None      # 官方编译的 Ribbon COM DLL（win64）
    ribbon_xml = None
    ribbon_json = None
    for i, a in enumerate(argv):
        if a == "--source" and i + 1 < len(argv):
            source = os.path.abspath(argv[i + 1])
        if a == "--folder" and i + 1 < len(argv):
            folder = os.path.abspath(argv[i + 1])
        if a == "--ribbon-dll" and i + 1 < len(argv):
            ribbon_dll = os.path.abspath(argv[i + 1])
        if a == "--ribbon-xml" and i + 1 < len(argv):
            ribbon_xml = os.path.abspath(argv[i + 1])
        if a == "--ribbon-json" and i + 1 < len(argv):
            ribbon_json = os.path.abspath(argv[i + 1])
    if not source or not folder:
        print(__doc__)
        return 1
    if not os.path.isfile(source):
        print("[致命] 安装源不存在: %s" % source)
        return 2
    os.makedirs(folder, exist_ok=True)

    msaccess = find_msaccess()
    print("[MSAccess] %s" % msaccess)

    set_install_settings(folder)

    # Ribbon COM 加载项：复制官方编译 DLL / Ribbon.xml / Ribbon.json 到安装文件夹，
    # 并注册 COM Add-Ins（对应官方 VerifyComAddIn 的产出与 DllRegisterServer 的注册键）。
    if ribbon_dll and os.path.isfile(ribbon_dll):
        import shutil
        shutil.copy2(ribbon_dll, os.path.join(folder, "MSAccessVCSLib_win64.dll"))
        if ribbon_xml and os.path.isfile(ribbon_xml):
            shutil.copy2(ribbon_xml, os.path.join(folder, "Ribbon.xml"))
        if ribbon_json and os.path.isfile(ribbon_json):
            shutil.copy2(ribbon_json, os.path.join(folder, "Ribbon.json"))
        addins_key = r"Software\Microsoft\Office\Access\Addins\MSAccessVCSLib.AddInRibbon"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, addins_key) as k:
            winreg.SetValueEx(k, "FriendlyName", 0, winreg.REG_SZ,
                              "Ribbon integration for MSAccessVCS add-in")
            winreg.SetValueEx(k, "Description", 0, winreg.REG_SZ,
                              "Microsoft Access COM add-in to add Fluent UI ribbon support to Access add-in project")
            winreg.SetValueEx(k, "ProgID", 0, winreg.REG_SZ, "MSAccessVCSLib.AddInRibbon")
            winreg.SetValueEx(k, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        print("[Ribbon] DLL/Ribbon.xml/Ribbon.json 已装入 %s，COM Add-Ins 已注册（LoadBehavior=3）" % folder)
    else:
        print("[Ribbon] 未提供 --ribbon-dll，跳过 Ribbon COM 加载项（仅 Menu Add-Ins + API 通道可用）")

    # 弹窗守卫（捕获消息框内容 + 自动点击按钮）
    fwguard.start_guard()
    try:
        fwguard.kill_access(os.path.dirname(source))
        time.sleep(2)
        cmd = '"%s" /cmd INSTALL' % source
        print("[安装] %s %s" % (msaccess, cmd))
        proc = subprocess.Popen([msaccess, source, "/cmd", "INSTALL"],
                                cwd=os.path.dirname(source))
        try:
            proc.wait(timeout=300)
            print("[安装] 进程退出，代码=%s" % proc.returncode)
        except subprocess.TimeoutExpired:
            print("[安装] 超时，强制结束")
            proc.kill()
            return 3
        time.sleep(3)
    finally:
        fwguard.kill_access(os.path.dirname(source))
        fwguard.stop_guard()
        for s in fwguard.dialogs_seen()[:10]:
            print("[弹窗] %r 正文=%s" % (s[0], (s[1][:160] if s[1] else "")))

    ok = verify_install(folder)
    print("[结论] 安装%s" % ("成功" if ok else "失败，请检查上述验证项"))
    return 0 if ok else 4


if __name__ == "__main__":
    raise SystemExit(main())
