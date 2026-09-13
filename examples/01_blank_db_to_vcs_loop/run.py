#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""示例 01：从空库到「导入模块 -> VCS 导出 .src -> 漂移核验」的端到端闭环。

演示版本控制作为贯穿方法：库一建好就导出文本源作为单一真相源，改完再核验。
前提：本机已装 Access（2010+）与技能目录运行位的 VCS 加载项
      Microsoft Access Version Control System\\Version Control.accda，且 Python 已装 pywin32。
运行：python run.py   （退出码 0 = 闭环全过；1 = 任一步失败）
"""
import os
import sys
import shutil
import subprocess
import tempfile

# 让示例能直接 import 本技能的弹窗守卫（不依赖人工配置路径）
_SKILL = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILL, "scripts"))
import fwguard  # noqa: E402

SCRIPTS = os.path.join(_SKILL, "scripts")
ADDIN = os.path.join(_SKILL, "Microsoft Access Version Control System", "Version Control.accda")


def step(msg):
    print("\n=== " + msg + " ===")


def run_script(rel, *args):
    cmd = [sys.executable, os.path.join(SCRIPTS, rel)] + list(args)
    print("+ " + " ".join(cmd))
    return subprocess.run(cmd).returncode


def main():
    # 0) 前提检查
    step("0) 前提检查")
    try:
        import win32com.client  # noqa: F401
    except Exception:
        print("缺少 pywin32：请先 `pip install pywin32`"); return 1
    if not os.path.isfile(ADDIN):
        print("未找到 VCS 加载项正式安装版：\n  %s\n请先安装 Version Control.accda。" % ADDIN)
        return 1
    print("OK：pywin32 可用；加载项=%s" % ADDIN)

    work = tempfile.mkdtemp(prefix="mads_ex01_")
    db = os.path.join(work, "Sample.accdb")
    src = os.path.join(work, "Sample.src")
    mod = os.path.join(work, "basSample.bas")
    try:
        # 1) 建空库
        step("1) 建空库（win32com.NewCurrentDatabase，零手动）")
        fwguard.kill_access(work)
        from win32com.client.gencache import EnsureDispatch
        app = EnsureDispatch("Access.Application")
        app.Visible = True
        if os.path.isfile(db):
            os.remove(db)
        app.NewCurrentDatabase(db)
        app.CloseCurrentDatabase()
        app.Quit(2)
        print("已建空库：%s" % db)

        # 2) 写示例模块文本（UTF-8，含 Attribute VB_Name）
        step("2) 写示例模块文本")
        code = (
            'Attribute VB_Name = "basSample"\r\n'
            'Option Compare Database\r\n'
            'Option Explicit\r\n'
            '\r\n'
            'Public Function Hello() As String\r\n'
            '    Hello = "你好，Access VCS 闭环"\r\n'
            'End Function\r\n'
        )
        with open(mod, "w", encoding="utf-8") as f:
            f.write(code)
        print("已写模块：%s" % mod)

        # 3) 经 VCS 通道原子导入（rebuild 会转 GBK 无 BOM + Import + 读回比对）
        step("3) 导入模块（自动化，不手工粘贴）")
        if run_script("rebuild_module_from_src.py", db, mod) != 0:
            print("导入失败"); return 1

        # 4) 导出 .src 作为单一真相源
        step("4) 导出 .src（版本控制：库一建好就纳入管理）")
        if os.path.isdir(src):
            shutil.rmtree(src, ignore_errors=True)
        if run_script("vcs_drive.py", "FullExport", src, "--db", db, "--settle", "30") != 0:
            print("导出失败（检查加载项/库状态）"); return 1

        # 5) 漂移核验：实时库 vs .src（断言化：0=一致通过；3=有漂移失败）
        step("5) 漂移核验（改完必跑）")
        rc = run_script("vcs_consistency_check.py", db, src)
        if rc == 0:
            print("PASS：实时库与 .src 完全一致，闭环成立。")
        elif rc == 3:
            print("FAIL：存在漂移（vcs_consistency_check 退出码 3），闭环未成立。"); return 1
        else:
            print("核验异常退出码=%s" % rc); return 1
        return 0
    finally:
        fwguard.kill_access(work)
        try:
            shutil.rmtree(work, ignore_errors=True)
            print("\n[清理] 临时目录已删除：%s" % work)
        except Exception:
            print("[清理] 请手动删除：%s" % work)


if __name__ == "__main__":
    raise SystemExit(main())
