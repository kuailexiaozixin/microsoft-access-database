#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""示例 02：增量回写——改文本源 -> 原子重导入 -> 断言生效。

演示日常开发最常拍的一下：在 .src 文本上改一处，用版本控制原子重导入回库，
再用 Application.Run 复核改动生效。前提同示例 01（Access + 技能目录运行位的
VCS 加载项 + pywin32）。
运行：python run.py   （退出码 0 = 闭环全过；1 = 任一步失败）
"""
import os
import sys
import shutil
import subprocess
import tempfile

_SKILL = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILL, "scripts"))
import fwguard  # noqa: E402

SCRIPTS = os.path.join(_SKILL, "scripts")
ADDIN = os.path.join(_SKILL, "Microsoft Access Version Control System", "Version Control.accda")

SAMPLE_CODE = (
    'Attribute VB_Name = "basSample"\r\n'
    'Option Compare Database\r\n'
    'Option Explicit\r\n'
    '\r\n'
    'Public Function Hello() As String\r\n'
    '    Hello = "%s"\r\n'
    'End Function\r\n'
)


def step(msg):
    print("\n=== " + msg + " ===")


def run_script(rel, *args):
    cmd = [sys.executable, os.path.join(SCRIPTS, rel)] + list(args)
    print("+ " + " ".join(cmd))
    return subprocess.run(cmd).returncode


def run_func(db, func):
    """开库 -> Application.Run(func) -> 关库，返回字符串结果。"""
    from win32com.client.gencache import EnsureDispatch
    fwguard.kill_access(os.path.dirname(db))
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    app.OpenCurrentDatabase(db)
    try:
        v = app.Run(func)
        if isinstance(v, tuple):
            v = v[0]
        return v
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        fwguard.kill_access(os.path.dirname(db))


def main():
    step("0) 前提检查")
    try:
        import win32com.client  # noqa: F401
    except Exception:
        print("缺少 pywin32：请先 `pip install pywin32`"); return 1
    if not os.path.isfile(ADDIN):
        print("未找到 VCS 加载项正式安装版：\n  %s\n请先安装 Version Control.accda。" % ADDIN)
        return 1
    print("OK：pywin32 可用；加载项=%s" % ADDIN)

    work = tempfile.mkdtemp(prefix="mads_ex02_")
    db = os.path.join(work, "Sample.accdb")
    mod = os.path.join(work, "basSample.bas")
    try:
        # 1) 建空库
        step("1) 建空库")
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

        # 2) 写初值模块并导入
        step("2) 写初值 v1 并导入")
        with open(mod, "w", encoding="utf-8") as f:
            f.write(SAMPLE_CODE % "v1")
        if run_script("rebuild_module_from_src.py", db, mod) != 0:
            print("导入失败"); return 1

        # 3) 复核初值
        step("3) 复核 Hello()（期望 v1）")
        v1 = run_func(db, "Hello")
        print("Hello() = %r" % v1)
        if v1 != "v1":
            print("初值异常"); return 1

        # 4) 在文本源上改：v1 -> v2（模拟文本化开发）
        step("4) 改文本源：v1 -> v2")
        with open(mod, "w", encoding="utf-8") as f:
            f.write(SAMPLE_CODE % "v2")
        print("已改 %s 的返回值为 v2" % mod)

        # 5) 原子重导入（Remove + Import + 读回比对，不手工粘贴）
        step("5) 原子重导入回库")
        if run_script("rebuild_module_from_src.py", db, mod) != 0:
            print("回写失败"); return 1

        # 6) 复核改动生效
        step("6) 复核 Hello()（期望 v2）")
        v2 = run_func(db, "Hello")
        print("Hello() = %r" % v2)
        if v2 == "v2":
            print("PASS：改文本源 -> 回写 -> 生效，闭环成立。")
            return 0
        print("FAIL：回写未生效（Hello()=%r，期望 v2）" % v2); return 1
    finally:
        fwguard.kill_access(work)
        try:
            shutil.rmtree(work, ignore_errors=True)
            print("\n[清理] 临时目录已删除：%s" % work)
        except Exception:
            print("[清理] 请手动删除：%s" % work)


if __name__ == "__main__":
    raise SystemExit(main())
