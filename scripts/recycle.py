# -*- coding: utf-8 -*-
r"""把文件/目录移入 Windows 回收站（绝不永久删除）。本机铁律的工具化实现。

用法:
    python recycle.py <文件或目录> [<文件或目录> ...]

原理：SHFileOperationW + FOF_ALLOWUNDO（走回收站）+ FOF_NOCONFIRMATION（不弹确认框）。
为什么不用 os.remove / shutil.rmtree：本机铁律要求「删除 = 回收站」，可随时还原。

返回：0 = 全部成功；5 = 有目标不存在；6 = 有回收失败。

踩过的坑（2026-09-11）：
  · 保留设备名（`nul`/`con`/`aux`…）会让 SHFileOperation **整批**返回 124，
    连累同批的正常文件。对策：先用 `\\?\` 前缀改名再回收，并**把可疑项单独成批**。
  · 因此本脚本按 50 个一批提交；失败时能缩小到具体批次。
"""
import os, sys, ctypes
from ctypes import wintypes

FO_DELETE = 3
FOF_SILENT = 0x0004
FOF_NOCONFIRMATION = 0x0010
FOF_ALLOWUNDO = 0x0040
FOF_NOERRORUI = 0x0400
FOF_NOCONFIRMMKDIR = 0x0200


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", wintypes.UINT),
        ("pFrom", wintypes.LPCWSTR),
        ("pTo", wintypes.LPCWSTR),
        ("fFlags", ctypes.c_uint16),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", wintypes.LPCWSTR),
    ]


def recycle(paths):
    """返回 (成功列表, 不存在列表, 失败列表)"""
    ok, missing, failed = [], [], []
    batch, n = [], 0
    for p in paths:
        if not os.path.exists(p):
            missing.append(p)
            continue
        batch.append(os.path.abspath(p))
        if len(batch) == 50:                       # 分批，便于定位失败
            _do(batch, ok, failed)
            batch = []
    if batch:
        _do(batch, ok, failed)
    return ok, missing, failed


def _do(items, ok, failed):
    src = "\0".join(items) + "\0\0"
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = src
    op.pTo = None
    op.fFlags = (FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
                 | FOF_NOERRORUI | FOF_NOCONFIRMMKDIR)
    op.fAnyOperationsAborted = False
    res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if res == 0 and not op.fAnyOperationsAborted:
        ok.extend(items)
        for p in items:
            print("  已移入回收站: %s" % p)
    else:
        failed.extend(items)
        print("  回收失败(码 %s, 中止=%s): %s" % (res, bool(op.fAnyOperationsAborted), items))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    ok, missing, failed = recycle(sys.argv[1:])
    print("\n成功 %d / 不存在 %d / 失败 %d" % (len(ok), len(missing), len(failed)))
    for p in missing:
        print("  不存在: %s" % p)
    return 0 if not failed and not missing else (5 if missing else 6)


if __name__ == "__main__":
    raise SystemExit(main())
