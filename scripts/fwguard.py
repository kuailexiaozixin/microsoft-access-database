# -*- coding: utf-8 -*-
"""Access 自动化通用工具：弹窗守卫 + 进程清理。所有脚本统一 import 本模块。

关键经验：
- VBE 错误框（标题 'Microsoft Visual Basic'）的按钮文本常被渲染截断为单字
  （如 '继续'->'续'、'结束'->'束'），所以匹配必须按【首字】。
- 有些错误框根本没有 Button 子控件，此时退回 PostMessage(WM_COMMAND, IDOK) + 回车键。
"""
import os, time, glob, threading
import win32gui, win32con

_stop = threading.Event()
_logged = []

# 首字 -> 期望点击；顺序即优先级
_FIRSTCHAR = ("确", "结", "束", "O", "o", "是", "Y", "y", "好", "继", "续", "试", "助")


def children(h):
    out = []
    def cb(ch, _):
        try:
            out.append((ch, win32gui.GetClassName(ch), win32gui.GetWindowText(ch)))
        except Exception:
            pass
    try:
        win32gui.EnumChildWindows(h, cb, None)
    except Exception:
        pass
    return out


# 危险字样的对话框：一律点"取消/否"，绝不点"是"
_DANGER_WORDS = ("另存", "保存", "覆盖", "替换", "删除", "确认修改")


def _try_dismiss(h):
    """返回 True 表示已尝试点击"""
    kids = children(h)
    buttons = [(c, t) for c, cls, t in kids if cls == "Button"]
    # 0) 危险字样 -> 取消/否
    blob = win32gui.GetWindowText(h) + " " + " ".join(x[2] for x in kids)
    if any(w in blob for w in _DANGER_WORDS):
        for pref in ("取", "C", "否", "N"):
            for ch, bt in buttons:
                if bt.replace("&", "").strip().startswith(pref):
                    try:
                        win32gui.SendMessage(ch, win32con.BM_CLICK, 0, 0)
                        return True
                    except Exception:
                        pass
    # 1) 按首字匹配
    for pref in _FIRSTCHAR:
        for ch, bt in buttons:
            b = bt.replace("&", "").strip()
            if b.startswith(pref):
                try:
                    win32gui.SendMessage(ch, win32con.BM_CLICK, 0, 0)
                    return True
                except Exception:
                    pass
    # 2) 无按钮或未匹配：发 IDOK / 回车 / 关闭
    try:
        win32gui.PostMessage(h, win32con.WM_COMMAND, 1, 0)
    except Exception:
        pass
    try:
        win32gui.PostMessage(h, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(h, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    except Exception:
        pass
    return False


def _hit_list():
    hits = []
    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return
            cls = win32gui.GetClassName(h)
            t = win32gui.GetWindowText(h)
            if cls == "#32770" and "Microsoft Visual Basic for Applications" not in t:
                hits.append(h)
        except Exception:
            pass
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return hits


def _guard_loop(verbose=True):
    while not _stop.is_set():
        for h in _hit_list():
            t = win32gui.GetWindowText(h)
            kids = children(h)
            statics = [x[2] for x in kids if x[1] == "Static" and x[2].strip()]
            buttons = [x[2] for x in kids if x[1] == "Button"]
            key = (t, tuple(statics))
            if key not in _logged:
                _logged.append(key)
                if verbose:
                    print("    [弹窗] %r 正文=%s 按钮=%s" % (
                        t, statics[0][:130] if statics else "", buttons), flush=True)
            _try_dismiss(h)
        time.sleep(0.3)


def start_guard(verbose=True):
    th = threading.Thread(target=_guard_loop, args=(verbose,), daemon=True)
    th.start()
    return th


def stop_guard():
    _stop.set()


def dialogs_seen():
    return list(_logged)


def kill_access(base_dir):
    """向后兼容的无差别清理（会误杀用户正在使用的 Access，仅用于可控的脚本收尾）。"""
    os.system("taskkill /F /IM MSACCESS.EXE >nul 2>&1")
    time.sleep(1)
    remove_locks(base_dir)


def remove_locks(base_dir):
    for f in glob.glob(os.path.join(base_dir, "*.laccdb")):
        try:
            os.remove(f)
        except Exception:
            pass


def snapshot_access():
    """记录当前所有 MSACCESS.EXE 进程 ID（用于区分本次新拉起的实例）。"""
    out = set()
    try:
        text = os.popen('tasklist /FI "IMAGENAME eq MSACCESS.EXE" /FO CSV /NH').read()
        for line in text.splitlines():
            if "MSACCESS" in line:
                parts = line.split(",")
                if len(parts) > 1:
                    out.add(int(parts[1].strip('"')))
    except Exception:
        pass
    return out


def kill_new_access(before_pids, base_dir=None):
    """只终止 before_pids 之外新出现的 Access 进程（避免误杀用户实例）。

    对高完整性孤儿进程（Access is denied）会报告但跳过——它们不占 COM 单实例，
    重启系统后消失。返回未能清理的 PID 集合。
    """
    import win32api, win32con
    now = snapshot_access()
    leftover = set()
    for pid in now - before_pids:
        try:
            h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
            if h:
                win32api.TerminateProcess(h, 1)
                print("    [清理] 已终止本次新实例 PID %s" % pid, flush=True)
        except Exception as e:
            leftover.add(pid)
            print("    [清理] 无法终止 PID %s（%s，可能为高完整性孤儿）" % (pid, str(e)[:60]), flush=True)
    if base_dir:
        remove_locks(base_dir)
    return leftover


def get_proj(app, db_path):
    """按 FileName 定位目标 VBProject（绝不使用 ActiveVBProject）"""
    for p in app.VBE.VBProjects:
        try:
            fn = p.FileName
        except Exception:
            fn = ""
        if fn and os.path.normcase(fn) == os.path.normcase(db_path):
            return p
    return None
