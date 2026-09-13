# -*- coding: utf-8 -*-
"""从 .src 原子重建 Access VBA 模块（GBK 无 BOM + 读回逐行比对）。

用法:
    python rebuild_module_from_src.py <主库.accdb> <src模块文件> [<src模块文件> ...]

要点（踩过的坑都在这里）:
  1. 必须按 FileName 定位 VBProject —— ActiveVBProject 会指向 VCS 加载项。
  2. 导入文件必须 GBK 无 BOM，否则中文乱码、.cls 还会退化为 type=1。
  3. 用 Remove + Import 原子替换，不要用 InsertLines 灌大段代码（会部分写入）。
  4. 导入后读回比对，不一致要打印首个差异行。
  5. 脚本必须带弹窗守卫，否则运行时错误框会永久阻塞。
  6. 【2026-09-11 实测】Import 建出的组件名取自文件头部的 Attribute VB_Name，
     不是文件名 → 移除旧组件必须按 Attribute 名匹配，否则会漏删留下僵尸模块。
"""
import os, sys, time, shutil, datetime, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

SCALAR_HEAD = ("Attribute ", "VERSION ", "BEGIN", "END", "MultiUse", "BaseClass")

# 匹配 Attribute VB_Name = "xxx"（.cls/.bas 头部都有）
RE_VBNAME = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE)


def component_name(src_path, text):
    """组件名 = 文件头部 Attribute VB_Name；没有就退回文件名（不含扩展名）。"""
    for ln in text.split("\r\n")[:30]:
        m = RE_VBNAME.match(ln)
        if m:
            return m.group(1)
    return os.path.splitext(os.path.basename(src_path))[0]


def to_gbk_crlf(src_path, out_dir):
    raw = open(src_path, "rb").read()
    t = raw.decode("utf-8-sig")
    t = t.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if not t.endswith("\r\n"):
        t += "\r\n"
    dst = os.path.join(out_dir, os.path.basename(src_path))
    open(dst, "wb").write(t.encode("gbk"))
    return dst, t


def strip_head(lines):
    """剥掉 VBE 不保存的头部行（.cls 的 VERSION/BEGIN/END/Attribute 等）"""
    out, started = [], False
    for ln in lines:
        s = ln.strip()
        if not started:
            if any(s.startswith(h) for h in SCALAR_HEAD):
                continue
            started = True
        out.append(ln)
    return out


def compare(src_text, live_text, label):
    s = strip_head(src_text.split("\r\n"))
    l = live_text.split("\r\n")
    while l and l[-1] == "":
        l.pop()
    while s and s[-1] == "":
        s.pop()
    if s == l:
        print("    [比对] %s 一致（%d 行）" % (label, len(l)))
        return True
    print("    [比对] %s 不一致：源 %d 行 / 库 %d 行" % (label, len(s), len(l)))
    for i in range(max(len(s), len(l))):
        a = s[i] if i < len(s) else "<无>"
        b = l[i] if i < len(l) else "<无>"
        if a != b:
            print("        首个差异 行%d\n          源: %r\n          库: %r" % (i + 1, a[:90], b[:90]))
            break
    return False


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 1
    db = os.path.abspath(sys.argv[1])
    mods = [os.path.abspath(p) for p in sys.argv[2:]]
    base = os.path.dirname(db)
    tmp = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "fw_rebuild")
    os.makedirs(tmp, exist_ok=True)

    preps = []
    for m in mods:
        dst, text = to_gbk_crlf(m, tmp)
        comp = component_name(m, text)     # 按 Attribute VB_Name 取组件名（实测规则）
        stem = os.path.splitext(os.path.basename(m))[0]
        preps.append((comp, dst, text))
        print("[准备] %s -> GBK 无 BOM（组件名=%s%s）"
              % (stem, comp, "" if comp == stem else "，注意与文件名不同"))

    fwguard.start_guard()
    fwguard.kill_access(base)
    bak = db + ".bak_rebuild_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(db, bak)
    print("[备份] %s" % os.path.basename(bak))

    app = EnsureDispatch("Access.Application")
    app.Visible = True
    app.OpenCurrentDatabase(db)
    try:
        # VBE 已在 get_proj 枚举 VBProjects 时完成初始化
        proj = fwguard.get_proj(app, db)
        if proj is None:
            print("[致命] 按 FileName 定位不到目标工程，中止")
            return 2
        print("[工程] %s" % proj.Name)

        for comp, path, text in preps:
            removed = 0
            for c in list(proj.VBComponents):
                if c.Name.lower() == comp.lower():
                    proj.VBComponents.Remove(c)
                    print("[移除] %s" % comp)
                    removed += 1
                    break
            if not removed:
                print("[提示] 库中没有同名组件 %s，按新建处理" % comp)
            c = proj.VBComponents.Import(path)
            actual = c.Name
            print("[导入] %s type=%s 行数=%d" % (actual, c.Type, c.CodeModule.CountOfLines))
            if actual.lower() != comp.lower():
                print("[警告] 实际组件名(%s) 与 预期(%s) 不一致，请检查文件头部 Attribute VB_Name" % (actual, comp))
            compare(text, c.CodeModule.Lines(1, c.CodeModule.CountOfLines), actual)
            try:
                app.DoCmd.Save(5, actual)   # acModule = 5
                print("[保存] %s" % actual)
            except Exception as e:
                print("[保存失败] %s: %s" % (actual, str(e)[:120]))
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
    print("done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
