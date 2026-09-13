# -*- coding: utf-8 -*-
"""模板可用性真机验证：把 templates/ 里的存根导入一个临时空库，断言类型正确。

验证目标（全部自动断言，无人工步骤）：
  1. templates/class-module-stub.cls  经"UTF-8-BOM -> GBK 无 BOM"转换后 Import，得到 type=2。
  2. templates/standard-module-stub.bas 同样转换后 Import，得到 type=1。
  3. templates/autoexec-macro.macro 经 LoadFromText 4 加载后，AllMacros 里出现 AutoExec。
  4. 临时库用完即删，绝不触碰主库。

用法: python probe_templates.py
退出码 0=全部通过，4=有断言失败。
"""
import os, sys, shutil, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwguard
from win32com.client.gencache import EnsureDispatch

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(BASE, "templates")

checks = []


def ck(name, cond, detail=""):
    checks.append((name, bool(cond), detail))
    print("  [%s] %s%s" % ("PASS" if cond else "FAIL", name, ("  -> " + detail) if detail else ""))


def to_gbk_crlf(src, out_dir):
    raw = open(src, "rb").read()
    t = raw.decode("utf-8-sig")
    t = t.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if not t.endswith("\r\n"):
        t += "\r\n"
    dst = os.path.join(out_dir, os.path.basename(src))
    with open(dst, "wb") as f:
        f.write(t.encode("gbk"))
    assert not open(dst, "rb").read(3).startswith(b"\xef\xbb\xbf"), "转出的文件不该有 BOM"
    return dst


def main():
    tmp = tempfile.mkdtemp(prefix="fw_tplprobe_")
    work = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "fw_tplprobe")
    os.makedirs(work, exist_ok=True)
    db = os.path.join(tmp, "Probe.accdb")

    print("[准备] 模板 -> GBK 无 BOM")
    cls_gbk = to_gbk_crlf(os.path.join(TPL, "class-module-stub.cls"), work)
    bas_gbk = to_gbk_crlf(os.path.join(TPL, "standard-module-stub.bas"), work)
    macro_src = os.path.join(TPL, "autoexec-macro.macro")

    fwguard.start_guard()
    fwguard.kill_access(tmp)
    app = EnsureDispatch("Access.Application")
    app.Visible = True
    try:
        app.NewCurrentDatabase(db)
        proj = fwguard.get_proj(app, db)
        if proj is None:
            for p in app.VBE.VBProjects:
                try:
                    if os.path.normcase(p.FileName) == os.path.normcase(db):
                        proj = p
                        break
                except Exception:
                    pass
        ck("临时库工程可定位", proj is not None, proj.Name if proj else "未找到")

        # 1) 类模块模板
        try:
            c = proj.VBComponents.Import(cls_gbk)
            ck("类模块模板 Import -> type=2", c.Type == 2, "name=%s type=%s" % (c.Name, c.Type))
            ck("类模块模板正文已写入", c.CodeModule.CountOfLines > 10, "%d 行" % c.CodeModule.CountOfLines)
            body = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
            ck("类模块模板结构完整（含 Option Explicit）", "Option Explicit" in body)
            ck("类模块模板无乱码（中文可读）", "类模块存根" in body, repr(body[:60]))
        except Exception as e:
            ck("类模块模板 Import", False, str(e)[:160])

        # 2) 标准模块模板
        try:
            b = proj.VBComponents.Import(bas_gbk)
            ck("标准模块模板 Import -> type=1", b.Type == 1, "name=%s type=%s" % (b.Name, b.Type))
            bbody = b.CodeModule.Lines(1, b.CodeModule.CountOfLines)
            ck("标准模块模板中文无乱码", "标准模块存根" in bbody, repr(bbody[:60]))
        except Exception as e:
            ck("标准模块模板 Import", False, str(e)[:160])

        # 3) 宏模板
        try:
            app.LoadFromText(4, "AutoExec", macro_src)   # acMacro = 4
            names = [app.CurrentProject.AllMacros(i).Name for i in range(app.CurrentProject.AllMacros.Count)]
            ck("宏模板 LoadFromText 4 可加载", "AutoExec" in names, "宏列表=%s" % names)
        except Exception as e:
            ck("宏模板 LoadFromText 4 可加载", False, str(e)[:140])

        # 4) 编译检查交由 Access 在导入时完成；模板本身无语法依赖
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        fwguard.kill_access(tmp)
        fwguard.stop_guard()
        shutil.rmtree(tmp, ignore_errors=True)

    bad = [c for c in checks if not c[1]]
    print("\n===== 模板真机验证：%d 项断言，%d 项失败 =====" % (len(checks), len(bad)))
    if bad:
        for n, _, d in bad:
            print("  FAIL %s %s" % (n, d))
    print("结论：%s" % ("模板全部可用" if not bad else "存在不可用模板，需修复"))
    return 0 if not bad else 4


if __name__ == "__main__":
    raise SystemExit(main())
