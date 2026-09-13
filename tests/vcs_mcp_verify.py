# -*- coding: utf-8 -*-
"""MCP 端到端验证：启动 msaccess-vcs-mcp 服务器，走真实 JSON-RPC 调用，硬断言后返回退出码。

用法: python vcs_mcp_verify.py --db <宿主库.accdb>
宿主库需含可调用函数 Hello()（返回 "VCS OK"）与模块 basSample，见 tests/README.md。
返回: 0 = 全部 6 步通过；1 = 任一步失败。
"""
import json, os, subprocess, sys, time

_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MCP_DIR = os.path.join(_SKILL_DIR, "msaccess-vcs-mcp")
sys.path.insert(0, os.path.join(_SKILL_DIR, "scripts"))
import fwguard
PY = sys.executable

if "--db" not in sys.argv:
    print(__doc__)
    raise SystemExit(1)
DB = os.path.abspath(sys.argv[sys.argv.index("--db") + 1])

before = fwguard.snapshot_access()

proc = subprocess.Popen(
    [PY, "-m", "msaccess_vcs_mcp.main"],
    cwd=MCP_DIR,
    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    text=True, encoding="utf-8", bufsize=1,
)

import threading, queue
_q = queue.Queue()
def _reader():
    for line in proc.stdout:
        _q.put(("out", line))
    _q.put(("eof", None))
threading.Thread(target=_reader, daemon=True).start()

def send(obj):
    line = json.dumps(obj, ensure_ascii=False)
    proc.stdin.write(line + "\n")
    proc.stdin.flush()

def recv(timeout=120):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            kind, line = _q.get(timeout=1.0)
        except queue.Empty:
            continue
        if kind == "eof":
            return None
        line = line.strip()
        if not line:
            continue
        try:
            return json.loads(line)
        except Exception as e:
            print("  非 JSON 行: %r (%s)" % (line[:120], e))
            continue
    return {"timeout": True}

def call(method, params, mid):
    send({"jsonrpc": "2.0", "id": mid, "method": method, "params": params})
    while True:
        resp = recv()
        if resp is None:
            return {"error": "EOF"}
        if resp.get("id") == mid or "result" in resp or "error" in resp:
            if resp.get("id") == mid:
                return resp

def tool_ok(r):
    """MCP 工具返回的 structuredContent.success 是否为 True。"""
    try:
        return r.get("result", {}).get("structuredContent", {}).get("success") is True
    except Exception:
        return False

ok_all = True

# 1. initialize
print("== 1. initialize ==")
r = call("initialize", {
    "protocolVersion": "2024-11-05",
    "capabilities": {},
    "clientInfo": {"name": "verify", "version": "1.0"},
}, 1)
si = (r.get("result") or {}).get("serverInfo")
print("  server:", si)
ok_all = ok_all and bool(si)

# 2. initialized 通知
send({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})

# 3. tools/list
print("== 2. tools/list ==")
r = call("tools/list", {}, 2)
tools = [t["name"] for t in (r.get("result") or {}).get("tools", [])]
print("  工具数:", len(tools))
has_expected = all(t in tools for t in
                   ("vcs_export_database", "vcs_list_objects", "vcs_call_vba",
                    "vcs_compile_vba", "vcs_export_object", "vcs_import_object"))
ok_all = ok_all and has_expected
print("  关键工具齐备:", "OK" if has_expected else "缺失")

# 4. vcs_list_objects
print("== 3. vcs_list_objects(database_path) ==")
r = call("tools/call", {"name": "vcs_list_objects", "arguments": {"database_path": DB}}, 3)
good = tool_ok(r)
print("  返回:", json.dumps(r.get("result") or {}, ensure_ascii=False)[:600])
print("  断言: %s" % ("OK" if good else "失败"))
ok_all = ok_all and good

# 5. vcs_call_vba（调用库内 Hello 函数验证运行链路）
print("== 4. vcs_call_vba(Hello) ==")
r = call("tools/call", {
    "name": "vcs_call_vba",
    "arguments": {"database_path": DB, "function_name": "Hello"},
}, 4)
res = r.get("result") or {}
good = tool_ok(r) and (res.get("structuredContent", {}).get("result") == "VCS OK")
print("  返回:", json.dumps(res, ensure_ascii=False)[:400])
print("  断言: %s" % ("OK" if good else "失败（Hello 应返回 VCS OK）"))
ok_all = ok_all and good

# 6. vcs_compile_vba
print("== 5. vcs_compile_vba ==")
r = call("tools/call", {"name": "vcs_compile_vba", "arguments": {"database_path": DB}}, 5)
good = tool_ok(r)
print("  返回:", json.dumps(r.get("result") or {}, ensure_ascii=False)[:300])
print("  断言: %s" % ("OK" if good else "失败"))
ok_all = ok_all and good

# 7. vcs_export_object（单对象导出，验证写读闭环 + 错误 91 修复）
print("== 6. vcs_export_object(modules, basSample) ==")
r = call("tools/call", {
    "name": "vcs_export_object",
    "arguments": {"database_path": DB, "object_type": "modules", "object_name": "basSample"},
}, 6)
good = tool_ok(r) and bool((r.get("result") or {}).get("structuredContent", {}).get("logPath"))
print("  返回:", json.dumps(r.get("result") or {}, ensure_ascii=False)[:400])
print("  断言: %s" % ("OK" if good else "失败（应返回 success + logPath）"))
ok_all = ok_all and good

proc.stdin.close()
try:
    proc.wait(timeout=10)
except Exception:
    proc.kill()

# 收尾：清理 MCP 本次新拉起的 Access 实例（只杀新进程，不误伤用户实例）
fwguard.kill_new_access(before)

print("===== MCP 端到端判定: %s =====" % ("全部通过" if ok_all else "存在失败项"))
raise SystemExit(0 if ok_all else 1)
