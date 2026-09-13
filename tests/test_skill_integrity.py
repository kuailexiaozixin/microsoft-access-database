#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/test_skill_integrity.py

纯 Python 冒烟测试（不需要 Access / COM，可在任意环境直接跑）：
  1. 校验 SKILL.md 引用的 12 个子项目目录都存在；
  2. 校验 references/ 与关键文件齐全；
  3. 校验模板文件的存放编码符合 iron-laws：标准模块 / 类模块以 UTF-8-BOM 存放（与 .src 一致，
     导入前由脚本转 GBK 无 BOM）；宏无 BOM；
  4. 扫描技能自有文件，确认无自研框架专有残留（clsLog / Framework_* / 门禁数字 / TS_* 链接表等；
     `changelog.md` 是唯一记载变化处，豁免扫描）。

运行：
  python tests/test_skill_integrity.py
退出码 0 = 全绿，1 = 有失败项。
"""

import os
import re

# 自研框架残留 token（changelog.md 除外：它是技能内唯一记载变化的地方，允许提及历史上的残留词）
# 注意："Framework.accdb" 须排除 EdonSoft 底座真实文件名 "EdonSoft Development Framework.accdb"，
#       用负向后行断言 (?<!Development ) 精确匹配"恰好叫 Framework.accdb"的残留。
TOKEN_RE = {t: re.compile(r"(?<!Development )" + re.escape(t)) for t in ("Framework.accdb",)}
import sys
import codecs

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 11 个子项目（与 SKILL.md 中引用完全一致；ms-access-ai-skill 已于 v2.9 移除）
SUBPROJECTS = [
    "盟威Access快速开发平台V2.7.0版(64位)",
    "Access BOM Management System",
    "Access DatePicker",
    "Access VBA Modules Collection",
    "AccessAI",
    "Edonsoft Development Framework_x64",
    "edonsoft-docs",
    "Microsoft Access Version Control System",
    "msaccess-vcs-addin",
    "msaccess-vcs-mcp",
    "Version_Control_v5.0.1",
]

# 已移入 examples/ 的子项目（作为可复用的积木型案例）
SUBPROJECTS_IN_EXAMPLES = {
    "Access BOM Management System",
    "Access DatePicker",
    "Access VBA Modules Collection",
    "AccessAI",
}

# 已移入 Edonsoft 框架文件夹内的子项目（作为其自带文档）
SUBPROJECTS_IN_EDONSOFT = {
    "edonsoft-docs",
}

# 自研框架专有残留（不应出现在技能自有文件里）
# 注：clsLog 已于 v2.17 移出禁词——经核实它是 Edonsoft 框架的可见源码模块
#   （framework-overview.md 第 4 节枚举：basCoreReference/basAI/basZip/basButton/
#   basFormFontTools/clsLog/JsonConverter/Module_DatePicker/Module_YearMonthPicker），
#   SKILL.md 引用其为合法借鉴目标；框架目录本身在 SKIP_DIRS 中，扫描不会误伤。
FORBIDDEN_TOKENS = [
    "Framework_RunAllTests",
    "Framework_VCSBackup",
    "Framework_API",
    "v3.11",
    "v3.7",
    "v3.8",
    "v3.6",
    "269 断言",
    "257 断言",
    "44 分组",
    "TS_* 链接表",
    "12 个 TS_",
    "edonsoft-gap",
    "Core_GetVersion",
    "FrameWork",  # 自研框架目录名（不应存在于技能目录中）
    "Framework.accdb",  # 自研框架主库文件名（脚本中不得残留此默认路径）
    "FRAME_DIR",  # 自研框架路径变量名（脚本中不得残留）
]

# 扫描时跳过的目录（第三方子项目 + 本测试自身目录 + examples 案例 + VCS 快照等无需扫描处）
SKIP_DIRS = set(SUBPROJECTS + [".git", "tests", "examples"])

REQUIRED_FILES = [
    "SKILL.md",
    "README.md",
    "docs/troubleshooting.md",
    "changelog.md",
    "references/README.md",
    "references/iron-laws.md",
    "references/vba-encoding-rules.md",
    "references/vba-com-automation.md",
    "references/vcs-three-channels.md",
    "references/vcs-text-editing.md",
    "references/access-architecture-network.md",
    "references/framework-comparison.md",
    # 对象级手册（9 份，上移自 ms-access-ai-skill，一对象一份）
    "references/forms.md",
    "references/reports.md",
    "references/images.md",
    "references/conditional-formatting.md",
    "references/queries.md",
    "references/tables-and-relationships.md",
    "references/vba.md",
    "references/project-config.md",
    "references/other-objects.md",
    # 完整业务样例（examples/CustomerOrders，9 份手册的活样例，v2.6 上移）
    "examples/CustomerOrders/CustomerOrders.accdb",
    "examples/CustomerOrders/prompt.txt",
    "examples/CustomerOrders/CustomerOrders.accdb.src/vcs-options.json",
    "examples/CustomerOrders/CustomerOrders.accdb.src/vcs-index.idx",  # VCS 5.0.1 导出为二进制索引（v4 为 vcs-index.json）
    "examples/CustomerOrders/CustomerOrders.accdb.src/forms/fCustomerList.bas",
    "examples/CustomerOrders/CustomerOrders.accdb.src/queries/qOrders.sql",
    "examples/CustomerOrders/CustomerOrders.accdb.src/reports/rOrderReport.bas",
    "examples/CustomerOrders/CustomerOrders.accdb.src/modules/modNavigation.bas",
    "examples/CustomerOrders/CustomerOrders.accdb.src/tbldefs/tOrders.xml",
    "examples/CustomerOrders/CustomerOrders.accdb.src/relations/tCustomerstOrders.json",
    "templates/standard-module-stub.bas",
    "templates/class-module-stub.cls",
    "templates/autoexec-macro.macro",
    "assets/manifest.json",
    "scripts/fwguard.py",
    "scripts/vcs_consistency_check.py",
    "scripts/rebuild_module_from_src.py",
    "scripts/install_vcs_addin.py",
    "tests/vcs_com_verify.py",
    "tests/vcs_mcp_verify.py",
]


def read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def has_bom_utf8(b):
    return b[:3] == codecs.BOM_UTF8


def collect_own_files():
    out = []
    for dirpath, dirnames, filenames in os.walk(SKILL_ROOT):
        rel = os.path.relpath(dirpath, SKILL_ROOT)
        top = rel.split(os.sep)[0] if rel != "." else ""
        if top in SKIP_DIRS:
            dirnames[:] = []
            continue
        for fn in filenames:
            out.append(os.path.join(dirpath, fn))
    return out


def main():
    failures = []
    checks = 0

    # 1) 子项目目录存在（部分已移入 examples/ 或 Edonsoft 框架文件夹内）
    for sp in SUBPROJECTS:
        checks += 1
        if sp in SUBPROJECTS_IN_EXAMPLES:
            p = os.path.join(SKILL_ROOT, "examples", sp)
        elif sp in SUBPROJECTS_IN_EDONSOFT:
            p = os.path.join(SKILL_ROOT, "Edonsoft Development Framework_x64", sp)
        else:
            p = os.path.join(SKILL_ROOT, sp)
        if not os.path.isdir(p):
            failures.append("子项目目录缺失: %s" % sp)

    # 2) 关键文件齐全
    for rf in REQUIRED_FILES:
        checks += 1
        p = os.path.join(SKILL_ROOT, rf)
        if not os.path.isfile(p):
            failures.append("关键文件缺失: %s" % rf)

    # 3) 模板编码：标准模块 / 类模块 = UTF-8-BOM（与 .src 一致，导入前由脚本转 GBK 无 BOM）
    for fn in ("standard-module-stub.bas", "class-module-stub.cls"):
        p = os.path.join(SKILL_ROOT, "templates", fn)
        if os.path.isfile(p):
            checks += 1
            if not has_bom_utf8(read_bytes(p)):
                failures.append("%s 应为 UTF-8-BOM（导入前由脚本转 GBK 无 BOM）" % fn)

    mac = os.path.join(SKILL_ROOT, "templates", "autoexec-macro.macro")
    if os.path.isfile(mac):
        checks += 1
        if has_bom_utf8(read_bytes(mac)):
            failures.append("autoexec-macro.macro 不应带 BOM")

    # 4) 无框架专有残留（changelog.md 除外：它是技能内唯一记载变化的地方，允许提及历史上的残留词）
    own = collect_own_files()
    for fp in own:
        if os.path.basename(fp) == "changelog.md":
            continue
        try:
            with open(fp, "r", encoding="utf-8") as f:
                text = f.read()
        except Exception:
            continue
        for tok in FORBIDDEN_TOKENS:
            if tok in TOKEN_RE:
                hit = TOKEN_RE[tok].search(text)
            else:
                hit = (tok in text)
            if hit:
                rel = os.path.relpath(fp, SKILL_ROOT)
                failures.append("残留 '%s' 出现在 %s" % (tok, rel))

    # 报告
    print("=" * 60)
    print("技能完整性冒烟测试")
    print("扫描根: %s" % SKILL_ROOT)
    print("检查项: %d" % checks)
    if failures:
        print("-" * 60)
        print("失败 %d 项：" % len(failures))
        for i, f in enumerate(failures, 1):
            print("  %d. %s" % (i, f))
        print("=" * 60)
        return 1
    print("结果: 全部通过 (%d 项)" % checks)
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
