# 示例 02：增量回写——改文本源，原子重导入，断言生效

**一句话**：这是日常开发最常用的那一拍：在文本源 `.src` 上改一处逻辑，用版本控制把它原子重导入回库，再用 `Application.Run` 复核改动确实生效。对应 `SKILL.md` 阶段 3「导入也走自动化，不手工粘贴」与全阶段贯穿的「改 `.src` → 构建/回写 → 核验」闭环。

## 运行前提

- 同 `examples/01_blank_db_to_vcs_loop/README.md`：Access + VCS 加载项正式安装版 + pywin32。

## 运行

```bat
python run.py
```

## 它做了什么

1. 建空库、写初始模块 `basSample`（`Hello()` 返回初值 `"v1"`），用 `scripts/rebuild_module_from_src.py` 导入。
2. 用 Access `Application.Run("Hello")` 复核初值。
3. **直接在 `.src` 文本上改**：把返回值改成 `"v2"`（模拟 AI 代理或你在文本源里编辑）。
4. 再次调 `scripts/rebuild_module_from_src.py` 做原子重导入（Remove + Import + 读回比对）。
5. 再次 `Application.Run("Hello")` 复核，应返回 `"v2"`——证明「改文本源 → 回写」闭环成立，全程无手工粘贴。

## 你该看到的结果

- 第 2 步打印 `Hello() = v1`。
- 第 5 步打印 `Hello() = v2`，且第 4 步比对「一致」。

若第 5 步仍是 `v1`，说明回写没真正进库（典型是编码错导致乱码或 `Import` 退化），此时回到 `references/vba-encoding-rules.md` 与 `references/vcs-three-channels.md` 排查。

## 与技能文档的对应

- 回写走原子重建而非手工粘贴，属阶段 3：读 `SKILL.md`。
- 编码要求（`.src` UTF-8-BOM、导入 GBK 无 BOM + CRLF）：读 `references/vba-encoding-rules.md`。
