# Access VBA 导入导出编码铁律

> 一句话：**读的时候用 UTF-8 带 BOM 存 `.src`；写进 Access 的时候必须是 GBK 无 BOM。**
> 编码错了不会报"编码错误"，而是报**看不懂的业务错误**（SQL 语法错、类型不对、模块类型退化）。

## 一、按场景对照表

| 场景 | 正确编码 | 错误后果 |
|---|---|---|
| `.src` 源码落盘（真相源） | UTF-8 **带 BOM** + CRLF | 无（可读性最好） |
| `VBComponents.Import(.bas)` | **GBK 无 BOM** + CRLF | 中文变乱码 `鐢?`；含中文的 SQL 字符串被破坏 → 运行时报 `3075 语法错误(操作符丢失)` |
| `VBComponents.Import(.cls)` | **GBK 无 BOM** + CRLF | 带 BOM 时首行 `VERSION 1.0 CLASS` 不被识别 → 建成 **标准模块 type=1**；`Friend` 成员直接编译报错 |
| `Application.LoadFromText(5, 模块)` | **GBK 无 BOM** | 首行变 `锘緼ttribute` → 工程瘫痪 |
| `Application.LoadFromText(4, 宏)` | **GBK 无 BOM** | `-2146826003 不能创建输出文件` |
| `Application.LoadFromText(2, 窗体)` | UTF-16 **LE + BOM** + CRLF | 导入失败或窗体错乱 |
| `Application.SaveAsText`（导出） | 由 Access 决定，勿手工改 | — |

## 二、判定"是不是编码问题"的三个特征

1. 中文变成 `鐢?`、`鎺㈤拡` 这类**生僻字组合**（UTF-8 字节被当 GBK 解码）；
2. 报错出现在**运行时**而不是导入时，且报的是 **SQL 语法/查询表达式**错误；
3. 代码逻辑本身没改过，但"以前能跑"的测试突然失败。

## 三、唯一可靠的验证方式：读回逐行比对

导入后**必须**读回来比，别只看"Import 没报错"：

```python
live = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
# 与源文件逐行比较；.cls 需剥掉 VERSION/BEGIN/END/Attribute 头
```

不一致时打印**首个差异行**的源与库两侧内容，一眼就能看出是乱码还是漏行。

## 四、BOM 的三重危害（务必记住）

1. `.bas` 首行 `Attribute VB_Name` 前多出 BOM → 解析异常；
2. `.cls` 的 `VERSION 1.0 CLASS` 头不被识别 → **类模块退化为标准模块**；
3. 宏文本 `Version =196611` 首行被污染 → `LoadFromText` 直接失败。

**注意**：某个 `.cls` 曾"用 UTF-8 无 BOM 导入成功且 type=2"，是因为它**一个非 ASCII 字符都没有**（整份是英文注释）。一旦类里有中文，同样的写法就会乱码——不要被这种偶发成功误导。

## 五、编码自检（写文件前）

```python
t = open(src, "rb").read().decode("utf-8-sig")          # 读真相源
t = t.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")  # 统一 CRLF
open(dst, "wb").write(t.encode("gbk"))                  # 关键：GBK 无 BOM
```

若源文件确实含 GBK 无法表示的字符（极少），导入会抛异常，此时再评估是否改用 UTF-16 路径，不要默认用 UTF-8。
