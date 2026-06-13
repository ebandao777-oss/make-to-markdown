---
name: make-to-markdown
description: 工业级RAG Markdown物料生成技能，使用 markitdown 将各类文档和文件转换为 Markdown 格式。支持 .doc/.ppt 老格式自动预处理（Word/PowerPoint COM / LibreOffice）。启动时自动检测 OS/版本/能力，按平台选择最佳执行路径。触发词：转 Markdown / 转换文档 / markitdown / 文档转 md / 批量转换。当需要将 PDF、Word (.docx/.doc)、PowerPoint (.pptx/.ppt)、Excel (.xlsx, .xls)、HTML、CSV、JSON、XML、图片（含 EXIF/OCR）、音频（含语音转写）、ZIP 压缩包、YouTube 链接或 EPub 电子书转换为 Markdown 格式，为知识库提供统一的"通用语言"时触发此技能。
version: "3.6"
metadata:
  domain: "make-to-markdown"
  author: "智慧半岛"
  platform:
    windows: full
    linux: full
    macos: full
  openclaw:
    requires:
      bins:
        - uv
        - python
    emoji: "📚"
---



> **分级导航**：L1=摘要（本段）→ L2=红线+流程+命令速查（§⛔/§1-§3）→ L3=完整运维细节（[REFERENCE.md](REFERENCE.md)）

## ⛔ 核心红线 (Critical Constraints)

1. **源文件零修改**：转换过程只读操作。所有中间产物写入临时目录，完成后不保留可关联回源文件的痕迹。
2. **输出覆盖确认**：`-o` 输出文件已存在时必须暂停确认，禁止静默覆写。
3. **加密文档不重试**：检测到加密（`File is encrypted`）立即终止，提示用户解密后重试。
4. **旧格式无预处理环境阻断**：`.doc`/`.ppt` 且 COM 和 LibreOffice 均不可用时，暂停告知用户安装 Office 或 LibreOffice，禁止静默跳过。
5. **禁止绕过 convert.py**：所有转换必须经由 `scripts/convert.py` 统一入口。禁止裸调 `markitdown` / `uvx markitdown` / 原生库。
6. **禁止外部 API 泄露**：转换全程本地执行，禁止上传文档内容至任何外部 API。
7. **网络驱动器 + 执行器约束**：Windows 网络驱动器（如 `E:\Marvis_Data`）路径优先用 `python_executor`，脚本内部 `pathlib.Path` 自动适配平台。
8. **系统破坏操作禁令 (HARB)**：`rm -rf` / `diskpart clean` / `DROP TABLE` / `git reset --hard` 等破坏性命令不得自动执行，必须输出完整预览并等待用户确认。
9. **输出 UTF-8 + 批量确认 + 禁循环清洗**：输出必须 UTF-8 无乱码；批量转换前确认源/目标路径+数量+格式；禁止 `post_clean.py`↔`convert.py` 循环清洗。

> 完整禁令细则及 HARB 黑名单详见 [REFERENCE.md §4](REFERENCE.md)。

# make-to-markdown 智能转换器 (v3.6)

核心入口：`scripts/convert.py`。零人工干预：`依赖检测 → uv 安装 extras → 转换 → 降级兜底 → 后置清洗 → 输出`。

🟢 **最小可用示例**：`python scripts/convert.py input.docx`。版本自查：`python scripts/convert.py --version`。

## 1. 核心执行流程

只需调用一次 `scripts/convert.py`，脚本内部自动完成：

```
(.doc/.ppt?)→预处理为 .docx/.pptx → 依赖检测 → 自动安装 → markitdown → (失败)原生降级 → 后置清洗 → 输出
```

| 步骤 | 说明 |
|------|------|
| 预处理 | `.doc`/`.ppt` 自动转为 `.docx`/`.pptx`（见 §2.5） |
| 依赖 | 按扩展名自动检测+安装缺失 Python 模块，120s 超时 |
| markitdown extras | 自动补全可选依赖，pypdf 已就绪则跳过 |
| markitdown 转换 | CLI 调用 |
| 降级兜底 | 失败时自动切换 python-docx/openpyxl/python-pptx |
| 后置清洗 | 内联执行：去水印+页码+标题修复+表格补全+摘要注入 |
| 结果反馈 | 输出路径 + H 标题统计 + 表格计数 + 转换方式 |

## 2. 依赖速查

`convert.py` 自动处理依赖。以下为手动参考，完整映射表见 [REFERENCE.md §1](REFERENCE.md)。

| 格式 | 额外包 |
|------|--------|
| .docx/.doc | python-docx |
| .xlsx | openpyxl |
| .xls | xlrd |
| .pptx/.ppt | python-pptx |
| .pdf | pypdf |
| .epub | ebooklib |

无需额外依赖的格式：`.html` `.csv` `.json` `.xml` `.png` `.jpg` `.jpeg` `.gif` `.bmp` `.tiff` `.webp` `.mp3` `.wav` `.m4a` `.ogg` `.flac` `.zip`

```bash
# 一次性安装全部可选依赖（脚本自动处理）
uv tool install markitdown --with python-docx --with openpyxl --with python-pptx --with pypdf --with xlrd --with ebooklib
```

## 2.5 旧格式预处理

`.doc`/`.ppt` 不被 markitdown 直接支持，`convert.py` 自动尝试两种方式：

| | `.doc`→`.docx` | `.ppt`→`.pptx` |
|---|---|---|
| 方式1 | Word COM (Windows only) | PowerPoint COM (Windows only) |
| 方式2 | `soffice --headless --convert-to docx` | `soffice --headless --convert-to pptx` |

**LibreOffice 定位链**：`shutil.which("soffice")` → `$SOFFICE_PATH` → 常见安装路径 → 裸名兜底。详见 [REFERENCE.md §3](REFERENCE.md)。

> 🔴 **CHECKPOINT**：旧格式且 COM 和 LibreOffice 均不可用时，暂停并告知用户安装 Office 或 LibreOffice。`.xls` 无需预处理。

## 3. 命令速查

```bash
# 单文件（推荐）
python scripts/convert.py input.docx -o output.md

# 默认输出名 / 跳过摘要 / 静默
python scripts/convert.py input.pdf
python scripts/convert.py data.xlsx --no-summary
python scripts/convert.py report.pptx -q

# 批量转换（确认源/目标/数量/格式后执行）
python scripts/batch_convert.py <源目录> <输出目录> --ext .pdf .docx --clean

# 单独后置清洗（仅极端场景）
python scripts/post_clean.py output.md --check-only
```

> 🔴 **CHECKPOINT**：`-o` 输出已存在时暂停确认；批量转换前必须确认路径+数量+格式。

## 3.2 降级转换器

markitdown 失败时自动启用原生降级（输出同样经过后置清洗）：

| 格式 | 降级方案 | 能力 |
|------|---------|------|
| .docx/.doc | python-docx | 段落样式→标题、表格→MD 表格 |
| .xlsx/.xls | openpyxl | Sheet→H2 章节、数据行→MD 表格 |
| .pptx | python-pptx | 幻灯片→H2 章节、表格自动转换 |

## 4. 后置清洗

`convert.py` 内联清洗管线（6 类 15+ 正则）：去水印/页码/机密标记/版权声明 + 空白压缩 + 标题层级修复 + 表格分隔符补全 + 文档摘要注入。完整清洗项及正则模式详见 [REFERENCE.md §1](REFERENCE.md)。

## 5. 异常处理

| 错误类型 | 处理方式 |
|:---|:---|
| 依赖安装失败 | 终止，输出缺失包名，提示手动安装 |
| markitdown 失败 | 自动降级到原生转换器 |
| 加密文档 | 🔴 立即终止，提示解密后重试 |
| 旧格式无预处理环境 | 🔴 CHECKPOINT：暂停，告知安装 Office/LibreOffice |
| 批量单文件失败 | 跳过，记入 `_conversion_errors.log`，其余继续 |
| 输出路径不可写 | 终止，提示检查权限/磁盘空间 |

> 完整异常分类表（含 uv 不在 PATH、网络驱动器不可达等）见 [REFERENCE.md §10](REFERENCE.md)。

## 6. 输出反馈模板

```
{状态}将 `{源文件名}` 转换为 Markdown [{转换方式}] | H1={n} H2={n} H3={n} | 表格={n} | {文件大小}
输出: {绝对路径}
```

| 字段 | 取值 |
|:---|:---|
| 状态 | `成功` / `失败` / `部分成功` |
| 转换方式 | `markitdown` / `原生降级(python-docx)` / `原生降级(openpyxl)` / `原生降级(python-pptx)` |

批量：`批量转换完成 | 总计=N | 成功=N | 失败=N`，详情见 `_conversion_errors.log`。

## 7. 脚本清单

| 脚本 | 用途 |
|------|------|
| `scripts/convert.py` | 智能转换引擎（依赖补全 + markitdown + 降级 + 内联清洗） |
| `scripts/batch_convert.py` | 批量递归转换，保持目录结构 |
| `scripts/post_clean.py` | 单独后置清洗（极端场景备用） |
| `scripts/platform_detect.py` | 平台检测模块，`best_office_tool()` 路由 |

前置校验：`python -c "from pathlib import Path; assert Path('scripts/convert.py').exists()"`

## 8. 平台兼容性

Windows / Linux / macOS 全支持。`convert.py` 启动时自动检测平台能力，脚本使用 `pathlib.Path` 适配路径。详见 [REFERENCE.md §3](REFERENCE.md)。

## 9. 反模式禁令

1. **裸调 markitdown** → 必须 `python scripts/convert.py`
2. **shell_executor 执行网络驱动器脚本** → 用 `python_executor`
3. **加密文档反复重试** → 首次失败即终止
4. **旧格式无预处理时强行转换** → 告知用户安装 Office/LibreOffice
5. **批量转换不设 `--ext`** → 必须指定格式过滤
6. **重复清洗 convert.py 输出** → 禁止循环清洗

详细禁令清单见 [REFERENCE.md §4](REFERENCE.md)。

## 10. 验证检查表

转换完成后必须执行逐项检查，任一未通过即进入 §5 异常处理。

| # | 检查项 | 验证命令（替换 `<输出.md>`） |
|:---|:---|:---|
| V1 | 文件存在且非空 | `python -c "import os; assert os.stat('<输出.md>').st_size>0"` |
| V2 | UTF-8 无乱码 | `python -c "open('<输出.md>',encoding='utf-8').read()"` |
| V3 | 无残留水印 | `python -c "import re; t=open('<输出.md>').read(); assert not re.search(r'Generated by\|Page \\d+',t)"` |
| V4 | 标题层级正常 | H1 存在 + 无跳跃（H1→H3 无 H2） |
| V5 | 表格分隔符完整 | 无 `\|---\|` 缺失 |
| V6 | 批量文件数对应 | 输出目录文件数 = 源文件数 |

Agent 完成后在最终回复中确认：V1 存在非空 / V2+V3 抽检前 3 段无乱码水印 / V4 无跳跃 / V5 表格完整 / V6（批量时）数量一致。

> 详细验证 oneliner、回归测试用例（RT1-RT3）见 [REFERENCE.md §5](REFERENCE.md)。自检清单见 [REFERENCE.md §2](REFERENCE.md)。
*（内容由AI生成，仅供参考）*
