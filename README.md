# word-auto-formula-skill

这个仓库当前对外发布的 Codex skill 名称是 `document-formula-bridge`。仓库名保留了旧名字，但 skill 本身已经扩展成一个统一的“公式文档桥接”工具，覆盖 PDF、DOCX、Markdown、LaTeX、Word、MathType 之间的转换与提取。

## 仓库当前包含的能力

- 把 DOCX 中的 `$...$` / `$$...$$` 批量转换为 MathType 公式对象
- 预检 Word COM 和 MathType 模板是否可用
- 审计转换后的残留 `$...$` / `$$...$$`
- 将包含 MathType OLE 的 DOCX 导出为 Markdown，并保留公式图片
- 将 MathType OLE 批量转换为可编辑 TeX，并导出 raw-TeX Markdown
- 将学术 PDF 提取为 Markdown，并保留或识别公式
- 提供 Marker 与 Nougat 两条 PDF 提取路径
- 提供 Word / MathType / WPS / PDF 提取环境的常见排障说明

实际内容以 [`skills/document-formula-bridge`](./skills/document-formula-bridge) 目录为准。

## 适用场景

- Word 文档中已经写了大量 `$...$` / `$$...$$`，需要批量转成可编辑公式
- 需要先确认 `Word.Application` 当前是不是被 WPS 劫持
- 需要审计一个已转换文档里还剩多少 LaTeX 片段
- 需要把含 MathType 公式的 DOCX 导出为 Markdown，并保留公式外观
- 需要把 Word 中的 MathType OLE 公式转成可编辑 TeX
- 需要把学术 PDF 提取成 Markdown，并尽量保住公式质量
- 需要在 Word / PDF / Markdown / LaTeX 之间复用统一工作流，而不是每次手工拼装

## 环境要求

- 通用：
  - Windows
  - PowerShell
- Word / MathType 工作流：
  - Microsoft Word 桌面版
  - MathType，且 Office Support 已正确安装
- DOCX -> Markdown 工作流：
  - Python 3
- PDF 提取工作流：
  - Python 环境，且已预装 `marker-pdf` 或 `nougat-ocr`

不建议使用 WPS 代替 Microsoft Word 执行 COM 自动化。WPS 可能能显示已有公式对象，但这套正向转换与部分反向工作流依赖的是 Word COM。

## 给 Codex 安装

这个仓库里的 skill 不在仓库根目录，而是在 `skills/document-formula-bridge/` 下。因此：

- 直接只发仓库根链接给 Codex，不够稳妥
- 推荐直接发 skill 子路径链接

推荐让同学直接对 Codex 这样说：

```text
请安装这个 skill：
https://github.com/a851445115/word-auto-formula-skill/tree/main/skills/document-formula-bridge
```

也可以这样说：

```text
请从仓库 a851445115/word-auto-formula-skill 安装 skill skills/document-formula-bridge
```

安装完成后，重启 Codex 以加载新 skill。

## 为什么不要只发仓库根链接

Codex 的 GitHub skill 安装器最终需要定位到一个包含 `SKILL.md` 的目录。这个仓库的 `SKILL.md` 位于：

```text
skills/document-formula-bridge/SKILL.md
```

所以只给下面这个根链接：

```text
https://github.com/a851445115/word-auto-formula-skill
```

安装器无法直接判断要安装哪个子目录。

## 使用方式

### 1. 让 Codex 直接调用这个 skill

安装后，可以直接在对话里描述需求，例如：

```text
请用 document-formula-bridge 把这个 docx 里的 $...$ 和 $$...$$ 公式批量转成 MathType，并先做一次 preflight。
```

或者：

```text
请用 document-formula-bridge 把这个 PDF 提取成 Markdown，优先保留公式质量。
```

### 2. 手动运行脚本

Word / MathType 正向转换：

```powershell
$SKILL_DIR = "$env:USERPROFILE\.codex\skills\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx"
```

DOCX 中公式保真导出为 Markdown：

```powershell
$SKILL_DIR = "$env:USERPROFILE\.codex\skills\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode formula-preserved
```

DOCX 中公式转 raw TeX：

```powershell
$SKILL_DIR = "$env:USERPROFILE\.codex\skills\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode latex-raw
```

PDF 提取为 Markdown：

```powershell
$SKILL_DIR = "$env:USERPROFILE\.codex\skills\document-formula-bridge"
$PDF_PYTHON = "D:\anaconda3\envs\pdf-extractor\python.exe"  # 按本机环境调整
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  -o "C:\path\to\paper_output\paper.md"
```

## 仓库结构

```text
word-auto-formula-skill/
├─ README.md
└─ skills/
   └─ document-formula-bridge/
      ├─ SKILL.md
      ├─ agents/
      │  └─ openai.yaml
      ├─ scripts/
      │  ├─ convert-docx-latex-to-formulas.ps1
      │  ├─ audit-docx-formulas.ps1
      │  ├─ export-docx-to-md.ps1
      │  ├─ convert-docx-mathtype-to-latex.ps1
      │  ├─ convert-docx-assets-to-png.ps1
      │  ├─ extract-docx-formula-preserved.py
      │  ├─ pdf2md_marker.py
      │  ├─ pdf2latex.py
      │  ├─ marker_openai_compat_service.py
      │  └─ marker_llm_antigravity.py
      └─ references/
         ├─ docx-to-markdown.md
         ├─ pdf-to-markdown.md
         └─ troubleshooting.md
```

## 常见问题

### 1. 运行时打开的是 WPS，不是 Word

先看 [`skills/document-formula-bridge/references/troubleshooting.md`](./skills/document-formula-bridge/references/troubleshooting.md)。核心是把 `Word.Application` 重新注册回 Microsoft Word。

### 2. 脚本提示找不到 MathType 模板

确认 MathType 的 Office Support 已安装；必要时运行脚本时显式传入 `-MathTypeTemplatePath`。

### 3. 为什么转换后还有少量 `$...$`

这通常是残留清理问题，不一定表示主流程失败。先用审计脚本定位样例，再决定是否加 `-AggressiveCleanup` 或手工修正。

### 4. 为什么导出的 Markdown 里公式是图片，不是 TeX

你用的是 `formula-preserved` 模式。这个模式优先保留公式外观。如果你要可编辑 TeX，请改用 `latex-raw`。

### 5. PDF 提取时应该优先用 Marker 还是 Nougat

中文、混排、复杂公式、扫描件，优先用 Marker。英文优先且更看重速度时，再考虑 Nougat。

### 6. 为什么 PDF 工作流还要求单独的 Python 环境

PDF 提取依赖 `marker-pdf` 或 `nougat-ocr`。这个 skill 只是把工作流和脚本整合到一个入口里，不会在运行时替你安装这些依赖。

## 说明

如果你打算把这份 skill 长期分享给同学，建议优先把 GitHub 仓库维护成“可直接安装、可直接看懂、能力描述和实际文件一致”的状态。对外发布时，README 和 [`skills/document-formula-bridge/SKILL.md`](./skills/document-formula-bridge/SKILL.md) 的能力范围最好保持同步。
