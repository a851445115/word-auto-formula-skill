---
name: document-formula-bridge
description: Use when working on Windows with Word, MathType, PDF, DOCX, Markdown, and LaTeX files, especially when formulas must move between these formats without losing editability or visual fidelity.
---

# Document Formula Bridge

Use this skill on Windows when formulas must move safely between PDF, DOCX, Markdown, and LaTeX. It combines Word + MathType automation with academic PDF extraction.

## Quick Start

Main scripts:

- DOCX LaTeX -> MathType: `scripts/convert-docx-latex-to-formulas.ps1`
- DOCX audit: `scripts/audit-docx-formulas.ps1`
- DOCX MathType OLE -> Markdown: `scripts/export-docx-to-md.ps1`
- DOCX MathType OLE -> editable TeX DOCX: `scripts/convert-docx-mathtype-to-latex.ps1`
- PDF -> Markdown with Marker: `scripts/pdf2md_marker.py`
- PDF -> Markdown/LaTeX with Nougat: `scripts/pdf2latex.py`
- Marker OpenAI-compatible helper: `scripts/marker_openai_compat_service.py`
- DOCX export notes: `references/docx-to-markdown.md`
- PDF extraction notes: `references/pdf-to-markdown.md`
- Word/MathType troubleshooting: `references/troubleshooting.md`

## Supported Workflows

## 1. LaTeX text in DOCX -> MathType formulas

Use this when a `.docx` contains `$...$` or `$$...$$` text that should become native MathType objects.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx"
```

Optional commands:

```powershell
# Best-effort residual cleanup
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -AggressiveCleanup
```

```powershell
# Preflight only
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -PreflightOnly
```

## 2. DOCX(MathType OLE) -> Markdown with formula images

Use `formula-preserved` when formulas must stay visually exact in Markdown.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode formula-preserved
```

Default outputs:

- `{basename}_formula-preserved.md`
- `{basename}_formula-preserved_assets\`

Requirements:

- PowerShell
- Python 3

## 3. DOCX(MathType OLE) -> editable TeX DOCX copy -> Markdown with raw TeX

Use `latex-raw` when the formulas should become editable TeX text in a copied `.docx` and in the exported Markdown.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode latex-raw
```

Default outputs:

- `{basename}_latex.docx`
- `{basename}_latex_raw.md`
- `{basename}_latex_raw_assets\`

Requirements:

- Microsoft Word
- MathType
- Python 3

## 4. Direct MathType OLE -> TeX conversion on a working copy

Use this lower-level helper when only the copied `.docx` is needed.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-mathtype-to-latex.ps1" `
  -SourcePath "C:\path\to\document.docx" `
  -DestinationPath "C:\path\to\document_latex.docx"
```

## 5. Academic PDF -> Markdown with formulas using Marker

Use Marker by default for Chinese + English papers, scanned PDFs, or formula-heavy academic PDFs.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
$PDF_PYTHON = "D:\anaconda3\envs\pdf-extractor\python.exe"  # adjust to your machine
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  -o "C:\path\to\paper_output\paper.md"
```

Optional commands:

```powershell
# Page-0 smoke test
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --page-range "0" `
  -o "C:\path\to\paper_output\page_01.md"
```

```powershell
# Force OCR for scanned PDFs
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --force-ocr `
  -o "C:\path\to\paper_output\paper.md"
```

```powershell
# Use current Codex-compatible provider settings for Marker LLM mode
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --codex-gpt-5-2 `
  -o "C:\path\to\paper_output\paper.llm.md"
```

## 6. English-first academic PDF -> Markdown/LaTeX with Nougat

Use Nougat when the paper is English-only or English-first and you want a faster OCR path.

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
$PDF_PYTHON = "D:\anaconda3\envs\pdf-extractor\python.exe"  # adjust to your machine
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2latex.py" `
  "C:\path\to\paper.pdf" `
  -o "C:\path\to\paper_output\paper.mmd"
```

## Guardrails

- Do not use WPS as a substitute for Microsoft Word when Word COM automation is required.
- Keep the original file untouched unless the script explicitly documents in-place behavior.
- Treat `formula-preserved` and `latex-raw` as different outputs for different needs. One prioritizes visual fidelity; the other prioritizes editability.
- When a document carries revision meaning through red text, keep those color cues in the Markdown export.
- For MathType-heavy `.docx` sources that will later be extracted as PDFs, prefer Microsoft Word native PDF export before running Marker or Nougat.
- Prefer Marker for Chinese or mixed-language papers, scanned PDFs, and complex math-heavy layouts.
- Prefer Nougat only for English-first PDFs when speed matters more than multilingual robustness.
- Keep PDF extraction self-contained. Use an existing Python environment with `marker-pdf` or `nougat-ocr`; do not start installing packages mid-task.
- For PDF output, prefer one dedicated output folder per source PDF so the `.md` file and image assets stay together.

## Known Good Baseline

- `Word.Application` should identify itself as `Microsoft Word`, not WPS.
- A validated MathType template path on this machine was `C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm`.
- Forward bulk conversion worked reliably through reflection-based `InvokeMember('Run', ...)`.
- Reverse OLE-to-TeX extraction worked reliably by selecting `Equation.DSMT4` inline shapes and running `MathTypeCommands.UILib.MTCommand_TeXToggle`.
- A validated Windows PDF extraction environment on this machine was `D:\anaconda3\envs\pdf-extractor\python.exe`.

Read `references/docx-to-markdown.md`, `references/pdf-to-markdown.md`, and `references/troubleshooting.md` when the environment drifts away from that baseline.
