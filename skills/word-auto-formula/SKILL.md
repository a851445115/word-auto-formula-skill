---
name: word-auto-formula
description: Use when working on Windows with Word, MathType, DOCX, Markdown, and LaTeX files, especially when formulas must be batch-converted into MathType, exported from MathType OLE into editable TeX, or preserved while extracting a DOCX document to Markdown.
---

# Word Auto Formula

Use this skill on Windows when formulas must move safely between Word, Markdown, and LaTeX.

## Quick Start

Main scripts:

- forward conversion: `scripts/convert-docx-latex-to-formulas.ps1`
- forward audit: `scripts/audit-docx-formulas.ps1`
- reverse export wrapper: `scripts/export-docx-to-md.ps1`
- reverse TeX helper: `scripts/convert-docx-mathtype-to-latex.ps1`
- troubleshooting: `references/troubleshooting.md`
- reverse export notes: `references/docx-to-markdown.md`

## Supported Workflows

## 1. LaTeX text in DOCX -> MathType formulas

Use this when a `.docx` contains `$...$` or `$$...$$` text that should become native MathType objects.

```powershell
$SKILL_DIR = "C:\path\to\word-auto-formula"
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
$SKILL_DIR = "C:\path\to\word-auto-formula"
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
$SKILL_DIR = "C:\path\to\word-auto-formula"
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
$SKILL_DIR = "C:\path\to\word-auto-formula"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-mathtype-to-latex.ps1" `
  -SourcePath "C:\path\to\document.docx" `
  -DestinationPath "C:\path\to\document_latex.docx"
```

## Guardrails

- Do not use WPS as a substitute for Microsoft Word when Word COM automation is required.
- Keep the original file untouched unless the script explicitly documents in-place behavior.
- Treat `formula-preserved` and `latex-raw` as different outputs for different needs. One prioritizes visual fidelity; the other prioritizes editability.
- When a document carries revision meaning through red text, keep those color cues in the Markdown export.
- Fix styling drift separately from formula conversion logic.

## Known Good Baseline

- `Word.Application` should identify itself as `Microsoft Word`, not WPS.
- A validated MathType template path on this machine was `C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm`.
- Forward bulk conversion worked reliably through reflection-based `InvokeMember('Run', ...)`.
- Reverse OLE-to-TeX extraction worked reliably by selecting `Equation.DSMT4` inline shapes and running `MathTypeCommands.UILib.MTCommand_TeXToggle`.

Read `references/troubleshooting.md` when the environment drifts away from that baseline.
