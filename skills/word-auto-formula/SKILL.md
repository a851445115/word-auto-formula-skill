---
name: word-auto-formula
description: Insert formulas into Microsoft Word documents on Windows by converting LaTeX-delimited equations into native MathType objects through Word COM automation and the MathType add-in. Use when a user wants to batch-replace `$...$` or `$$...$$` in `.docx` files, verify whether `Word.Application` is still hijacked by WPS, audit residual LaTeX after conversion, or rerun the MathType macro safely on a working copy of a document.
---

# Word Auto Formula

Use this skill on Windows only. Run the bundled PowerShell scripts instead of rebuilding the Word + MathType COM workflow from scratch.

## Quick Start

- Main conversion script: `scripts/convert-docx-latex-to-formulas.ps1`
- Audit script: `scripts/audit-docx-formulas.ps1`
- Troubleshooting guide: `references/troubleshooting.md`

```powershell
$SKILL_DIR = "C:\path\to\word-auto-formula"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx"
```

Default behavior:
- Work on a copy named `{filename}_mathtype.docx`
- Verify that `Word.Application` resolves to real Microsoft Word
- Load `MathType Commands 2016.dotm` from common install paths
- Run `MTCommand_OnTexToggle` over the whole document
- Save the converted document and audit the remaining visible LaTeX snippets
- Print MathType object counts plus sample leftovers for inspection

Optional cleanup:
- Add `-AggressiveCleanup` to attempt delimiter cleanup and targeted residual passes
- Treat that mode as best-effort: use it only when the stable conversion pass leaves obvious leftovers and the user accepts slower cleanup work

## Workflow

1. Resolve the input `.docx` path and default to a copy unless the user explicitly requests in-place edits.
2. Run `scripts/convert-docx-latex-to-formulas.ps1`.
3. If preflight reports that WPS still owns `Word.Application`, stop and follow `references/troubleshooting.md`.
4. If the final audit still shows leftover `$...$` or `$$...$$`, inspect the reported samples and retry with `-Visible` for live debugging before making manual fixes.
5. If the document contains tables encoded as plain text or LaTeX table blocks, keep or rebuild them as Word tables instead of forcing the whole block through MathType.

## Commands

```powershell
# Convert on a working copy
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx"

# Convert in place and keep Word visible during debugging
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -InPlace `
  -Visible
```

```powershell
# Enable best-effort residual cleanup after the stable bulk pass
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -AggressiveCleanup
```

```powershell
# Verify Word/MathType wiring without touching the document
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\convert-docx-latex-to-formulas.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -PreflightOnly
```

```powershell
# Audit an already-converted document
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\audit-docx-formulas.ps1" `
  -InputPath "C:\path\to\document_mathtype.docx"
```

## Guardrails

- Do not use WPS for this workflow. WPS can host MathType OLE objects but is not reliable for the Word automation path that actually converts TeX at scale.
- Keep the document closed before running the conversion script.
- Preserve the original file unless the user explicitly requests `-InPlace`.
- Treat residual snippets after the stable bulk pass as targeted cleanup work, not as proof that the conversion workflow failed.
- When the user reports formatting drift after conversion, fix Word styles separately from MathType conversion.

## Known Good Baseline

- `Word.Application` should identify itself as `Microsoft Word`, not WPS.
- A validated MathType template path on this machine was `C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm`.
- The macro invocation that worked reliably from PowerShell was reflection-based `InvokeMember('Run', ...)`, not a plain `Application.Run(...)` call.

Read `references/troubleshooting.md` when the environment drifts away from that baseline.
