# Word Auto Formula Bidirectional Design

**Date:** 2026-03-19

**Status:** Approved scope, pending implementation

## Goal

Evolve `word-auto-formula` from a one-way `LaTeX text -> Word/MathType formulas` skill into a Windows-only bridge skill that also supports reverse export from Word `.docx` files to Markdown in two validated modes:

1. `formula-preserved`
   Export Word text to Markdown while preserving MathType OLE formulas as rendered images.
2. `latex-raw`
   Convert MathType OLE formulas inside a working-copy `.docx` into editable TeX text, then export the document to Markdown with raw LaTeX preserved in the text stream.

## Approved Scope

Keep the existing forward path unchanged:

- `LaTeX-delimited text in DOCX -> MathType formulas`

Add the reverse paths validated in this session:

- `DOCX(MathType OLE) -> Markdown with formula images`
- `DOCX(MathType OLE) -> DOCX copy with TeX text -> Markdown with raw TeX`

## Non-Goals

- No generic `Markdown -> DOCX` renderer in this iteration.
- No OCR or image-to-LaTeX recognition.
- No in-place modification of the original input document by default.
- No attempt to normalize all Word layout details into semantic Markdown.

## Baseline Findings

The current repository contains only:

- `scripts/convert-docx-latex-to-formulas.ps1`
- `scripts/audit-docx-formulas.ps1`
- `SKILL.md` describing forward conversion only

It does not yet contain:

- a user-facing `DOCX -> Markdown` workflow
- a `MathType OLE -> editable TeX` extraction path
- asset conversion for formula preview images
- documentation for reverse export modes

This baseline is the failing pre-change state for the new capability.

## Technical Findings Reused From This Session

- Math formulas in the manuscript are stored as `Equation.DSMT4` MathType OLE objects, not Word `m:oMath`.
- The reliable editable-TeX conversion path is:
  - open a `.docx` working copy in Word COM
  - iterate `InlineShapes`
  - select each `Equation.DSMT4`
  - run `MathTypeCommands.UILib.MTCommand_TeXToggle`
- This worked on real documents:
  - reply doc: `109/109` formulas converted
  - manuscript doc: `244/244` formulas converted, with `12` remaining inline shapes that were regular images rather than formulas
- For formula-preserved export, preview images embedded in the `.docx` package can be extracted and converted to PNG without changing the source document.

## Proposed Repository Changes

## Skill Surface

Keep the existing skill name `word-auto-formula`, but broaden the description and commands to cover both directions.

Primary user-facing commands after the change:

- forward conversion:
  - `scripts/convert-docx-latex-to-formulas.ps1`
- reverse export:
  - `scripts/export-docx-to-md.ps1 -Mode formula-preserved`
  - `scripts/export-docx-to-md.ps1 -Mode latex-raw`

## Internal Script Layout

Retain the current forward scripts and add:

- `scripts/export-docx-to-md.ps1`
  - high-level wrapper
  - accepts input `.docx`, output directory, export mode
  - unpacks the document into a temporary directory
  - dispatches to the appropriate low-level conversion path
- `scripts/convert-docx-mathtype-to-latex.ps1`
  - copies the input `.docx`
  - converts MathType OLE formulas to TeX text in the copy
  - returns counts and output path as structured JSON
- `scripts/convert-docx-assets-to-png.ps1`
  - converts extracted EMF/WMF/image assets to PNG for Markdown embedding
- `scripts/extract-docx-formula-preserved.py`
  - parses WordprocessingML
  - extracts text, tables, normal images, and MathType preview images
  - preserves red-font revisions as HTML color spans
  - emits Markdown

## Output Conventions

`formula-preserved` mode:

- produces `{basename}_formula-preserved.md`
- produces `{basename}_formula-preserved_assets/`

`latex-raw` mode:

- produces `{basename}_latex.docx`
- produces `{basename}_latex_raw.md`
- produces `{basename}_latex_raw_assets/`

## Dependency Model

`formula-preserved` mode:

- requires Python 3
- requires PowerShell
- does not require Word COM for extraction itself

`latex-raw` mode:

- requires Microsoft Word COM
- requires MathType
- requires Python 3
- must run on a working copy

## Guardrails

- Preserve the original input document unless the user explicitly asks otherwise.
- Keep the forward conversion workflow stable; do not regress existing behavior.
- In `latex-raw` mode, operate on a copied file only.
- Treat normal inline images and MathType preview images separately.
- Preserve visible red text in Markdown export because the manuscript workflow depends on revision color cues.

## Verification Plan

Use the real documents already validated in this session as smoke-test inputs:

- `审稿意见回复_318.docx`
- `26-0247_修改稿_红字_318.docx`

Verify:

- `formula-preserved` mode emits Markdown and PNG assets with non-zero formula counts
- `latex-raw` mode emits a copied `.docx`, converts OLE formulas to TeX text, and exports Markdown with raw TeX visible
- the updated `SKILL.md` accurately documents all supported paths and prerequisites

## Installation Outcome

After repository changes are verified:

- commit to `main`
- push to GitHub remote
- copy the finished skill folder into `C:\Users\cr\.codex\skills\word-auto-formula`
