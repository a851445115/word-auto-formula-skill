# Word Auto Formula Bidirectional Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Extend `word-auto-formula` into a bidirectional Windows skill that keeps the existing LaTeX-to-MathType conversion path and adds `DOCX -> Markdown` export in `formula-preserved` and `latex-raw` modes.

**Architecture:** Add one high-level export wrapper and three low-level helpers. Reuse the already validated Word COM + MathType TeX-toggle path for `latex-raw`, and reuse direct package extraction for `formula-preserved`. Update `SKILL.md`, troubleshooting, and agent metadata so the skill is usable from Codex without reconstructing the workflow.

**Tech Stack:** PowerShell, Python 3, Microsoft Word COM, MathType, WordprocessingML XML, System.Drawing

---

### Task 1: Add the reverse-export wrapper

**Files:**
- Create: `skills/word-auto-formula/scripts/export-docx-to-md.ps1`
- Test: manual PowerShell smoke runs against local `.docx` inputs

**Step 1: Write the failing smoke command**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "skills/word-auto-formula/scripts/export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode formula-preserved
```

Expected before implementation:
- script file missing

**Step 2: Create the wrapper with minimal behavior**

Implement:

- path resolution
- mode validation: `formula-preserved` or `latex-raw`
- output naming
- temp unpack directory creation
- dispatch to helper scripts
- structured JSON summary

**Step 3: Run the wrapper in `formula-preserved` mode**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "skills/word-auto-formula/scripts/export-docx-to-md.ps1" `
  -InputPath "C:\Users\cr\Desktop\审稿意见修改\319\审稿意见回复_318.docx" `
  -Mode formula-preserved
```

Expected:
- Markdown file created
- asset directory created
- JSON summary reports non-zero output paths and counts

**Step 4: Run the wrapper in `latex-raw` mode**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "skills/word-auto-formula/scripts/export-docx-to-md.ps1" `
  -InputPath "C:\Users\cr\Desktop\审稿意见修改\319\审稿意见回复_318.docx" `
  -Mode latex-raw
```

Expected:
- copied `_latex.docx` created
- Markdown file created
- JSON summary includes converted formula counts

### Task 2: Add the MathType-to-TeX helper

**Files:**
- Create: `skills/word-auto-formula/scripts/convert-docx-mathtype-to-latex.ps1`
- Test: manual PowerShell smoke run against local `.docx` inputs

**Step 1: Write the failing smoke command**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "skills/word-auto-formula/scripts/convert-docx-mathtype-to-latex.ps1" `
  -SourcePath "C:\path\to\document.docx" `
  -DestinationPath "C:\path\to\document_latex.docx"
```

Expected before implementation:
- script file missing

**Step 2: Write minimal implementation**

Implement:

- copy source to destination
- open destination in Word COM
- iterate inline shapes backwards
- select only `Equation.DSMT4`
- run `MathTypeCommands.UILib.MTCommand_TeXToggle`
- save and return JSON counts

**Step 3: Run the helper**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "skills/word-auto-formula/scripts/convert-docx-mathtype-to-latex.ps1" `
  -SourcePath "C:\Users\cr\Desktop\审稿意见修改\319\26-0247_修改稿_红字_318.docx" `
  -DestinationPath "C:\Users\cr\Desktop\审稿意见修改\319\repo-smoke\26-0247_修改稿_红字_318_latex.docx"
```

Expected:
- JSON reports successful conversions
- destination document exists

### Task 3: Add the asset-conversion and XML extraction helpers

**Files:**
- Create: `skills/word-auto-formula/scripts/convert-docx-assets-to-png.ps1`
- Create: `skills/word-auto-formula/scripts/extract-docx-formula-preserved.py`
- Test: manual wrapper run and direct extractor run

**Step 1: Write the failing smoke command**

Run:

```powershell
python "skills/word-auto-formula/scripts/extract-docx-formula-preserved.py" --help
```

Expected before implementation:
- script file missing

**Step 2: Write minimal implementation**

Implement:

- read unpacked `word/document.xml` and relationships
- extract headings, paragraphs, tables, normal images, and MathType preview images
- preserve red font as HTML color spans
- emit Markdown
- convert extracted assets to PNG through the PowerShell helper

**Step 3: Run direct extraction smoke test**

Run through the wrapper or direct script on:

- `C:\Users\cr\Desktop\审稿意见修改\319\审稿意见回复_318.docx`
- `C:\Users\cr\Desktop\审稿意见修改\319\26-0247_修改稿_红字_318.docx`

Expected:
- Markdown emitted successfully
- formula-preserved path reports non-zero formula count
- manuscript export keeps red-font spans

### Task 4: Update the skill documentation and metadata

**Files:**
- Modify: `skills/word-auto-formula/SKILL.md`
- Modify: `skills/word-auto-formula/references/troubleshooting.md`
- Modify: `skills/word-auto-formula/agents/openai.yaml`
- Create: `skills/word-auto-formula/references/docx-to-markdown.md`

**Step 1: Write the failing documentation gap**

Current expected failure:
- `SKILL.md` documents only forward conversion
- no user-facing reverse export command exists in docs

**Step 2: Update docs minimally but completely**

Document:

- forward workflow
- `formula-preserved` export
- `latex-raw` export
- prerequisites per mode
- safe working-copy behavior
- troubleshooting for Word COM and MathType requirements

**Step 3: Verify docs against real commands**

Check:

- every documented command exists
- command names and parameters match the scripts exactly

### Task 5: Verify, install, commit, and push

**Files:**
- Modify: repo working tree
- Install to: `C:\Users\cr\.codex\skills\word-auto-formula`

**Step 1: Run end-to-end verification**

Run:

- `git diff --stat`
- forward preflight on existing script
- reverse export in both modes on both real documents

Expected:
- no missing-script errors
- outputs created successfully
- no original source file overwritten

**Step 2: Install the skill locally**

Copy:

- `skills/word-auto-formula` -> `C:\Users\cr\.codex\skills\word-auto-formula`

Expected:
- local Codex skills directory contains the updated skill

**Step 3: Commit**

Run:

```bash
git add skills/word-auto-formula docs/plans
git commit -m "feat: add bidirectional word formula export skill"
```

**Step 4: Push**

Run:

```bash
git push origin main
```

Expected:
- remote `main` includes the new reverse-export capability
