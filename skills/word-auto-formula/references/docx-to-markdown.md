# DOCX to Markdown

This skill supports two reverse-export modes.

## 1. `formula-preserved`

Use this mode when the document contains MathType OLE formulas and you need the Markdown output to preserve the visible formulas exactly as rendered.

Command:

```powershell
$SKILL_DIR = "C:\path\to\word-auto-formula"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode formula-preserved
```

Output:

- `{basename}_formula-preserved.md`
- `{basename}_formula-preserved_assets\`

Requirements:

- PowerShell
- Python 3

This mode does not need Word COM just to read the `.docx` package.

## 2. `latex-raw`

Use this mode when the document contains MathType OLE formulas and you want an editable TeX representation in both a copied `.docx` and the exported Markdown.

Command:

```powershell
$SKILL_DIR = "C:\path\to\word-auto-formula"
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$SKILL_DIR\scripts\export-docx-to-md.ps1" `
  -InputPath "C:\path\to\document.docx" `
  -Mode latex-raw
```

Output:

- `{basename}_latex.docx`
- `{basename}_latex_raw.md`
- `{basename}_latex_raw_assets\`

Requirements:

- Microsoft Word
- MathType
- Python 3

This mode works on a copied document by default. The original input file is not modified.

## Notes

- The extractor reads the main document body. It does not attempt full-fidelity conversion of headers, footers, tracked-change metadata, comments, or every Word-specific layout feature.
- Visible red text is preserved as HTML color spans so revision cues survive in Markdown.
- Tables are emitted as HTML tables because that is safer than forcing lossy Markdown table reconstruction on merged or styled Word tables.
