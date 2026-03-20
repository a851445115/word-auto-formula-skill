# PDF to Markdown

This skill supports two PDF extraction paths.

## 1. Marker

Use Marker by default for:

- Chinese or mixed Chinese/English papers
- scanned PDFs
- formula-heavy academic PDFs
- cases where image extraction and layout robustness matter more than speed

Command:

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
$PDF_PYTHON = "D:\anaconda3\envs\pdf-extractor\python.exe"  # adjust to your machine
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  -o "C:\path\to\paper_output\paper.md"
```

Useful variants:

```powershell
# Page-0 smoke test
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --page-range "0" `
  -o "C:\path\to\paper_output\page_01.md"
```

```powershell
# Force OCR
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --force-ocr `
  -o "C:\path\to\paper_output\paper.md"
```

```powershell
# Use current Codex-compatible provider settings for LLM enhancement
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2md_marker.py" `
  "C:\path\to\paper.pdf" `
  --codex-gpt-5-2 `
  -o "C:\path\to\paper_output\paper.llm.md"
```

## 2. Nougat

Use Nougat when:

- the paper is English-only or English-first
- you want a faster OCR path
- you do not need Marker's stronger multilingual handling

Command:

```powershell
$SKILL_DIR = "C:\path\to\document-formula-bridge"
$PDF_PYTHON = "D:\anaconda3\envs\pdf-extractor\python.exe"  # adjust to your machine
& $PDF_PYTHON "$SKILL_DIR\scripts\pdf2latex.py" `
  "C:\path\to\paper.pdf" `
  -o "C:\path\to\paper_output\paper.mmd"
```

## Preparation From DOCX Sources

If the source material is a `.docx` before it becomes a PDF, check whether it contains many MathType or embedded OLE equations.

Use this rule:

- For MathType-heavy Word files, prefer Microsoft Word native export to PDF before running Marker or Nougat.
- Do not prefer LibreOffice, generic DOCX-to-PDF converters, or online conversion services for MathType-heavy files.

## Output Layout

Prefer one dedicated output folder per PDF:

```text
paper_output/
  paper.md
  paper_images/
```

This keeps the final `.md` file and image assets together.

## Guardrails

- Do not start installing packages during extraction. Use an existing Python environment that already contains `marker-pdf` or `nougat-ocr`.
- Marker is slower but usually more robust for mixed-language papers and formula-heavy pages.
- If Marker LLM mode fails with `401 INVALID_API_KEY`, rerun with explicit `--openai-base-url`, `--openai-api-key`, and `--openai-model`, or fall back to non-LLM mode.
- If you split a PDF by page range, keep every partial `.md` and image directory inside the same dedicated output folder.
