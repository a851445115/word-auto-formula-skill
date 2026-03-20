# Troubleshooting

Read this file when:

- the forward conversion script cannot reach real Microsoft Word
- MathType does not load
- some LaTeX fragments remain after the bulk pass
- the reverse `latex-raw` export fails to convert `Equation.DSMT4` objects into editable TeX
- the reverse `formula-preserved` export produces unexpected asset or Markdown output

Default script behavior is intentionally conservative:
- The stable path runs the whole-document MathType toggle, saves, and audits the result.
- The optional `-AggressiveCleanup` mode attempts delimiter cleanup and targeted retries, but it is slower and more fragile.

## 1. `Word.Application` still resolves to WPS

Symptoms:
- `New-Object -ComObject Word.Application` opens WPS instead of Microsoft Word.
- The conversion script errors before opening the document.

Repair sequence:
1. Close WPS and Word completely.
2. Reassign `.doc`, `.docx`, and `.dotm` default apps to Microsoft Word in Windows settings.
3. Re-register Word COM from `Win + R` or PowerShell:

```powershell
& "C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE" /r
```

4. Launch Word once, then close it.
5. Verify the COM target:

```powershell
$word = New-Object -ComObject Word.Application
try {
  [pscustomobject]@{
    Name = $word.Name
    Version = $word.Version
    Path = $word.Path
  }
} finally {
  $word.Quit() | Out-Null
}
```

Expected result:
- `Name`: `Microsoft Word`
- `Path`: an Office directory such as `C:\Program Files (x86)\Microsoft Office\Root\Office16`

Do not continue until WPS is out of the COM path.

## 2. MathType template is missing

Common template paths:
- `C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm`
- `C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2016.dotm`
- `C:\Program Files\MathType\Office Support\32\MathType Commands 2016.dotm`
- `C:\Program Files\MathType\Office Support\64\MathType Commands 2016.dotm`

If the script cannot find the template automatically, rerun it with `-MathTypeTemplatePath`.

## 3. PowerShell cannot call the MathType macro directly

The plain `Application.Run(...)` pattern was unreliable from PowerShell because of COM ref-binding. The validated workaround is reflection:

```powershell
$word.GetType().InvokeMember(
  "Run",
  [System.Reflection.BindingFlags]::InvokeMethod,
  $null,
  $word,
  @("MTCommand_OnTexToggle", $null)
) | Out-Null
```

The bundled script already uses that form.

## 3a. Reverse `latex-raw` export uses a different MathType command

For reverse extraction from existing MathType OLE formulas to editable TeX text, the validated command in this session was:

```powershell
$word.Run("MathTypeCommands.UILib.MTCommand_TeXToggle")
```

That command is selection-based:

1. select the target `InlineShape`
2. run `MathTypeCommands.UILib.MTCommand_TeXToggle`

Do not replace the forward bulk-toggle command with this one. The two paths solve different problems.

## 4. Residual `$...$` or `$$...$$` remain after the bulk pass

What the script already does:
1. Whole-document `MTCommand_OnTexToggle`
2. Remove stray `$` markers around `Equation.DSMT4`
3. Rerun display-formula matches
4. Replace leftover inline fragments one by one through a temporary MathType-generated object

If leftovers remain:
1. Run the converter with `-Visible`.
2. Run `scripts/audit-docx-formulas.ps1` to collect sample snippets.
3. Replace the reported fragments manually or rerun on only that selection.

Known edge case:
- Fragments like `$15$` may require a MathType-safe wrapper such as `\mathrm{15}` before they convert cleanly.

## 5. Reverse export says `formula_count = 0` in `latex-raw` mode

That can be expected.

In `latex-raw` mode, the workflow first converts MathType OLE formulas into plain TeX text inside a copied `.docx`. After that conversion, the Markdown extractor no longer sees those formulas as OLE objects or preview images, so the exported Markdown summary may show zero preserved formulas while still containing raw TeX text.

Check the Markdown body itself before treating this as a failure.

## 6. `formula-preserved` export works but `latex-raw` export fails

That usually means:

- the `.docx` package can be read directly
- but Word COM or MathType is unavailable for the TeX-toggle step

Check:

1. Microsoft Word launches through `Word.Application`
2. MathType is installed
3. the selected formula objects report `ProgID = Equation.DSMT4`

If only Markdown extraction is needed, stay on `formula-preserved`.

## 7. Asset extraction produced images but formatting is not fully faithful

The reverse extractor is intentionally conservative:

- paragraphs and headings are exported in document order
- tables are emitted as HTML tables
- formula previews are preserved as images when still present
- text color is preserved when it matters for revision markup

It does not attempt complete Word layout reconstruction in Markdown.

## 8. Do not use WPS as a substitute host

WPS can sometimes host existing `Equation.DSMT4` objects, but it was not reliable for the insertion workflow:
- WPS hosting was able to expose some OLE data paths.
- `SetData(MathML)` failed with `DV_E_FORMATETC`.
- The larger blocker was WPS taking over `Word.Application`, which broke the stable automation route entirely.

Use Microsoft Word for conversion and reserve WPS, if needed, for viewing only.
