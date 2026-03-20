from __future__ import annotations

import argparse
import json
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from html import escape
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

HEADING_LEVELS = {
    "Title": 1,
    "Subtitle": 2,
    "Heading1": 1,
    "Heading2": 2,
    "Heading3": 3,
    "Heading4": 4,
    "Heading5": 5,
    "Heading6": 6,
}


def qn(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def local_name(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def load_relationships(path: Path) -> dict[str, str]:
    tree = ET.parse(path)
    root = tree.getroot()
    rels: dict[str, str] = {}
    for rel in root.findall("pr:Relationship", NS):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            rels[rel_id] = target
    return rels


def parse_pt_size(style: str, key: str) -> str | None:
    match = re.search(rf"{re.escape(key)}\s*:\s*([0-9.]+pt)", style)
    return match.group(1) if match else None


def emu_to_pt(value: str | None) -> str | None:
    if not value:
        return None
    try:
        pt = int(value) / 12700.0
    except ValueError:
        return None
    return f"{pt:.2f}pt"


def normalize_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_color(color_val: str | None) -> str | None:
    if not color_val:
        return None

    value = color_val.lower()
    if value in {"auto", "000000"}:
        return None

    if re.fullmatch(r"[0-9a-f]{6}", value):
        channels = [int(value[idx : idx + 2], 16) for idx in (0, 2, 4)]
        if max(channels) < 96 and max(channels) - min(channels) < 24:
            return None
        return f"#{value}"

    return None


def format_text(text: str, style: dict[str, str | bool | None]) -> str:
    if not text:
        return ""

    value = escape(text)
    value = value.replace("\t", "&emsp;")

    if style.get("bold"):
        value = f"<strong>{value}</strong>"
    if style.get("italic"):
        value = f"<em>{value}</em>"

    vert = style.get("vert")
    if vert == "superscript":
        value = f"<sup>{value}</sup>"
    elif vert == "subscript":
        value = f"<sub>{value}</sub>"

    color = style.get("color")
    if color:
        value = f'<span style="color:{color}">{value}</span>'

    return value


@dataclass
class Context:
    source_dir: Path
    assets_dir: Path
    assets_rel: str
    converter_script: Path
    rels: dict[str, str]
    asset_lookup: dict[str, str] = field(default_factory=dict)
    manifest: list[dict[str, str]] = field(default_factory=list)
    formula_count: int = 0
    image_count: int = 0
    table_count: int = 0

    def ensure_asset(self, target: str, kind: str) -> str:
        source_path = (self.source_dir / "word" / Path(target)).resolve()
        key = str(source_path)
        if key not in self.asset_lookup:
            index = len(self.asset_lookup) + 1
            file_name = f"{kind}_{index:04d}.png"
            dest_path = (self.assets_dir / file_name).resolve()
            self.asset_lookup[key] = file_name
            self.manifest.append({"src": key, "dst": str(dest_path)})
        return f"{self.assets_rel}/{self.asset_lookup[key]}".replace("\\", "/")


def run_style(run: ET.Element) -> dict[str, str | bool | None]:
    style: dict[str, str | bool | None] = {
        "bold": False,
        "italic": False,
        "color": None,
        "vert": None,
    }
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return style

    style["bold"] = rpr.find("w:b", NS) is not None
    style["italic"] = rpr.find("w:i", NS) is not None

    color = rpr.find("w:color", NS)
    color_val = color.get(qn("w", "val")) if color is not None else None
    style["color"] = normalize_color(color_val)

    vert = rpr.find("w:vertAlign", NS)
    vert_val = vert.get(qn("w", "val")) if vert is not None else None
    if vert_val in {"superscript", "subscript"}:
        style["vert"] = vert_val

    return style


def build_img_tag(src: str, alt: str, width: str | None, height: str | None, inline: bool = True) -> str:
    style_parts: list[str] = []
    if width:
        style_parts.append(f"width:{width}")
    if height:
        style_parts.append(f"height:{height}")
    if inline:
        style_parts.append("vertical-align:middle")
    style_attr = f' style="{"; ".join(style_parts)}"' if style_parts else ""
    return f'<img src="{escape(src)}" alt="{escape(alt)}"{style_attr}/>'


def extract_from_object(obj: ET.Element, ctx: Context) -> str:
    shape = obj.find(".//v:shape", NS)
    imagedata = obj.find(".//v:imagedata", NS)
    ole = obj.find(".//o:OLEObject", NS)

    style = shape.get("style", "") if shape is not None else ""
    width = parse_pt_size(style, "width")
    height = parse_pt_size(style, "height")

    if ole is not None and ole.get("ProgID") == "Equation.DSMT4" and imagedata is not None:
        rel_id = imagedata.get(qn("r", "id"))
        target = ctx.rels.get(rel_id or "")
        if target:
            ctx.formula_count += 1
            asset = ctx.ensure_asset(target, "formula")
            return build_img_tag(asset, f"formula-{ctx.formula_count:04d}", width, height)

    if imagedata is not None:
        rel_id = imagedata.get(qn("r", "id"))
        target = ctx.rels.get(rel_id or "")
        if target:
            ctx.image_count += 1
            asset = ctx.ensure_asset(target, "image")
            return build_img_tag(asset, f"image-{ctx.image_count:04d}", width, height)

    return ""


def extract_from_drawing(drawing: ET.Element, ctx: Context) -> str:
    blip = drawing.find(".//a:blip", NS)
    if blip is None:
        return ""

    rel_id = blip.get(qn("r", "embed"))
    target = ctx.rels.get(rel_id or "")
    if not target:
        return ""

    extent = drawing.find(".//wp:extent", NS)
    width = emu_to_pt(extent.get("cx")) if extent is not None else None
    height = emu_to_pt(extent.get("cy")) if extent is not None else None

    ctx.image_count += 1
    asset = ctx.ensure_asset(target, "image")
    return build_img_tag(asset, f"image-{ctx.image_count:04d}", width, height)


def extract_inline(node: ET.Element, ctx: Context) -> str:
    tag = local_name(node.tag)

    if tag == "t":
        return node.text or ""
    if tag == "tab":
        return "\t"
    if tag in {"br", "cr"}:
        return "<br/>"
    if tag in {"noBreakHyphen", "softHyphen", "hyphen"}:
        return "-"
    if tag == "sym":
        char = node.get(qn("w", "char")) or node.get("char")
        if char:
            try:
                return chr(int(char, 16))
            except ValueError:
                return ""
        return ""
    if tag == "object":
        return extract_from_object(node, ctx)
    if tag == "pict":
        return extract_from_object(node, ctx)
    if tag == "drawing":
        return extract_from_drawing(node, ctx)
    if tag in {"footnoteReference", "endnoteReference"}:
        note_id = node.get(qn("w", "id")) or node.get("id")
        return f"[^{note_id}]" if note_id else ""
    if tag == "instrText":
        return ""

    parts: list[str] = []
    for child in list(node):
        parts.append(extract_inline(child, ctx))
    return "".join(parts)


def extract_run(run: ET.Element, ctx: Context) -> str:
    style = run_style(run)
    parts: list[str] = []
    for child in list(run):
        if local_name(child.tag) == "rPr":
            continue
        fragment = extract_inline(child, ctx)
        if not fragment:
            continue
        if local_name(child.tag) in {"t", "tab", "sym", "noBreakHyphen", "softHyphen", "hyphen"}:
            parts.append(format_text(fragment, style))
        else:
            parts.append(fragment)
    return "".join(parts)


def paragraph_style(paragraph: ET.Element) -> str | None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.get(qn("w", "val"))


def render_paragraph(paragraph: ET.Element, ctx: Context, in_table: bool = False) -> str:
    style_id = paragraph_style(paragraph)
    pieces: list[str] = []

    for child in list(paragraph):
        tag = local_name(child.tag)
        if tag == "r":
            pieces.append(extract_run(child, ctx))
        elif tag in {"hyperlink", "smartTag", "sdt", "ins", "moveTo"}:
            for grandchild in list(child):
                if local_name(grandchild.tag) == "r":
                    pieces.append(extract_run(grandchild, ctx))
                else:
                    pieces.append(extract_inline(grandchild, ctx))
        elif tag in {"del", "moveFrom"}:
            continue
        elif tag == "pPr":
            continue
        else:
            pieces.append(extract_inline(child, ctx))

    content = normalize_whitespace("".join(pieces))
    if not content:
        return ""

    if in_table:
        return content

    if style_id in HEADING_LEVELS:
        return f'{"#" * HEADING_LEVELS[style_id]} {content}'
    return content


def render_table(table: ET.Element, ctx: Context) -> str:
    ctx.table_count += 1
    rows: list[str] = ["<table>"]
    for row in table.findall("w:tr", NS):
        rows.append("<tr>")
        for cell in row.findall("w:tc", NS):
            fragments: list[str] = []
            for child in list(cell):
                tag = local_name(child.tag)
                if tag == "p":
                    para = render_paragraph(child, ctx, in_table=True)
                    if para:
                        fragments.append(para)
                elif tag == "tbl":
                    fragments.append(render_table(child, ctx))
            rows.append(f"<td>{'<br/>'.join(fragments)}</td>")
        rows.append("</tr>")
    rows.append("</table>")
    return "\n".join(rows)


def render_document(source_dir: Path, output_md: Path, assets_dir: Path, assets_rel: str, converter_script: Path) -> dict[str, int]:
    document_xml = source_dir / "word" / "document.xml"
    rels_xml = source_dir / "word" / "_rels" / "document.xml.rels"

    if not document_xml.exists():
        raise FileNotFoundError(f"Missing {document_xml}")
    if not rels_xml.exists():
        raise FileNotFoundError(f"Missing {rels_xml}")

    assets_dir.mkdir(parents=True, exist_ok=True)

    ctx = Context(
        source_dir=source_dir,
        assets_dir=assets_dir,
        assets_rel=assets_rel,
        converter_script=converter_script,
        rels=load_relationships(rels_xml),
    )

    tree = ET.parse(document_xml)
    root = tree.getroot()
    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("Document body not found")

    blocks: list[str] = []
    for child in list(body):
        tag = local_name(child.tag)
        if tag == "p":
            para = render_paragraph(child, ctx)
            if para:
                blocks.append(para)
        elif tag == "tbl":
            blocks.append(render_table(child, ctx))

    if ctx.manifest:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as handle:
            json.dump(ctx.manifest, handle, ensure_ascii=False, indent=2)
            manifest_path = Path(handle.name)

        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(converter_script),
                    "-ManifestPath",
                    str(manifest_path),
                ],
                check=True,
            )
        finally:
            manifest_path.unlink(missing_ok=True)

    note = [
        f"# {output_md.stem}",
        "",
        "> Auto-extracted from DOCX in document order.",
        "> MathType OLE formulas are preserved as rendered images unless the document has already been converted to raw TeX text in a working copy.",
        f"> Counts: formulas={ctx.formula_count}, images={ctx.image_count}, tables={ctx.table_count}.",
        "",
    ]
    output_md.write_text("\n".join(note + blocks) + "\n", encoding="utf-8")

    return {
        "formula_count": ctx.formula_count,
        "image_count": ctx.image_count,
        "table_count": ctx.table_count,
        "asset_count": len(ctx.asset_lookup),
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract DOCX text while preserving MathType formulas as image assets.")
    parser.add_argument("--source-dir", required=True, help="Unpacked DOCX directory")
    parser.add_argument("--output-md", required=True, help="Output markdown path")
    parser.add_argument("--assets-dir", required=True, help="Output asset directory")
    parser.add_argument("--assets-rel", required=True, help="Asset path relative to the markdown file")
    parser.add_argument("--converter-script", required=True, help="PowerShell asset conversion script path")
    args = parser.parse_args(list(argv) if argv is not None else None)

    stats = render_document(
        source_dir=Path(args.source_dir),
        output_md=Path(args.output_md),
        assets_dir=Path(args.assets_dir),
        assets_rel=args.assets_rel,
        converter_script=Path(args.converter_script),
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
