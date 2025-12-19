import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_MML = "http://www.w3.org/1998/Math/MathML"

W = f"{{{NS_W}}}"
A = f"{{{NS_A}}}"
R = f"{{{NS_R}}}"
WP = f"{{{NS_WP}}}"
V = f"{{{NS_V}}}"
O = f"{{{NS_O}}}"
M = f"{{{NS_M}}}"
PR = f"{{{NS_PR}}}"
MML = f"{{{NS_MML}}}"

ET.register_namespace("", NS_MML)

EMU_PER_PX = 9525  # 914400 EMU per inch / 96 px per inch
PX_PER_PT = 96 / 72


def _emu_to_px(value: str | None) -> int | None:
    if not value:
        return None
    try:
        return max(1, int(round(int(value) / EMU_PER_PX)))
    except ValueError:
        return None


def _style_pt_to_px(style: str | None) -> tuple[int | None, int | None]:
    if not style:
        return None, None

    width = None
    height = None

    width_match = re.search(r"width:([0-9.]+)pt", style)
    if width_match:
        width = max(1, int(round(float(width_match.group(1)) * PX_PER_PT)))

    height_match = re.search(r"height:([0-9.]+)pt", style)
    if height_match:
        height = max(1, int(round(float(height_match.group(1)) * PX_PER_PT)))

    return width, height


def _run_has_highlight(run: ET.Element) -> bool:
    """
    Detect highlight/shading on a run.

    Word may store highlight as <w:highlight w:val="yellow"/> or as shading
    <w:shd w:fill="FFFF00"/> when users use the paint/highlight UI.
    """
    rpr = run.find(f"./{W}rPr")
    if rpr is None:
        return False

    highlight = rpr.find(f"./{W}highlight")
    if highlight is not None:
        val = (
            highlight.get(f"{W}val")
            or highlight.get("val")
            or highlight.get("w:val")
            or ""
        )
        if val and val.lower() not in {"none", "auto"}:
            return True

    shd = rpr.find(f"./{W}shd")
    if shd is not None:
        fill = shd.get(f"{W}fill") or shd.get("fill") or shd.get("w:fill") or ""
        if fill and fill.lower() not in {"none", "auto", "000000"}:
            return True

    return False


_OMML_OPERATORS = {
    "+",
    "-",
    "−",
    "±",
    "∓",
    "*",
    "×",
    "∙",
    "·",
    "/",
    "÷",
    "=",
    "<",
    ">",
    "≤",
    "≥",
    "(",
    ")",
    "[",
    "]",
    "{",
    "}",
    "|",
    "‖",
    ":",
    ";",
    ",",
    ".",
    "!",
    "?",
}


def _mml(tag: str, text: str | None = None, attrib: dict[str, str] | None = None) -> ET.Element:
    elem = ET.Element(f"{MML}{tag}", attrib or {})
    if text is not None:
        elem.text = text
    return elem


def _wrap_mrow(nodes: list[ET.Element]) -> ET.Element:
    if not nodes:
        return _mml("mrow")
    if len(nodes) == 1:
        return nodes[0]
    row = _mml("mrow")
    for node in nodes:
        row.append(node)
    return row


def _omml_get_val(elem: ET.Element | None) -> str | None:
    if elem is None:
        return None
    return elem.get(f"{M}val") or elem.get("val")


def _omml_text_to_mathml_nodes(text: str) -> list[ET.Element]:
    nodes: list[ET.Element] = []
    if not text:
        return nodes

    buf = ""
    mode: str | None = None  # "alpha" | "num" | "space"

    def flush():
        nonlocal buf, mode
        if not buf:
            return
        if mode == "space":
            nodes.append(_mml("mtext", " "))
        elif mode == "num":
            nodes.append(_mml("mn", buf))
        elif mode == "alpha":
            nodes.append(_mml("mi", buf))
        else:
            nodes.append(_mml("mtext", buf))
        buf = ""
        mode = None

    for ch in text:
        if ch.isspace():
            if mode != "space":
                flush()
            buf = " "
            mode = "space"
            continue

        if ch.isalpha():
            if mode not in ("alpha",):
                flush()
            buf += ch
            mode = "alpha"
            continue

        if ch.isdigit() or (mode == "num" and ch in {".", ","}):
            if mode not in ("num",):
                flush()
            buf += ch
            mode = "num"
            continue

        flush()
        if ch in _OMML_OPERATORS:
            nodes.append(_mml("mo", ch))
        else:
            nodes.append(_mml("mtext", ch))

    flush()
    return nodes


def _omml_child_expr(parent: ET.Element, tag_local: str) -> ET.Element | None:
    container = parent.find(f"./{M}{tag_local}")
    if container is None:
        return None
    expr = container.find(f"./{M}e")
    return expr if expr is not None else container


def _omml_nodes_to_mathml_nodes(elem: ET.Element | None) -> list[ET.Element]:
    if elem is None:
        return []

    tag = elem.tag

    # Containers
    if tag in (f"{M}oMath", f"{M}oMathPara", f"{M}e"):
        nodes: list[ET.Element] = []
        for child in list(elem):
            if child.tag.endswith("Pr") or child.tag.endswith("ctrlPr"):
                continue
            nodes.extend(_omml_nodes_to_mathml_nodes(child))
        return nodes

    # Text runs
    if tag == f"{M}r":
        text_parts: list[str] = []
        for t in elem.findall(f"./{M}t"):
            if t.text:
                text_parts.append(t.text)
        if text_parts:
            return _omml_text_to_mathml_nodes("".join(text_parts))

        nodes: list[ET.Element] = []
        for child in list(elem):
            if child.tag.endswith("Pr") or child.tag.endswith("ctrlPr"):
                continue
            nodes.extend(_omml_nodes_to_mathml_nodes(child))
        return nodes

    if tag == f"{M}t":
        return _omml_text_to_mathml_nodes(elem.text or "")

    # Sub / Sup
    if tag == f"{M}sSub":
        base = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e")))
        sub = _wrap_mrow(_omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sub")))
        msub = _mml("msub")
        msub.append(base)
        msub.append(sub)
        return [msub]

    if tag == f"{M}sSup":
        base = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e")))
        sup = _wrap_mrow(_omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sup")))
        msup = _mml("msup")
        msup.append(base)
        msup.append(sup)
        return [msup]

    if tag == f"{M}sSubSup":
        base = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e")))
        sub = _wrap_mrow(_omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sub")))
        sup = _wrap_mrow(_omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sup")))
        msubsup = _mml("msubsup")
        msubsup.append(base)
        msubsup.append(sub)
        msubsup.append(sup)
        return [msubsup]

    # Fractions
    if tag == f"{M}f":
        num = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}num/{M}e") or elem.find(f"./{M}num")))
        den = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}den/{M}e") or elem.find(f"./{M}den")))
        mfrac = _mml("mfrac")
        mfrac.append(num)
        mfrac.append(den)
        return [mfrac]

    # Roots
    if tag == f"{M}rad":
        deg = elem.find(f"./{M}deg/{M}e") or elem.find(f"./{M}deg")
        radicand = _wrap_mrow(_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e")))
        deg_nodes = _omml_nodes_to_mathml_nodes(deg) if deg is not None else []

        if deg_nodes:
            mroot = _mml("mroot")
            mroot.append(radicand)
            mroot.append(_wrap_mrow(deg_nodes))
            return [mroot]

        msqrt = _mml("msqrt")
        msqrt.append(radicand)
        return [msqrt]

    # Delimiters
    if tag == f"{M}d":
        dpr = elem.find(f"./{M}dPr")
        open_chr = _omml_get_val(dpr.find(f"./{M}begChr") if dpr is not None else None)
        close_chr = _omml_get_val(dpr.find(f"./{M}endChr") if dpr is not None else None)
        attrib: dict[str, str] = {}
        if open_chr:
            attrib["open"] = open_chr
        if close_chr:
            attrib["close"] = close_chr
        mfenced = _mml("mfenced", attrib=attrib)
        for node in _omml_nodes_to_mathml_nodes(elem.find(f"./{M}e")):
            mfenced.append(node)
        return [mfenced]

    # N-ary (sum, integral, ...)
    if tag == f"{M}nary":
        nary_pr = elem.find(f"./{M}naryPr")
        op = _omml_get_val(nary_pr.find(f"./{M}chr") if nary_pr is not None else None) or "∑"
        op_node = _mml("mo", op)

        sub_nodes = _omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sub"))
        sup_nodes = _omml_nodes_to_mathml_nodes(_omml_child_expr(elem, "sup"))

        if sub_nodes and sup_nodes:
            wrapper = _mml("munderover")
            wrapper.append(op_node)
            wrapper.append(_wrap_mrow(sub_nodes))
            wrapper.append(_wrap_mrow(sup_nodes))
            return [wrapper, *_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e"))]

        if sub_nodes:
            wrapper = _mml("munder")
            wrapper.append(op_node)
            wrapper.append(_wrap_mrow(sub_nodes))
            return [wrapper, *_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e"))]

        if sup_nodes:
            wrapper = _mml("mover")
            wrapper.append(op_node)
            wrapper.append(_wrap_mrow(sup_nodes))
            return [wrapper, *_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e"))]

        return [op_node, *_omml_nodes_to_mathml_nodes(elem.find(f"./{M}e"))]

    # Fallback: keep best-effort ordering of children.
    nodes: list[ET.Element] = []
    for child in list(elem):
        if child.tag.endswith("Pr") or child.tag.endswith("ctrlPr"):
            continue
        nodes.extend(_omml_nodes_to_mathml_nodes(child))
    return nodes


def _omml_to_mathml(elem: ET.Element, *, display: str = "inline") -> str | None:
    nodes = _omml_nodes_to_mathml_nodes(elem)
    if not nodes:
        return None
    attrib: dict[str, str] = {}
    if display and display != "inline":
        attrib["display"] = display
    math = _mml("math", attrib=attrib)
    row = _mml("mrow")
    for node in nodes:
        row.append(node)
    math.append(row)
    return ET.tostring(math, encoding="unicode")


def parse_flow(docx_path: str, image_dir: str = "extracted_images"):
    blocks: list[dict] = []

    with zipfile.ZipFile(docx_path) as z:
        document_xml = z.read("word/document.xml")
        rels_xml = z.read("word/_rels/document.xml.rels")

    root = ET.fromstring(document_xml)
    rels_root = ET.fromstring(rels_xml)

    # Map rId -> image filename.
    rel_map: dict[str, str] = {}
    for rel in rels_root.findall(f".//{PR}Relationship"):
        rel_type = rel.get("Type", "")
        if "image" in rel_type:
            rel_map[rel.get("Id")] = Path(rel.get("Target", "")).name

    def parse_run(run: ET.Element) -> list[dict]:
        run_blocks: list[dict] = []
        run_highlight = _run_has_highlight(run)

        for child in list(run):
            # TEXT
            if child.tag == f"{W}t" and child.text:
                run_blocks.append(
                    {
                        "type": "text",
                        "value": child.text,
                        "highlight": run_highlight,
                    }
                )

            # MATH (OMML)
            if child.tag in (f"{M}oMath", f"{M}oMathPara"):
                mathml = _omml_to_mathml(
                    child,
                    display="block" if child.tag == f"{M}oMathPara" else "inline",
                )
                if mathml:
                    run_blocks.append(
                        {
                            "type": "math",
                            "mathml": mathml,
                            "highlight": run_highlight,
                        }
                    )

            # TAB
            if child.tag == f"{W}tab":
                run_blocks.append(
                    {
                        "type": "text",
                        "value": "\t",
                        "highlight": run_highlight,
                    }
                )

            # LINE BREAKS INSIDE PARAGRAPH
            if child.tag in (f"{W}br", f"{W}cr"):
                kind = "line"
                if child.tag == f"{W}br":
                    br_type = child.get(f"{W}type") or child.get("type")
                    if br_type in ("page", "column"):
                        kind = "page"
                run_blocks.append({"type": "newline", "kind": kind})

            # IMAGE (DrawingML)
            if child.tag == f"{W}drawing":
                blip = child.find(f".//{A}blip")
                if blip is None:
                    continue
                rid = blip.get(f"{R}embed")
                filename = rel_map.get(rid)
                if not filename:
                    continue

                extent = child.find(f".//{WP}extent")
                width = _emu_to_px(extent.get("cx") if extent is not None else None)
                height = _emu_to_px(extent.get("cy") if extent is not None else None)

                run_blocks.append(
                    {
                        "type": "image",
                        "kind": "image",
                        "src": f"{image_dir}/{filename}",
                        "width": width,
                        "height": height,
                        "highlight": run_highlight,
                    }
                )

            # IMAGE (VML / OLE equation preview)
            if child.tag in (f"{W}object", f"{W}pict"):
                imagedata = child.find(f".//{V}imagedata")
                if imagedata is None:
                    continue
                rid = imagedata.get(f"{R}id") or imagedata.get(f"{R}href")
                filename = rel_map.get(rid)
                if not filename:
                    continue

                shape = child.find(f".//{V}shape")
                width, height = _style_pt_to_px(shape.get("style") if shape is not None else None)

                kind = "image"
                ole = child.find(f".//{O}OLEObject")
                if ole is not None and "Equation" in (ole.get("ProgID") or ""):
                    kind = "formula"

                run_blocks.append(
                    {
                        "type": "image",
                        "kind": kind,
                        "src": f"{image_dir}/{filename}",
                        "width": width,
                        "height": height,
                        "highlight": run_highlight,
                    }
                )

        return run_blocks

    def parse_paragraph(paragraph: ET.Element) -> list[dict]:
        paragraph_blocks: list[dict] = []

        def iter_inlines(container: ET.Element):
            for child in list(container):
                if child.tag == f"{W}hyperlink":
                    yield from iter_inlines(child)
                    continue

                # Content control wrapper
                if child.tag == f"{W}sdt":
                    sdt_content = child.find(f"./{W}sdtContent")
                    if sdt_content is not None:
                        yield from iter_inlines(sdt_content)
                    continue

                # SmartTag wrapper
                if child.tag == f"{W}smartTag":
                    yield from iter_inlines(child)
                    continue

                yield child

        for inline in iter_inlines(paragraph):
            if inline.tag == f"{W}r":
                paragraph_blocks.extend(parse_run(inline))
                continue

            if inline.tag in (f"{M}oMath", f"{M}oMathPara"):
                mathml = _omml_to_mathml(
                    inline,
                    display="block" if inline.tag == f"{M}oMathPara" else "inline",
                )
                if mathml:
                    paragraph_blocks.append(
                        {
                            "type": "math",
                            "mathml": mathml,
                            "highlight": False,
                        }
                    )
                continue

        paragraph_blocks.append({"type": "newline", "kind": "paragraph"})
        return paragraph_blocks

    def parse_table(table: ET.Element) -> dict:
        rows: list[list[dict]] = []

        for tr in table.findall(f"./{W}tr"):
            row: list[dict] = []
            for tc in tr.findall(f"./{W}tc"):
                colspan = 1
                tc_pr = tc.find(f"./{W}tcPr")
                if tc_pr is not None:
                    grid_span = tc_pr.find(f"./{W}gridSpan")
                    if grid_span is not None:
                        raw = grid_span.get(f"{W}val") or grid_span.get("val")
                        try:
                            colspan = max(1, int(raw)) if raw else 1
                        except ValueError:
                            colspan = 1

                cell_blocks: list[dict] = []
                for child in iter_blocks(tc):
                    if child.tag == f"{W}p":
                        cell_blocks.extend(parse_paragraph(child))
                    elif child.tag == f"{W}tbl":
                        cell_blocks.append(parse_table(child))

                if cell_blocks and cell_blocks[-1].get("type") == "newline":
                    cell_blocks.pop()

                row.append({"blocks": cell_blocks, "colspan": colspan})
            rows.append(row)

        return {"type": "table", "rows": rows}

    def iter_blocks(container: ET.Element):
        for child in list(container):
            if child.tag in (f"{W}p", f"{W}tbl"):
                yield child
                continue

            # Content control wrapper
            if child.tag == f"{W}sdt":
                sdt_content = child.find(f"./{W}sdtContent")
                if sdt_content is not None:
                    yield from iter_blocks(sdt_content)
                continue

            # SmartTag wrapper (rare, but appears in some docs)
            if child.tag == f"{W}smartTag":
                yield from iter_blocks(child)
                continue

    body = root.find(f"{W}body")
    if body is None:
        return blocks

    # Iterate top-level block elements in body order (keeps tables intact).
    for child in iter_blocks(body):
        if child.tag == f"{W}p":
            blocks.extend(parse_paragraph(child))
        elif child.tag == f"{W}tbl":
            blocks.append(parse_table(child))
            blocks.append({"type": "newline", "kind": "paragraph"})

    return blocks
