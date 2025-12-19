from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Tuple


def split_flow_into_lines(flow: List[Dict[str, Any]]) -> List[Tuple[List[Dict[str, Any]], str | None]]:
    lines: List[Tuple[List[Dict[str, Any]], str | None]] = []
    current: List[Dict[str, Any]] = []

    for block in flow:
        if isinstance(block, dict) and block.get("type") == "newline":
            kind = block.get("kind") or "line"
            lines.append((current, kind))
            current = []
            continue
        current.append(block)

    lines.append((current, None))
    return lines


def line_plain_text(blocks: Iterable[Dict[str, Any]]) -> str:
    return "".join(b.get("value", "") for b in blocks if isinstance(b, dict) and b.get("type") == "text")


def strip_prefix_from_blocks(blocks: List[Dict[str, Any]], chars_to_strip: int) -> List[Dict[str, Any]]:
    remaining = chars_to_strip
    out: List[Dict[str, Any]] = []

    for block in blocks:
        if not isinstance(block, dict):
            continue

        if remaining <= 0:
            out.append(block)
            continue

        if block.get("type") != "text":
            # Question prefix should be text-only; if formatting inserts non-text blocks
            # in the prefix area, stop stripping to avoid corrupting structure.
            remaining = 0
            out.append(block)
            continue

        value = block.get("value") or ""
        if remaining >= len(value):
            remaining -= len(value)
            continue

        new_value = value[remaining:]
        remaining = 0
        if new_value:
            out.append({**block, "value": new_value})

    return out


def _block_any_highlight(block: Dict[str, Any]) -> bool:
    if bool(block.get("highlight")):
        return True
    if block.get("type") == "table":
        for row in block.get("rows", []) or []:
            for cell in row or []:
                for inner in cell.get("blocks", []) or []:
                    if isinstance(inner, dict) and _block_any_highlight(inner):
                        return True
    return False


def blocks_any_highlight(blocks: Iterable[Dict[str, Any]], *, marker_highlight: bool = False) -> bool:
    if marker_highlight:
        return True
    return any(isinstance(b, dict) and _block_any_highlight(b) for b in blocks)


def blocks_to_text(blocks: Iterable[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for block in blocks:
        if not isinstance(block, dict):
            continue

        btype = block.get("type")
        if btype == "text":
            parts.append(block.get("value", ""))
        elif btype == "newline":
            parts.append("\n")
        elif btype == "image":
            src = block.get("src") or ""
            kind = block.get("kind") or "image"
            parts.append(f"[{kind}:{src}]")
        elif btype == "math":
            parts.append("[math]")
        elif btype == "table":
            parts.append("[table]")

    text = "".join(parts)
    return " ".join(text.replace("\n", " ").split()).strip()


def clean_option_text(text: str, *, qtype: str) -> str:
    text = text.strip()
    if qtype == "fa":
        # Inline FA options commonly use '.' as a delimiter between options.
        text = re.sub(r"\s*\.$", "", text)
    return text.strip()

