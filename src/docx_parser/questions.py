from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Tuple

from .question_blocks import (
    blocks_any_highlight,
    blocks_to_text,
    clean_option_text,
    line_plain_text,
    split_flow_into_lines,
    strip_prefix_from_blocks,
)


QUESTION_START_RE = re.compile(
    r"^\s*(câu|cau)\s*(\d+)\s*[:.)-]?\s*",
    re.IGNORECASE,
)

PART_TITLE_RE = re.compile(r"^\s*(phần|phan|part)\s+[ivxlcdm\d]+\s*$", re.IGNORECASE)

OPTION_MARKER_RE = re.compile(
    r"(?<!\w)(?:(?P<fa>[A-D])\s*\.\s*|(?P<tof>[a-d])\s*\)\s*)",
    re.IGNORECASE,
)

SHORT_ANSWER_RE = re.compile(r"\{([^{}]+)\}")


def _parse_question_blocks(number: int, blocks: List[Dict[str, Any]]) -> Dict[str, Any]:
    stem_blocks: List[Dict[str, Any]] = []

    options: Dict[str, List[Dict[str, Any]]] = {}
    option_marker_highlight: Dict[str, bool] = {}

    detected_choice_type: str | None = None  # "fa" | "tof"
    current_option: str | None = None

    answer: str | None = None

    def append_to_current(target_blocks: List[Dict[str, Any]]):
        nonlocal current_option
        if current_option is None:
            stem_blocks.extend(target_blocks)
            return
        options.setdefault(current_option, []).extend(target_blocks)

    for block in blocks:
        if not isinstance(block, dict):
            continue

        if block.get("type") != "text":
            append_to_current([block])
            continue

        value = block.get("value") or ""
        highlight = bool(block.get("highlight"))

        idx = 0
        while idx < len(value):
            # Short answer has priority: once detected, we stop interpreting option markers.
            brace_match = None if answer is not None else SHORT_ANSWER_RE.search(value, idx)
            opt_match = None if answer is not None else OPTION_MARKER_RE.search(value, idx)

            next_match = None
            if brace_match and opt_match:
                next_match = brace_match if brace_match.start() <= opt_match.start() else opt_match
            else:
                next_match = brace_match or opt_match

            if next_match is None:
                seg = value[idx:]
                if seg:
                    append_to_current([{**block, "value": seg}])
                break

            seg_before = value[idx : next_match.start()]
            if seg_before:
                append_to_current([{**block, "value": seg_before}])

            if next_match is brace_match:
                answer = (brace_match.group(1) or "").strip()
                idx = brace_match.end()
                continue

            # Option marker
            marker = opt_match
            marker_type = "fa" if marker.group("fa") else "tof"
            label = (marker.group("fa") or marker.group("tof") or "").lower()

            if detected_choice_type is None:
                detected_choice_type = marker_type
            # If mixed markers appear, keep the first detected type and still split by markers
            # so options don't get merged.

            current_option = label
            options.setdefault(label, [])
            option_marker_highlight[label] = option_marker_highlight.get(label, False) or highlight

            idx = marker.end()

    qtype: str
    if answer is not None:
        qtype = "sa"
    elif detected_choice_type is not None:
        qtype = detected_choice_type
    else:
        qtype = "unknown"

    stem_text = blocks_to_text(stem_blocks)

    if qtype == "sa":
        return {
            "id": number,
            "number": number,
            "type": "sa",
            "stem": stem_text,
            "stem_blocks": stem_blocks,
            "answer": answer or "",
        }

    if qtype in {"fa", "tof"}:
        # Always return a/b/c/d buckets to avoid dropping options.
        ordered_labels = ["a", "b", "c", "d"]
        option_items: List[Dict[str, Any]] = []

        for label in ordered_labels:
            blocks_for_option = options.get(label, [])
            text = clean_option_text(blocks_to_text(blocks_for_option), qtype=qtype)
            marker_hl = option_marker_highlight.get(label, False)
            correct = blocks_any_highlight(blocks_for_option, marker_highlight=marker_hl)

            option_items.append(
                {
                    "label": label,
                    "text": text,
                    "blocks": blocks_for_option,
                    "correct": bool(correct),
                }
            )

        return {
            "id": number,
            "number": number,
            "type": qtype,
            "stem": stem_text,
            "stem_blocks": stem_blocks,
            "options": option_items,
        }

    return {
        "id": number,
        "number": number,
        "type": qtype,
        "stem": stem_text,
        "stem_blocks": stem_blocks,
        "options": [],
        "answer": None,
    }


def parse_questions(*, texts: List[Dict[str, Any]] | None, flow: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Parse questions using only the processed JSON (texts + flow).

    - Question boundary: a line starting with "Câu X." (X is a number).
    - FA: A./B./C./D. markers (case-insensitive, allows spaces like "A .").
    - TOF: a)/b)/c)/d) markers (case-insensitive, allows spaces like "a )").
    - SA: contains "{answer}" anywhere in the question content.

    Highlight is taken from `flow` blocks (per-run highlight), and correctness is
    computed within each option scope (not per entire line).
    """
    _ = texts  # kept for signature compatibility / future heuristics

    lines = split_flow_into_lines(flow)
    raw_questions: List[Dict[str, Any]] = []
    current: Dict[str, Any] | None = None

    for line_blocks, break_kind in lines:
        plain = line_plain_text(line_blocks)
        has_non_text = any(
            isinstance(b, dict) and (b.get("type") not in {None, "text"})
            for b in line_blocks
        )

        # Keep image-only paragraphs (e.g. a table screenshot). Previously these were
        # dropped because `plain` only considers text blocks.
        if not plain.strip() and has_non_text:
            if current is not None:
                current["blocks"].extend(line_blocks)
                if break_kind is not None:
                    current["blocks"].append({"type": "newline", "kind": break_kind})
            continue

        if not plain.strip():
            if current is not None and break_kind is not None:
                current["blocks"].append({"type": "newline", "kind": break_kind})
            continue

        if PART_TITLE_RE.match(plain.strip()):
            # Section headers split questions (common in exam docs).
            if current is not None:
                raw_questions.append(current)
                current = None
            continue

        m = QUESTION_START_RE.match(plain)
        if m:
            if current is not None:
                raw_questions.append(current)

            number = int(m.group(2))
            stripped = strip_prefix_from_blocks(line_blocks, m.end())
            current = {"number": number, "blocks": stripped}
        else:
            if current is None:
                continue
            current["blocks"].extend(line_blocks)

        if current is not None and break_kind is not None:
            current["blocks"].append({"type": "newline", "kind": break_kind})

    if current is not None:
        raw_questions.append(current)

    parsed: List[Dict[str, Any]] = []
    for idx, q in enumerate(raw_questions, start=1):
        parsed_question = _parse_question_blocks(idx, q["blocks"])
        parsed_question["source_number"] = q["number"]
        parsed.append(parsed_question)
    return parsed
