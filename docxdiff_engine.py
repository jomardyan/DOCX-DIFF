"""Structured DOCX extraction and comparison engine."""

import difflib
import re
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


class ComparisonCancelled(RuntimeError):
    """Raised when a structured comparison is cancelled."""


@dataclass
class DocumentBlock:
    """A logical document block with stable structural metadata."""

    block_id: str
    kind: str
    text: str
    location: Dict[str, object]
    style: Optional[str] = None
    list_level: Optional[int] = None
    normalized_text: str = ""

    def legacy_line(self) -> str:
        """Return the historical line representation used by the CLI."""
        if self.kind == "table_cell":
            table = self.location.get("table", 1)
            row = self.location.get("row", 1)
            cell = self.location.get("cell", 1)
            return f"T{table}R{row}C{cell}|{self.text}"
        return f"P|{self.text}"

    def to_dict(self) -> dict:
        return {
            "id": self.block_id,
            "kind": self.kind,
            "text": self.text,
            "normalized_text": self.normalized_text,
            "location": self.location,
            "style": self.style,
            "list_level": self.list_level,
        }


@dataclass
class WordSpan:
    """A changed text span using character offsets in the original text."""

    change: str
    start: int
    end: int
    text: str

    def to_dict(self) -> dict:
        return {
            "change": self.change,
            "start": self.start,
            "end": self.end,
            "text": self.text,
        }


@dataclass
class BlockChange:
    """A comparison result for one aligned block pair."""

    change_id: str
    change_type: str
    old_block: Optional[DocumentBlock] = None
    new_block: Optional[DocumentBlock] = None
    old_spans: List[WordSpan] = field(default_factory=list)
    new_spans: List[WordSpan] = field(default_factory=list)
    moved_from: Optional[Dict[str, object]] = None
    moved_to: Optional[Dict[str, object]] = None

    def to_dict(self) -> dict:
        return {
            "id": self.change_id,
            "type": self.change_type,
            "old": self.old_block.to_dict() if self.old_block else None,
            "new": self.new_block.to_dict() if self.new_block else None,
            "old_spans": [span.to_dict() for span in self.old_spans],
            "new_spans": [span.to_dict() for span in self.new_spans],
            "moved_from": self.moved_from,
            "moved_to": self.moved_to,
        }


@dataclass
class ComparisonResult:
    """Complete structured comparison result."""

    blocks_a: List[DocumentBlock]
    blocks_b: List[DocumentBlock]
    changes: List[BlockChange]
    statistics: Dict[str, object]
    options: Dict[str, bool]

    @property
    def differences_found(self) -> bool:
        return any(change.change_type != "unchanged" for change in self.changes)

    def legacy_lines_a(self) -> List[str]:
        return [block.legacy_line() for block in self.blocks_a]

    def legacy_lines_b(self) -> List[str]:
        return [block.legacy_line() for block in self.blocks_b]

    def to_dict(self) -> dict:
        return {
            "schema_version": 2,
            "options": self.options,
            "statistics": self.statistics,
            "changes": [change.to_dict() for change in self.changes],
        }


def normalize_text(text: str, ignore_case: bool = False, ignore_whitespace: bool = False) -> str:
    """Normalize text for matching while retaining original block text separately."""
    normalized = " ".join(text.split()) if ignore_whitespace else text
    return normalized.casefold() if ignore_case else normalized


def _check_cancel(cancel_event: Optional[threading.Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ComparisonCancelled("Comparison cancelled")


def _paragraph_kind(paragraph: Paragraph, container_kind: str) -> Tuple[str, Optional[int]]:
    if container_kind == "table_cell":
        return "table_cell", _list_level(paragraph)

    style_name = paragraph.style.name if paragraph.style is not None else None
    if style_name and style_name.lower().startswith("heading"):
        return "heading", None

    list_level = _list_level(paragraph)
    if list_level is None and style_name and style_name.lower().startswith("list"):
        match = re.search(r"(\d+)$", style_name)
        list_level = max(0, int(match.group(1)) - 1) if match else 0
    if list_level is not None:
        return "list_item", list_level
    return container_kind, None


def _list_level(paragraph: Paragraph) -> Optional[int]:
    properties = paragraph._p.pPr
    if properties is None or properties.numPr is None:
        return None
    level = properties.numPr.ilvl
    return int(level.val) if level is not None else 0


def _iter_container_children(parent) -> Iterable[object]:
    if isinstance(parent, DocumentObject):
        element = parent.element.body
    elif isinstance(parent, _Cell):
        element = parent._tc
    else:
        element = parent._element

    for child in element.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _extract_container(
    parent,
    scope: str,
    blocks: List[DocumentBlock],
    counters: Dict[str, int],
    cancel_event: Optional[threading.Event],
    table_path: Optional[List[int]] = None,
    cell_location: Optional[Dict[str, object]] = None,
) -> None:
    paragraph_index = 0
    table_index = 0
    for child in _iter_container_children(parent):
        _check_cancel(cancel_event)
        if isinstance(child, Paragraph):
            paragraph_index += 1
            text = child.text or ""
            if not text.strip():
                continue
            counters["block"] += 1
            container_kind = "table_cell" if cell_location else {
                "body": "paragraph",
                "header": "header",
                "footer": "footer",
            }.get(scope, scope)
            kind, list_level = _paragraph_kind(child, container_kind)
            style_name = child.style.name if child.style is not None else None
            location = {
                "scope": scope,
                "paragraph": paragraph_index,
            }
            if table_path:
                location["table_path"] = list(table_path)
            if cell_location:
                location.update(cell_location)
            blocks.append(
                DocumentBlock(
                    block_id=f"block-{counters['block']}",
                    kind=kind,
                    text=text,
                    location=location,
                    style=style_name,
                    list_level=list_level,
                )
            )
        elif isinstance(child, Table):
            table_index += 1
            counters["table"] += 1
            current_table = counters["table"]
            current_path = list(table_path or []) + [table_index]
            seen_cells = set()
            for row_index, row in enumerate(child.rows, start=1):
                for cell_index, cell in enumerate(row.cells, start=1):
                    cell_key = id(cell._tc)
                    if cell_key in seen_cells:
                        continue
                    seen_cells.add(cell_key)
                    location = {
                        "table": current_table,
                        "row": row_index,
                        "cell": cell_index,
                    }
                    _extract_container(
                        cell,
                        scope,
                        blocks,
                        counters,
                        cancel_event,
                        table_path=current_path,
                        cell_location=location,
                    )


def extract_document_blocks(
    source,
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    cancel_event: Optional[threading.Event] = None,
) -> List[DocumentBlock]:
    """Extract body, table, header, and footer blocks in deterministic reading order."""
    document = Document(str(source)) if isinstance(source, (str, Path)) else source
    blocks: List[DocumentBlock] = []
    counters = {"block": 0, "table": 0}
    _extract_container(document, "body", blocks, counters, cancel_event)

    seen_parts = set()
    for section_index, section in enumerate(document.sections, start=1):
        for scope, container in (("header", section.header), ("footer", section.footer)):
            _check_cancel(cancel_event)
            part_key = str(container.part.partname)
            if part_key in seen_parts:
                continue
            seen_parts.add(part_key)
            start = len(blocks)
            _extract_container(container, scope, blocks, counters, cancel_event)
            for block in blocks[start:]:
                block.location["section"] = section_index

    for block in blocks:
        block.normalized_text = normalize_text(
            block.text,
            ignore_case=ignore_case,
            ignore_whitespace=ignore_whitespace,
        )
    return blocks


def _tokenize(text: str, ignore_whitespace: bool) -> List[Tuple[str, int, int]]:
    pattern = r"\w+|[^\w\s]" if ignore_whitespace else r"\s+|\w+|[^\w\s]"
    return [(match.group(0), match.start(), match.end()) for match in re.finditer(pattern, text)]


def compare_word_spans(
    old_text: str,
    new_text: str,
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
) -> Tuple[List[WordSpan], List[WordSpan]]:
    """Return changed character spans for an aligned block pair."""
    old_tokens = _tokenize(old_text, ignore_whitespace)
    new_tokens = _tokenize(new_text, ignore_whitespace)

    def key(token: Tuple[str, int, int]) -> str:
        value = token[0]
        return value.casefold() if ignore_case else value

    matcher = difflib.SequenceMatcher(
        None,
        [key(token) for token in old_tokens],
        [key(token) for token in new_tokens],
        autojunk=False,
    )
    old_spans: List[WordSpan] = []
    new_spans: List[WordSpan] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        old_change = "deleted" if tag == "delete" else "replaced"
        new_change = "inserted" if tag == "insert" else "replaced"
        if i1 < i2:
            start, end = old_tokens[i1][1], old_tokens[i2 - 1][2]
            old_spans.append(WordSpan(old_change, start, end, old_text[start:end]))
        if j1 < j2:
            start, end = new_tokens[j1][1], new_tokens[j2 - 1][2]
            new_spans.append(WordSpan(new_change, start, end, new_text[start:end]))
    return old_spans, new_spans


def _block_signature(block: DocumentBlock) -> str:
    return f"{block.kind}\0{block.normalized_text}"


def _pair_replacements(
    old_blocks: Sequence[DocumentBlock],
    new_blocks: Sequence[DocumentBlock],
) -> List[Tuple[Optional[DocumentBlock], Optional[DocumentBlock]]]:
    pairs: List[Tuple[Optional[DocumentBlock], Optional[DocumentBlock]]] = []
    count = max(len(old_blocks), len(new_blocks))
    for index in range(count):
        old = old_blocks[index] if index < len(old_blocks) else None
        new = new_blocks[index] if index < len(new_blocks) else None
        pairs.append((old, new))
    return pairs


def compare_blocks(
    blocks_a: List[DocumentBlock],
    blocks_b: List[DocumentBlock],
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    cancel_event: Optional[threading.Event] = None,
) -> ComparisonResult:
    """Align structured blocks and produce block- and word-level changes."""
    for block in blocks_a + blocks_b:
        block.normalized_text = normalize_text(
            block.text,
            ignore_case=ignore_case,
            ignore_whitespace=ignore_whitespace,
        )
    matcher = difflib.SequenceMatcher(
        None,
        [_block_signature(block) for block in blocks_a],
        [_block_signature(block) for block in blocks_b],
        autojunk=False,
    )
    changes: List[BlockChange] = []
    change_number = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        _check_cancel(cancel_event)
        if tag == "equal":
            for old, new in zip(blocks_a[i1:i2], blocks_b[j1:j2]):
                change_number += 1
                changes.append(BlockChange(f"change-{change_number}", "unchanged", old, new))
            continue

        if tag == "replace":
            pairs = _pair_replacements(blocks_a[i1:i2], blocks_b[j1:j2])
        elif tag == "delete":
            pairs = [(block, None) for block in blocks_a[i1:i2]]
        else:
            pairs = [(None, block) for block in blocks_b[j1:j2]]

        for old, new in pairs:
            change_number += 1
            if old is not None and new is not None:
                old_spans, new_spans = compare_word_spans(
                    old.text,
                    new.text,
                    ignore_case=ignore_case,
                    ignore_whitespace=ignore_whitespace,
                )
                change_type = "replaced"
            elif old is not None:
                old_spans = [WordSpan("deleted", 0, len(old.text), old.text)]
                new_spans = []
                change_type = "deleted"
            else:
                old_spans = []
                new_spans = [WordSpan("inserted", 0, len(new.text), new.text)]
                change_type = "added"
            changes.append(
                BlockChange(
                    f"change-{change_number}",
                    change_type,
                    old,
                    new,
                    old_spans,
                    new_spans,
                )
            )

    changes = _mark_moves(changes)
    additions = sum(change.change_type in ("added", "replaced") for change in changes)
    deletions = sum(change.change_type in ("deleted", "replaced") for change in changes)
    unchanged = sum(change.change_type == "unchanged" for change in changes)
    moved = sum(change.change_type == "moved" for change in changes)
    similarity = round(matcher.ratio() * 100, 2)
    statistics = {
        "additions": additions,
        "deletions": deletions,
        "unchanged": unchanged,
        "moved": moved,
        "replacements": sum(change.change_type == "replaced" for change in changes),
        "total_lines_a": len(blocks_a),
        "total_lines_b": len(blocks_b),
        "similarity": similarity,
    }
    return ComparisonResult(
        blocks_a,
        blocks_b,
        changes,
        statistics,
        {"ignore_case": ignore_case, "ignore_whitespace": ignore_whitespace},
    )


def _mark_moves(changes: List[BlockChange]) -> List[BlockChange]:
    deleted: Dict[str, List[BlockChange]] = {}
    for change in changes:
        if change.change_type == "deleted" and change.old_block is not None:
            deleted.setdefault(_block_signature(change.old_block), []).append(change)

    for change in changes:
        if change.change_type != "added" or change.new_block is None:
            continue
        candidates = deleted.get(_block_signature(change.new_block), [])
        if not candidates:
            continue
        source = candidates.pop(0)
        change.change_type = "moved"
        change.old_block = source.old_block
        change.old_spans = []
        change.new_spans = []
        change.moved_from = source.old_block.location
        change.moved_to = change.new_block.location
        source.change_type = "_move_source"

    filtered = [change for change in changes if change.change_type != "_move_source"]
    for index, change in enumerate(filtered, start=1):
        change.change_id = f"change-{index}"
    return filtered


def compare_documents(
    source_a,
    source_b,
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    cancel_event: Optional[threading.Event] = None,
) -> ComparisonResult:
    """Extract and compare two DOCX documents."""
    blocks_a = extract_document_blocks(
        source_a,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        cancel_event=cancel_event,
    )
    blocks_b = extract_document_blocks(
        source_b,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        cancel_event=cancel_event,
    )
    return compare_blocks(
        blocks_a,
        blocks_b,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        cancel_event=cancel_event,
    )
