"""Structured report rendering for DOCX Diff."""

import html
from typing import List, Optional

from docxdiff_engine import ComparisonResult, DocumentBlock, WordSpan


def render_html_spans(text: str, spans: List[WordSpan], css_class: str) -> str:
    """Render escaped text with changed character ranges highlighted."""
    if not spans:
        return html.escape(text)
    parts = []
    cursor = 0
    for span in sorted(spans, key=lambda item: item.start):
        parts.append(html.escape(text[cursor : span.start]))
        parts.append(
            f'<mark class="{css_class}">{html.escape(text[span.start : span.end])}</mark>'
        )
        cursor = span.end
    parts.append(html.escape(text[cursor:]))
    return "".join(parts)


def block_location_label(block: Optional[DocumentBlock]) -> str:
    """Return a concise human-readable structural location."""
    if block is None:
        return ""
    location = block.location
    scope = str(location.get("scope", "body")).title()
    if block.kind == "table_cell":
        return (
            f"{scope} / Table {location.get('table')} / "
            f"Row {location.get('row')} / Cell {location.get('cell')}"
        )
    if block.kind == "heading":
        return f"{scope} / {block.style or 'Heading'}"
    if block.kind == "list_item":
        return f"{scope} / List level {(block.list_level or 0) + 1}"
    return f"{scope} / {block.kind.replace('_', ' ').title()}"


def build_structured_html(
    result: ComparisonResult,
    a_name: str,
    b_name: str,
    application_version: str,
    generated_at: str,
    context_lines: int = 3,
    source_metadata: Optional[dict] = None,
    report_id: Optional[str] = None,
) -> str:
    """Build a standalone HTML report from the structured comparison model."""
    changed_indexes = [
        index for index, change in enumerate(result.changes) if change.change_type != "unchanged"
    ]
    visible_indexes = set()
    for index in changed_indexes:
        start = max(0, index - context_lines)
        end = min(len(result.changes), index + context_lines + 1)
        visible_indexes.update(range(start, end))

    rows = []
    previous_index = None
    for index, change in enumerate(result.changes):
        if changed_indexes and index not in visible_indexes:
            continue
        if previous_index is not None and index > previous_index + 1:
            rows.append('<tr class="separator"><td colspan="4">...</td></tr>')
        previous_index = index
        old_text = change.old_block.text if change.old_block else ""
        new_text = change.new_block.text if change.new_block else ""
        old_html = render_html_spans(old_text, change.old_spans, "word-del")
        new_html = render_html_spans(new_text, change.new_spans, "word-add")
        if change.change_type == "moved":
            location = (
                f"{block_location_label(change.old_block)} -> "
                f"{block_location_label(change.new_block)}"
            )
        else:
            location = block_location_label(change.new_block or change.old_block)
        rows.append(
            f'<tr class="change-{change.change_type}">'
            f'<td class="type">{html.escape(change.change_type)}</td>'
            f'<td class="location">{html.escape(location)}</td>'
            f'<td class="old">{old_html}</td>'
            f'<td class="new">{new_html}</td>'
            "</tr>"
        )

    provenance = ""
    if source_metadata:
        provenance = (
            f"<div><strong>Source A SHA-256:</strong> "
            f"{html.escape(source_metadata['file_a']['sha256'])}</div>"
            f"<div><strong>Source B SHA-256:</strong> "
            f"{html.escape(source_metadata['file_b']['sha256'])}</div>"
        )
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DOCX Structured Comparison</title>
<style>
body {{ font-family: "Segoe UI", sans-serif; margin: 24px; color: #24292f; }}
h1 {{ margin-bottom: 4px; }}
.meta {{ color: #57606a; margin-bottom: 18px; }}
.summary {{ display: flex; gap: 16px; margin: 16px 0; }}
.summary span {{ background: #f6f8fa; border: 1px solid #d0d7de; padding: 8px 12px; }}
table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
th, td {{ border: 1px solid #d0d7de; padding: 8px; vertical-align: top; }}
th {{ background: #f6f8fa; text-align: left; }}
.type {{ width: 8%; font-weight: 600; }}
.location {{ width: 17%; color: #57606a; }}
.old, .new {{ width: 37.5%; white-space: pre-wrap; overflow-wrap: anywhere; }}
.change-added .new, .change-moved .new {{ background: #e6ffec; }}
.change-deleted .old, .change-replaced .old {{ background: #ffebe9; }}
.change-replaced .new {{ background: #e6ffec; }}
.word-del {{ background: #ff818266; text-decoration: line-through; }}
.word-add {{ background: #46d16066; }}
.separator td {{ text-align: center; color: #57606a; }}
.footer {{ margin-top: 20px; color: #57606a; font-size: 0.9em; }}
</style>
</head>
<body>
<h1>DOCX Structured Comparison</h1>
<div class="meta">{html.escape(a_name)} compared with {html.escape(b_name)}</div>
<div class="summary">
<span>Similarity: {result.statistics['similarity']}%</span>
<span>Added: {result.statistics['additions']}</span>
<span>Deleted: {result.statistics['deletions']}</span>
<span>Replaced: {result.statistics['replacements']}</span>
<span>Moved: {result.statistics['moved']}</span>
</div>
<table>
<thead><tr><th>Change</th><th>Location</th><th>Original</th><th>Modified</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
<div class="footer">
<div>Generated by DOCX Diff v{html.escape(application_version)} at {html.escape(generated_at)}</div>
<div>Report ID: {html.escape(report_id or '')}</div>
{provenance}
</div>
</body>
</html>"""
