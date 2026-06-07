#!/usr/bin/env python3
"""DOCX Diff - Compare DOCX files and display differences.

A production-ready tool for comparing Microsoft Word DOCX files with support
for multiple output formats including unified diff, side-by-side, HTML, and JSON.
Supports both CLI and GUI modes.

Author: Hayk Jomardyan
License: MIT
"""

import argparse
import difflib
import hashlib
import html
import io
import json
import logging
import os
import queue
import sys
import tempfile
import tkinter as tk
import threading
import time
import uuid
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Optional, Tuple, Union

from docxdiff_engine import (
    ComparisonCancelled,
    ComparisonResult,
    WordSpan,
    compare_documents,
)
from docxdiff_reports import block_location_label, build_structured_html

# Version information
__version__ = "2.0.0"
__author__ = "Hayk Jomardyan"
__license__ = "MIT"

# Windows DPI awareness for crisp display on high-DPI screens
try:
    from ctypes import windll  # type: ignore[attr-defined]

    # Set DPI awareness for Windows 8.1 and later
    windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except Exception:
    try:
        # Fallback for Windows Vista and later
        windll.user32.SetProcessDPIAware()  # type: ignore
    except Exception:
        pass  # Not on Windows or DPI awareness not available

try:
    from docx import Document
except ImportError:
    print(
        "Error: python-docx is not installed. Install it with: pip install python-docx",
        file=sys.stderr,
    )
    sys.exit(2)


# Configuration constants
class Config:
    """Application configuration constants."""

    DEFAULT_CONTEXT_LINES = 3
    MAX_CONTEXT_LINES = 100
    MIN_CONTEXT_LINES = 0
    DEFAULT_FONT_SIZE = 10
    MIN_FONT_SIZE = 6
    MAX_FONT_SIZE = 24
    GUI_WINDOW_SIZE = "1600x900"
    GUI_MIN_WIDTH = 620
    GUI_MIN_HEIGHT = 520
    GUI_WIDE_BREAKPOINT = 2300
    GUI_COMPACT_BREAKPOINT = 1250
    GUI_NARROW_BREAKPOINT = 800
    SUPPORTED_EXTENSIONS = (".docx",)
    MAX_FILE_SIZE_MB = 100  # Maximum file size in MB
    ENCODING = "utf-8"
    AUDIT_SCHEMA_VERSION = 1


# ANSI color codes for terminal output
class Colors:
    """ANSI color codes for terminal output."""

    RED = "\033[91m"
    GREEN = "\033[92m"
    CYAN = "\033[96m"
    YELLOW = "\033[93m"
    RESET = "\033[0m"
    BOLD = "\033[1m"


# Configure logging
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger("docxdiff")


def utc_now_iso() -> str:
    """Return a timezone-aware UTC timestamp suitable for machine-readable reports."""
    return datetime.now(timezone.utc).isoformat()


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """Calculate a SHA-256 digest without loading the entire file into memory."""
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_source_metadata(path: Path, redact_path: bool = False) -> dict:
    """Build immutable source metadata for reports and audit records."""
    stat = path.stat()
    return {
        "path": path.name if redact_path else str(path),
        "size_bytes": stat.st_size,
        "sha256": sha256_file(path),
    }


def verify_source_metadata(path: Path, metadata: dict) -> None:
    """Fail if an input changed after its provenance metadata was captured."""
    stat = path.stat()
    if stat.st_size != metadata["size_bytes"] or sha256_file(path) != metadata["sha256"]:
        raise RuntimeError(f"Input file changed during comparison: {path}")


def atomic_write_text(path: Path, content: str) -> None:
    """Atomically replace a text file to avoid leaving partial reports."""
    target = path.resolve()
    parent = target.parent
    if not parent.exists():
        raise IOError(f"Output directory does not exist: {parent}")
    if not parent.is_dir():
        raise IOError(f"Output parent is not a directory: {parent}")

    temp_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding=Config.ENCODING,
            dir=str(parent),
            prefix=f".{target.name}.",
            suffix=".tmp",
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            temp_file.write(content)
            temp_file.flush()
            os.fsync(temp_file.fileno())
        os.replace(str(temp_path), str(target))
    except OSError as e:
        if temp_path is not None:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError:
                pass
        raise IOError(f"Failed to write output file {target}: {e}") from e


def append_audit_event(path: Path, event: dict) -> None:
    """Append one durable JSON Lines audit event."""
    target = path.resolve()
    if not target.parent.exists():
        raise IOError(f"Audit log directory does not exist: {target.parent}")

    record = json.dumps(event, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    try:
        with target.open("a", encoding=Config.ENCODING, newline="\n") as audit_file:
            audit_file.write(record + "\n")
            audit_file.flush()
            os.fsync(audit_file.fileno())
    except OSError as e:
        raise IOError(f"Failed to append audit log {target}: {e}") from e


def validate_output_targets(output_paths: List[Path], input_paths: List[Path]) -> Tuple[bool, str]:
    """Reject output collisions that could overwrite inputs or other reports."""
    resolved_inputs = {path.resolve() for path in input_paths}
    seen = set()
    for output_path in output_paths:
        resolved = output_path.resolve()
        if resolved in resolved_inputs:
            return False, f"Output path would overwrite an input file: {resolved}"
        if resolved in seen:
            return False, f"Multiple outputs use the same path: {resolved}"
        seen.add(resolved)
    return True, ""


def evaluate_policy(
    stats: dict,
    min_similarity: Optional[float] = None,
    max_changes: Optional[int] = None,
) -> List[str]:
    """Evaluate CI policy thresholds and return human-readable violations."""
    violations = []
    if min_similarity is not None and stats["similarity"] < min_similarity:
        violations.append(
            f"similarity {stats['similarity']:.2f}% is below required {min_similarity:.2f}%"
        )

    total_changes = stats["additions"] + stats["deletions"]
    if max_changes is not None and total_changes > max_changes:
        violations.append(f"change count {total_changes} exceeds allowed {max_changes}")
    return violations


def build_audit_event(
    run_id: str,
    generated_at: str,
    source_metadata: dict,
    options: dict,
    differences_found: bool,
    stats: dict,
    policy_violations: List[str],
    exit_code: int,
    duration_ms: float,
) -> dict:
    """Build a versioned audit event without document contents."""
    return {
        "schema_version": Config.AUDIT_SCHEMA_VERSION,
        "event_id": run_id,
        "timestamp": generated_at,
        "application": {"name": "docxdiff", "version": __version__},
        "sources": source_metadata,
        "options": options,
        "result": {
            "differences_found": differences_found,
            "statistics": stats,
            "policy_violations": policy_violations,
            "exit_code": exit_code,
            "duration_ms": duration_ms,
        },
    }


def validate_file(path: Path) -> Tuple[bool, str]:
    """Validate file exists, is readable, and has correct extension.

    Args:
        path: Path to the file to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not path.exists():
        return False, f"File not found: {path}"

    if not path.is_file():
        return False, f"Not a file: {path}"

    if path.suffix.lower() not in Config.SUPPORTED_EXTENSIONS:
        return (
            False,
            f"Unsupported file type: {path.suffix}. "
            f"Expected: {', '.join(Config.SUPPORTED_EXTENSIONS)}",
        )

    # Check file size
    file_size_mb = path.stat().st_size / (1024 * 1024)
    if file_size_mb > Config.MAX_FILE_SIZE_MB:
        return False, f"File too large: {file_size_mb:.1f}MB (max: {Config.MAX_FILE_SIZE_MB}MB)"

    # Check if file is readable
    if not os.access(path, os.R_OK):
        return False, f"File not readable: {path}"

    return True, ""


def iter_block_text(doc: Document):
    """Yield logical text blocks in reading order.

    Includes paragraphs and table cells, flattened as text lines.
    Uses a pragmatic approach since python-docx doesn't expose unified block iteration.

    Args:
        doc: Document object from python-docx

    Yields:
        str: Text blocks prefixed with P| for paragraphs or T{table}R{row}C{col}| for cells
    """
    try:
        # Paragraphs first in document order
        for p in doc.paragraphs:
            text = (p.text or "").strip()
            yield f"P|{text}"

        # Tables after paragraphs
        for ti, table in enumerate(doc.tables, start=1):
            for ri, row in enumerate(table.rows, start=1):
                for ci, cell in enumerate(row.cells, start=1):
                    cell_text = " ".join((cell.text or "").split())
                    yield f"T{ti}R{ri}C{ci}|{cell_text}"
    except Exception as e:
        logger.error(f"Error extracting text blocks: {e}")
        raise


def load_docx_lines(path: Path) -> List[str]:
    """Load and extract text lines from a DOCX file.

    Args:
        path: Path to the DOCX file

    Returns:
        List of text lines from the document

    Raises:
        ValueError: If file validation fails
        RuntimeError: If document loading fails
    """
    # Validate file
    is_valid, error_msg = validate_file(path)
    if not is_valid:
        raise ValueError(error_msg)

    try:
        logger.debug(f"Loading document: {path}")
        doc = Document(str(path))
        lines = list(iter_block_text(doc))

        # Filter out empty blocks
        lines = [ln for ln in lines if ln.split("|", 1)[1].strip() != ""]

        logger.debug(f"Extracted {len(lines)} lines from {path.name}")
        return lines

    except Exception as e:
        error_msg = f"Failed to load document {path.name}: {str(e)}"
        logger.error(error_msg)
        raise RuntimeError(error_msg) from e


def print_unified_diff(a_lines, b_lines, a_name, b_name, context_lines=3):
    diff = difflib.unified_diff(
        a_lines,
        b_lines,
        fromfile=a_name,
        tofile=b_name,
        lineterm="",
        n=context_lines,
    )
    any_output = False
    for line in diff:
        any_output = True
        print(line)
    return any_output


def print_colored_diff(
    a_lines: List[str], b_lines: List[str], a_name: str, b_name: str, context_lines: int = 3
) -> bool:
    """Print diff with ANSI color codes for terminal output.

    Args:
        a_lines: Lines from first file
        b_lines: Lines from second file
        a_name: Name/path of first file
        b_name: Name/path of second file
        context_lines: Number of context lines around differences

    Returns:
        True if differences were found, False otherwise
    """
    try:
        diff = difflib.unified_diff(
            a_lines,
            b_lines,
            fromfile=a_name,
            tofile=b_name,
            lineterm="",
            n=context_lines,
        )

        any_output = False
        for line in diff:
            any_output = True
            if line.startswith("---") or line.startswith("+++"):
                print(f"{Colors.BOLD}{Colors.CYAN}{line}{Colors.RESET}")
            elif line.startswith("@@"):
                print(f"{Colors.YELLOW}{line}{Colors.RESET}")
            elif line.startswith("-"):
                print(f"{Colors.RED}{line}{Colors.RESET}")
            elif line.startswith("+"):
                print(f"{Colors.GREEN}{line}{Colors.RESET}")
            else:
                print(line)

        return any_output
    except Exception as e:
        logger.error(f"Error printing colored diff: {e}")
        # Fallback to non-colored output
        return print_unified_diff(a_lines, b_lines, a_name, b_name, context_lines)


def calculate_diff_stats(a_lines: List[str], b_lines: List[str]) -> dict:
    """Calculate statistics about the differences between two files."""
    differ = difflib.Differ()
    diff = list(differ.compare(a_lines, b_lines))

    stats: Dict[str, Union[int, float]] = {
        "additions": sum(1 for line in diff if line.startswith("+ ")),
        "deletions": sum(1 for line in diff if line.startswith("- ")),
        "unchanged": sum(1 for line in diff if line.startswith("  ")),
        "total_lines_a": len(a_lines),
        "total_lines_b": len(b_lines),
    }

    # Calculate similarity ratio
    matcher = difflib.SequenceMatcher(None, a_lines, b_lines)
    stats["similarity"] = round(matcher.ratio() * 100, 2)

    return stats


def export_to_html(
    a_lines: List[str],
    b_lines: List[str],
    a_name: str,
    b_name: str,
    output_path: str,
    context_lines: int = 3,
    source_metadata: Optional[dict] = None,
    generated_at: Optional[str] = None,
    report_id: Optional[str] = None,
    structured_result: Optional[ComparisonResult] = None,
) -> bool:
    """Export diff to HTML file with syntax highlighting.

    Args:
        a_lines: Lines from first file
        b_lines: Lines from second file
        a_name: Name/path of first file
        b_name: Name/path of second file
        output_path: Path where HTML file will be saved
        context_lines: Number of context lines around differences

    Returns:
        True if export was successful

    Raises:
        IOError: If file cannot be written
    """
    try:
        logger.debug(f"Exporting HTML to: {output_path}")
        if structured_result is not None:
            html_content = build_structured_html(
                structured_result,
                a_name,
                b_name,
                application_version=__version__,
                generated_at=generated_at or utc_now_iso(),
                context_lines=context_lines,
                source_metadata=source_metadata,
                report_id=report_id,
            )
            atomic_write_text(Path(output_path), html_content)
            logger.info(f"HTML export successful: {output_path}")
            return True

        differ = difflib.HtmlDiff(wrapcolumn=80)
        html_content = differ.make_file(
            a_lines,
            b_lines,
            fromdesc=html.escape(a_name),
            todesc=html.escape(b_name),
            context=True,
            numlines=context_lines,
        )

        # Add custom styling and metadata
        custom_style = """
        <style>
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; }
            .diff { border: 1px solid #ddd; border-radius: 4px; }
            .diff_header { background-color: #e0e0e0; padding: 10px; font-weight: bold; }
            td.diff_header { text-align: right; color: #666; }
            .diff_next { background-color: #c0c0c0; }
            .diff_add { background-color: #d4edda; }
            .diff_chg { background-color: #fff3cd; }
            .diff_sub { background-color: #f8d7da; }
            .footer { margin-top: 20px; padding: 10px; background: #f5f5f5;
                      border-radius: 4px; font-size: 0.9em; color: #666; }
        </style>
        """

        html_content = html_content.replace("</head>", f"{custom_style}</head>")
        generated_at = generated_at or utc_now_iso()
        footer_parts = [
            f"Generated by DOCX Diff v{__version__}",
            f"Timestamp: {html.escape(generated_at)}",
        ]
        if report_id:
            footer_parts.append(f"Report ID: {html.escape(report_id)}")
        if source_metadata:
            footer_parts.extend(
                [
                    f"Source A SHA-256: {html.escape(source_metadata['file_a']['sha256'])}",
                    f"Source B SHA-256: {html.escape(source_metadata['file_b']['sha256'])}",
                ]
            )
        footer = '<div class="footer">' + "<br>".join(footer_parts) + "</div>"
        html_content = html_content.replace("</body>", f"{footer}</body>")

        output_file = Path(output_path)
        atomic_write_text(output_file, html_content)
        logger.info(f"HTML export successful: {output_path}")
        return True

    except IOError as e:
        error_msg = f"Failed to write HTML file: {str(e)}"
        logger.error(error_msg)
        raise IOError(error_msg) from e
    except Exception as e:
        error_msg = f"Unexpected error during HTML export: {str(e)}"
        logger.error(error_msg)
        raise RuntimeError(error_msg) from e


def export_to_json(
    a_lines: List[str],
    b_lines: List[str],
    a_name: str,
    b_name: str,
    output_path: str,
    context_lines: int = 3,
    source_metadata: Optional[dict] = None,
    generated_at: Optional[str] = None,
    report_id: Optional[str] = None,
    structured_result: Optional[ComparisonResult] = None,
) -> bool:
    """Export diff to JSON format with metadata and statistics.

    Args:
        a_lines: Lines from first file
        b_lines: Lines from second file
        a_name: Name/path of first file
        b_name: Name/path of second file
        output_path: Path where JSON file will be saved
        context_lines: Number of context lines around differences

    Returns:
        True if export was successful

    Raises:
        IOError: If file cannot be written
        json.JSONDecodeError: If JSON encoding fails
    """
    try:
        logger.debug(f"Exporting JSON to: {output_path}")
        differ = difflib.Differ()
        diff = list(differ.compare(a_lines, b_lines))

        changes = []
        for i, line in enumerate(diff):
            change_type = None
            content = line[2:] if len(line) > 2 else ""

            if line.startswith("+ "):
                change_type = "addition"
            elif line.startswith("- "):
                change_type = "deletion"
            elif line.startswith("? "):
                continue  # Skip hint lines
            elif line.startswith("  "):
                change_type = "unchanged"

            if change_type:
                changes.append({"line_number": i + 1, "type": change_type, "content": content})

        stats = (
            structured_result.statistics
            if structured_result is not None
            else calculate_diff_stats(a_lines, b_lines)
        )

        output = {
            "schema_version": 2 if structured_result is not None else 1,
            "metadata": {
                "version": __version__,
                "file_a": a_name,
                "file_b": b_name,
                "timestamp": generated_at or utc_now_iso(),
                "context_lines": context_lines,
                "report_id": report_id,
                "sources": source_metadata,
            },
            "statistics": stats,
            "changes": changes,
        }
        if structured_result is not None:
            output["structured_comparison"] = structured_result.to_dict()
            output["documents"] = {
                "file_a": {
                    "blocks": [block.to_dict() for block in structured_result.blocks_a]
                },
                "file_b": {
                    "blocks": [block.to_dict() for block in structured_result.blocks_b]
                },
            }

        output_file = Path(output_path)
        atomic_write_text(output_file, json.dumps(output, indent=2, ensure_ascii=False))
        logger.info(f"JSON export successful: {output_path}")
        return True

    except (IOError, json.JSONDecodeError) as e:
        error_msg = f"Failed to export JSON: {str(e)}"
        logger.error(error_msg)
        raise IOError(error_msg) from e
    except Exception as e:
        error_msg = f"Unexpected error during JSON export: {str(e)}"
        logger.error(error_msg)
        raise RuntimeError(error_msg) from e


def print_side_by_side(a_lines, b_lines, a_name, b_name, width=80):
    """Print side-by-side comparison."""
    col_width = width // 2 - 3

    print("=" * width)
    print(f"{a_name:<{col_width}} | {b_name:<{col_width}}")
    print("=" * width)

    differ = difflib.Differ()
    diff_result = list(differ.compare(a_lines, b_lines))

    a_idx = b_idx = 0
    for line in diff_result:
        if line.startswith("  "):  # Unchanged
            content = line[2:]
            left = content[:col_width]
            right = content[:col_width]
            print(f"{left:<{col_width}} | {right:<{col_width}}")
            a_idx += 1
            b_idx += 1
        elif line.startswith("- "):  # Deleted
            content = line[2:]
            left = content[:col_width]
            print(f"{left:<{col_width}} | {'':>{col_width}}")
            a_idx += 1
        elif line.startswith("+ "):  # Added
            content = line[2:]
            right = content[:col_width]
            print(f"{'':>{col_width}} | {right:<{col_width}}")
            b_idx += 1
        elif line.startswith("? "):  # Hint line
            continue

    print("=" * width)


def print_statistics(stats: dict, verbose: bool = False):
    """Print statistics about the diff."""
    print("\n" + "=" * 60)
    print("DIFF STATISTICS")
    print("=" * 60)
    print(f"Similarity:        {stats['similarity']}%")
    print(f"Lines Added:       {stats['additions']}")
    print(f"Lines Deleted:     {stats['deletions']}")
    print(f"Lines Unchanged:   {stats['unchanged']}")
    print(f"Total Lines (A):   {stats['total_lines_a']}")
    print(f"Total Lines (B):   {stats['total_lines_b']}")

    if verbose:
        total_changes = stats["additions"] + stats["deletions"]
        if total_changes > 0:
            add_percent = round((stats["additions"] / total_changes) * 100, 2)
            del_percent = round((stats["deletions"] / total_changes) * 100, 2)
            print("\nChange Breakdown:")
            print(f"  Additions:       {add_percent}%")
            print(f"  Deletions:       {del_percent}%")

    print("=" * 60)


def main():
    """Main entry point for CLI mode.

    Returns:
        int: Exit code (0=no differences, 1=differences found, 2=error, 3=policy violation)
    """
    started_at = time.perf_counter()
    run_id = str(uuid.uuid4())
    generated_at = utc_now_iso()

    parser = argparse.ArgumentParser(
        description=f"DOCX Diff v{__version__} - Compare DOCX files and display differences.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s file1.docx file2.docx
  %(prog)s file1.docx file2.docx --color --stats
  %(prog)s file1.docx file2.docx --html output.html
  %(prog)s file1.docx file2.docx --json output.json
  %(prog)s file1.docx file2.docx --side-by-side

For more information, visit: https://github.com/yourusername/docx-diff
        """,
    )

    # Version argument
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    # Positional arguments
    parser.add_argument("file_a", help="First DOCX file path")
    parser.add_argument("file_b", help="Second DOCX file path")

    # Comparison options
    compare_group = parser.add_argument_group("Comparison Options")
    compare_group.add_argument(
        "--context",
        "-c",
        type=int,
        default=Config.DEFAULT_CONTEXT_LINES,
        metavar="N",
        help=(
            f"Number of context lines (default: {Config.DEFAULT_CONTEXT_LINES}, "
            f"max: {Config.MAX_CONTEXT_LINES})"
        ),
    )
    compare_group.add_argument(
        "--ignore-case",
        "-i",
        action="store_true",
        help="Compare case-insensitively",
    )
    compare_group.add_argument(
        "--ignore-whitespace",
        "-w",
        action="store_true",
        help="Normalize whitespace before comparing",
    )

    # Output format options
    output_group = parser.add_argument_group("Output Format Options")
    output_group.add_argument(
        "--color",
        action="store_true",
        help="Display colorized output in terminal",
    )
    output_group.add_argument(
        "--side-by-side",
        "-y",
        action="store_true",
        help="Display differences side by side",
    )
    output_group.add_argument(
        "--stats",
        "-s",
        action="store_true",
        help="Show statistics summary",
    )
    output_group.add_argument(
        "--quiet",
        "-q",
        action="store_true",
        help="Suppress normal output, only show if files differ (exit code)",
    )
    output_group.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Show detailed statistics and information",
    )

    # Export options
    export_group = parser.add_argument_group("Export Options")
    export_group.add_argument(
        "--html",
        metavar="FILE",
        help="Export diff to HTML file",
    )
    export_group.add_argument(
        "--json",
        metavar="FILE",
        help="Export diff to JSON file",
    )
    export_group.add_argument(
        "--output",
        "-o",
        metavar="FILE",
        help="Write output to file instead of stdout",
    )

    governance_group = parser.add_argument_group("Governance and Automation")
    governance_group.add_argument(
        "--min-similarity",
        type=float,
        metavar="PERCENT",
        help="Fail policy checks if similarity is below this percentage (exit code 3)",
    )
    governance_group.add_argument(
        "--max-changes",
        type=int,
        metavar="N",
        help="Fail policy checks if additions plus deletions exceed N (exit code 3)",
    )
    governance_group.add_argument(
        "--audit-log",
        metavar="FILE",
        help="Append a structured JSONL audit record for the comparison",
    )
    governance_group.add_argument(
        "--audit-redact-paths",
        action="store_true",
        help="Store file names instead of absolute paths in the audit log",
    )

    args = parser.parse_args()

    # Enable verbose logging if requested
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    elif args.quiet:
        logger.setLevel(logging.ERROR)

    # Validate context lines
    if not Config.MIN_CONTEXT_LINES <= args.context <= Config.MAX_CONTEXT_LINES:
        print(
            f"Error: Context lines must be between {Config.MIN_CONTEXT_LINES} "
            f"and {Config.MAX_CONTEXT_LINES}",
            file=sys.stderr,
        )
        return 2

    if args.min_similarity is not None and not 0 <= args.min_similarity <= 100:
        print("Error: Minimum similarity must be between 0 and 100", file=sys.stderr)
        return 2

    if args.max_changes is not None and args.max_changes < 0:
        print("Error: Maximum changes cannot be negative", file=sys.stderr)
        return 2

    # Validate and load files
    try:
        path_a = Path(args.file_a).resolve()
        path_b = Path(args.file_b).resolve()
    except Exception as e:
        print(f"Error: Invalid file path: {e}", file=sys.stderr)
        return 2

    # Check if comparing the same file
    if path_a == path_b:
        print("Error: Both files point to the same location", file=sys.stderr)
        return 2

    output_paths = [
        Path(value)
        for value in (args.html, args.json, args.output, args.audit_log)
        if value is not None
    ]
    outputs_valid, output_error = validate_output_targets(output_paths, [path_a, path_b])
    if not outputs_valid:
        print(f"Error: {output_error}", file=sys.stderr)
        return 2

    source_metadata = None
    if args.html or args.json or args.audit_log:
        try:
            source_metadata = {
                "file_a": build_source_metadata(path_a, redact_path=args.audit_redact_paths),
                "file_b": build_source_metadata(path_b, redact_path=args.audit_redact_paths),
            }
        except OSError as e:
            print(f"Error: Failed to fingerprint input files: {e}", file=sys.stderr)
            return 2

    # Load and process files with error handling
    try:
        if args.verbose:
            print(f"Loading {path_a}...", file=sys.stderr)

        if args.verbose:
            print(f"Loading {path_b}...", file=sys.stderr)

        comparison_result = compare_documents(
            path_a,
            path_b,
            ignore_case=args.ignore_case,
            ignore_whitespace=args.ignore_whitespace,
        )
        a_lines = comparison_result.legacy_lines_a()
        b_lines = comparison_result.legacy_lines_b()

        if source_metadata:
            verify_source_metadata(path_a, source_metadata["file_a"])
            verify_source_metadata(path_b, source_metadata["file_b"])

    except (ValueError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"Unexpected error loading files: {e}", file=sys.stderr)
        logger.exception("Unexpected error in file loading")
        return 2

    stats = comparison_result.statistics
    differences_found = comparison_result.differences_found
    policy_violations = evaluate_policy(stats, args.min_similarity, args.max_changes)
    comparison_exit_code = 3 if policy_violations else (1 if differences_found else 0)

    def write_audit_record() -> bool:
        if not args.audit_log:
            return True
        audit_event = build_audit_event(
            run_id=run_id,
            generated_at=generated_at,
            source_metadata=source_metadata,
            options={
                "context_lines": args.context,
                "ignore_case": args.ignore_case,
                "ignore_whitespace": args.ignore_whitespace,
                "min_similarity": args.min_similarity,
                "max_changes": args.max_changes,
            },
            differences_found=differences_found,
            stats=stats,
            policy_violations=policy_violations,
            exit_code=comparison_exit_code,
            duration_ms=round((time.perf_counter() - started_at) * 1000, 2),
        )
        try:
            append_audit_event(Path(args.audit_log), audit_event)
            return True
        except IOError as e:
            print(f"Error: {e}", file=sys.stderr)
            return False

    # Export to HTML if requested
    if args.html:
        try:
            if args.verbose:
                print(f"Exporting to HTML: {args.html}...", file=sys.stderr)
            export_to_html(
                a_lines,
                b_lines,
                str(path_a),
                str(path_b),
                args.html,
                args.context,
                source_metadata=source_metadata,
                generated_at=generated_at,
                report_id=run_id,
                structured_result=comparison_result,
            )
            if not args.quiet:
                print(f"HTML export saved to: {args.html}")
        except (IOError, RuntimeError) as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    # Export to JSON if requested
    if args.json:
        try:
            if args.verbose:
                print(f"Exporting to JSON: {args.json}...", file=sys.stderr)
            export_to_json(
                a_lines,
                b_lines,
                str(path_a),
                str(path_b),
                args.json,
                args.context,
                source_metadata=source_metadata,
                generated_at=generated_at,
                report_id=run_id,
                structured_result=comparison_result,
            )
            if not args.quiet:
                print(f"JSON export saved to: {args.json}")
        except (IOError, RuntimeError) as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    # If only exporting, we might skip normal output
    if args.quiet:
        if not write_audit_record():
            return 2
        return comparison_exit_code

    output_buffer = io.StringIO() if args.output else None

    try:
        def render_output() -> bool:
            if not differences_found:
                print("No differences found.")
                if args.stats:
                    print_statistics(stats, verbose=args.verbose)
                return False

            if args.side_by_side:
                print_side_by_side(a_lines, b_lines, str(path_a), str(path_b))
                rendered_diff = differences_found
            elif args.color:
                rendered_diff = print_colored_diff(
                    a_lines,
                    b_lines,
                    a_name=str(path_a),
                    b_name=str(path_b),
                    context_lines=args.context,
                )
            else:
                rendered_diff = print_unified_diff(
                    a_lines,
                    b_lines,
                    a_name=str(path_a),
                    b_name=str(path_b),
                    context_lines=args.context,
                )

            if args.stats:
                print_statistics(stats, verbose=args.verbose)
            return rendered_diff

        if output_buffer is not None:
            with redirect_stdout(output_buffer):
                any_diff = render_output()
        else:
            any_diff = render_output()

        if output_buffer is not None:
            atomic_write_text(Path(args.output), output_buffer.getvalue())
            print(f"Output saved to: {args.output}")

    except Exception as e:
        print(f"Error during diff generation: {e}", file=sys.stderr)
        logger.exception("Unexpected error during diff generation")
        return 2

    if policy_violations:
        for violation in policy_violations:
            print(f"Policy violation: {violation}", file=sys.stderr)

    if not write_audit_record():
        return 2

    logger.debug(f"Comparison complete. Differences found: {differences_found}")
    return comparison_exit_code


# pylint: disable=too-many-public-methods
class DocxDiffGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Comparison Tool - Enhanced")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        initial_width = min(1400, max(Config.GUI_MIN_WIDTH, screen_width - 80))
        initial_height = min(900, max(Config.GUI_MIN_HEIGHT, screen_height - 120))
        self.root.geometry(f"{initial_width}x{initial_height}")
        self.root.minsize(Config.GUI_MIN_WIDTH, Config.GUI_MIN_HEIGHT)

        # Configure custom style for Compare button
        style = ttk.Style()
        style.configure(
            "Green.TButton", foreground="#0d8a3c", font=("Segoe UI", 9, "bold")  # Green text
        )
        # Map states for better visibility on Windows
        style.map(
            "Green.TButton",
            foreground=[("pressed", "#0a6e2e"), ("active", "#0d8a3c")],
            background=[("pressed", "#d4edda"), ("active", "#e8f5e9")],
        )

        # Load recently used files history
        try:
            self.history_file = Path.home() / ".docxdiff_history.json"
        except Exception:
            self.history_file = Path(".docxdiff_history.json")
        self.history: List[str] = []
        self.load_history()

        # Statistics tracking
        self.stats = {"additions": 0, "deletions": 0, "changes": 0, "similarity": 0.0}
        self.current_diff_lines = []
        self.current_comparison_result: Optional[ComparisonResult] = None
        self._comparison_thread: Optional[threading.Thread] = None
        self._comparison_cancel = threading.Event()
        self._comparison_queue = queue.Queue()

        # File selection frame - compact
        file_frame = ttk.LabelFrame(root, text="Files", padding=3)
        file_frame.pack(fill="x", padx=5, pady=(5, 2))

        # File A
        ttk.Label(file_frame, text="A:").grid(row=0, column=0, sticky="w", pady=2, padx=(5, 2))
        self.file_a_var = tk.StringVar()
        self.file_a_combo = ttk.Combobox(
            file_frame, textvariable=self.file_a_var, values=self.history, width=80
        )
        self.file_a_combo.grid(row=0, column=1, padx=2, sticky="ew")
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file("a"), width=8).grid(
            row=0, column=2, padx=2
        )

        # File B
        ttk.Label(file_frame, text="B:").grid(row=1, column=0, sticky="w", pady=2, padx=(5, 2))
        self.file_b_var = tk.StringVar()
        self.file_b_combo = ttk.Combobox(
            file_frame, textvariable=self.file_b_var, values=self.history, width=80
        )
        self.file_b_combo.grid(row=1, column=1, padx=2, sticky="ew")
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file("b"), width=8).grid(
            row=1, column=2, padx=2
        )

        # Swap Button
        ttk.Button(file_frame, text="⇅", command=self.swap_files, width=3).grid(
            row=0, column=3, rowspan=2, padx=5, pady=2, sticky="ns"
        )

        # Make entry fields expand
        file_frame.columnconfigure(1, weight=1)

        # Responsive options and controls area
        self.control_frame = ttk.Frame(root)
        self.control_frame.pack(fill="x", padx=5, pady=(2, 5))

        self.options_frame = ttk.LabelFrame(self.control_frame, text="Options", padding=3)
        self.stats_frame = ttk.LabelFrame(self.control_frame, text="Stats", padding=3)
        self.actions_frame = ttk.LabelFrame(self.control_frame, text="Actions", padding=3)

        self.context_label = ttk.Label(self.options_frame, text="Context:")
        self.context_var = tk.IntVar(value=3)
        self.context_spinbox = ttk.Spinbox(
            self.options_frame,
            from_=0,
            to=20,
            textvariable=self.context_var,
            width=6,
        )

        self.ignore_case_var = tk.BooleanVar(value=False)
        self.ignore_case_check = ttk.Checkbutton(
            self.options_frame,
            text="Ignore Case",
            variable=self.ignore_case_var,
        )

        self.ignore_whitespace_var = tk.BooleanVar(value=False)
        self.ignore_whitespace_check = ttk.Checkbutton(
            self.options_frame,
            text="Ignore Space",
            variable=self.ignore_whitespace_var,
        )

        self.font_label = ttk.Label(self.options_frame, text="Font:")
        self.font_size_var = tk.IntVar(value=10)
        self.font_spinbox = ttk.Spinbox(
            self.options_frame,
            from_=8,
            to=20,
            textvariable=self.font_size_var,
            width=6,
            command=self.update_font_size,
        )

        self.show_line_numbers_var = tk.BooleanVar(value=False)
        self.line_numbers_check = ttk.Checkbutton(
            self.options_frame,
            text="Line #",
            variable=self.show_line_numbers_var,
            command=self.toggle_line_numbers,
        )

        self.stats_label = ttk.Label(
            self.stats_frame,
            text="No comparison yet",
            font=("Segoe UI", 9),
            anchor=tk.W,
        )
        self.stats_label.pack(fill="x", padx=5)

        self.primary_actions = ttk.Frame(self.actions_frame)
        self.export_actions = ttk.Frame(self.actions_frame)
        self.navigation_actions = ttk.Frame(self.actions_frame)

        self.compare_button = ttk.Button(
            self.primary_actions,
            text="Compare",
            command=self.compare_files,
            width=8,
            style="Green.TButton",
        )
        self.compare_button.pack(side="left", padx=2)
        self.cancel_button = ttk.Button(
            self.primary_actions,
            text="Cancel",
            command=self.cancel_comparison,
            width=7,
            state=tk.DISABLED,
        )
        self.cancel_button.pack(side="left", padx=2)
        self.clear_button = ttk.Button(
            self.primary_actions,
            text="Clear",
            command=self.clear_results,
            width=6,
        )
        self.clear_button.pack(side="left", padx=2)

        self.export_txt_button = ttk.Button(
            self.export_actions,
            text="Export TXT",
            command=self.export_txt,
            width=9,
        )
        self.export_txt_button.pack(side="left", padx=2)
        self.export_html_button = ttk.Button(
            self.export_actions,
            text="Export HTML",
            command=self.export_html,
            width=10,
        )
        self.export_html_button.pack(side="left", padx=2)
        self.export_json_button = ttk.Button(
            self.export_actions,
            text="Export JSON",
            command=self.export_json,
            width=11,
        )
        self.export_json_button.pack(side="left", padx=2)

        self.search_button = ttk.Button(
            self.navigation_actions,
            text="Search",
            command=self.show_search_dialog,
            width=7,
        )
        self.search_button.pack(side="left", padx=2)
        self.prev_button = ttk.Button(
            self.navigation_actions,
            text="◄",
            command=self.prev_diff,
            width=3,
        )
        self.prev_button.pack(side="left", padx=1)
        self.next_button = ttk.Button(
            self.navigation_actions,
            text="►",
            command=self.next_diff,
            width=3,
        )
        self.next_button.pack(side="left", padx=1)

        self.filter_var = tk.BooleanVar(value=False)
        self.changes_only_check = ttk.Checkbutton(
            self.navigation_actions,
            text="Changes Only",
            variable=self.filter_var,
            command=self.refresh_structured_view,
        )
        self.changes_only_check.pack(side="left", padx=5)

        self._responsive_mode = None
        self._layout_after_id = None
        self.apply_responsive_layout(initial_width)

        # Results frame with tabs - maximized for diff viewing
        results_frame = ttk.Frame(root)
        results_frame.pack(fill="both", expand=True, padx=5, pady=(2, 2))

        # Notebook for different views
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill="both", expand=True)

        # Unified diff tab
        unified_frame = ttk.Frame(self.notebook)
        self.notebook.add(unified_frame, text="Unified Diff")

        # Compact toolbar for unified view
        unified_toolbar = ttk.Frame(unified_frame)
        unified_toolbar.pack(fill="x", padx=2, pady=2)

        ttk.Button(unified_toolbar, text="+", command=lambda: self.adjust_font(1), width=3).pack(
            side="left", padx=1
        )
        ttk.Button(unified_toolbar, text="-", command=lambda: self.adjust_font(-1), width=3).pack(
            side="left", padx=1
        )
        ttk.Separator(unified_toolbar, orient=tk.VERTICAL).pack(side="left", fill="y", padx=3)
        ttk.Button(unified_toolbar, text="Copy All", command=self.copy_all_diff, width=8).pack(
            side="left", padx=2
        )
        ttk.Button(
            unified_toolbar, text="Copy Selection", command=self.copy_selection, width=12
        ).pack(side="left", padx=2)

        # Configure the text widget with GitHub-like styling - bigger for more diff space
        self.results_text = scrolledtext.ScrolledText(
            unified_frame,
            wrap=tk.NONE,  # No wrapping for better diff viewing
            font=("Consolas", 10),  # Better monospace font
            bg="#ffffff",  # White background like GitHub
            relief=tk.FLAT,
            borderwidth=0,
            padx=8,
            pady=3,
        )

        # Add horizontal scrollbar for long lines
        h_scrollbar = ttk.Scrollbar(
            unified_frame, orient=tk.HORIZONTAL, command=self.results_text.xview
        )
        self.results_text.configure(xscrollcommand=h_scrollbar.set)

        self.results_text.pack(fill="both", expand=True)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Add context menu
        self.create_context_menu(self.results_text)

        # Side-by-side tab
        sidebyside_frame = ttk.Frame(self.notebook)
        self.notebook.add(sidebyside_frame, text="Side-by-Side")

        # Compact toolbar for side-by-side view
        sbs_toolbar = ttk.Frame(sidebyside_frame)
        sbs_toolbar.pack(fill="x", padx=2, pady=2)

        ttk.Button(sbs_toolbar, text="+", command=lambda: self.adjust_font(1), width=3).pack(
            side="left", padx=1
        )
        ttk.Button(sbs_toolbar, text="-", command=lambda: self.adjust_font(-1), width=3).pack(
            side="left", padx=1
        )
        ttk.Separator(sbs_toolbar, orient=tk.VERTICAL).pack(side="left", fill="y", padx=3)

        self.sync_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sbs_toolbar, text="Sync Scroll", variable=self.sync_scroll_var).pack(
            side="left", padx=5
        )

        ttk.Button(
            sbs_toolbar, text="Copy Left", command=lambda: self.copy_pane("left"), width=9
        ).pack(side="left", padx=2)
        ttk.Button(
            sbs_toolbar, text="Copy Right", command=lambda: self.copy_pane("right"), width=9
        ).pack(side="left", padx=2)

        # Create paned window for side-by-side view
        paned = ttk.PanedWindow(sidebyside_frame, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        # Left pane (File A) - compact header
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        ttk.Label(
            left_frame,
            text="File A (Original)",
            font=("Segoe UI", 9, "bold"),
            background="#ffdce0",
            padding=2,
        ).pack(fill="x")
        self.left_text = scrolledtext.ScrolledText(
            left_frame, wrap=tk.NONE, font=("Consolas", 10), padx=5, pady=2
        )
        self.left_text.pack(fill="both", expand=True)

        # Right pane (File B) - compact header
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        ttk.Label(
            right_frame,
            text="File B (Modified)",
            font=("Segoe UI", 9, "bold"),
            background="#d1f7d6",
            padding=2,
        ).pack(fill="x")
        self.right_text = scrolledtext.ScrolledText(
            right_frame, wrap=tk.NONE, font=("Consolas", 10), padx=5, pady=2
        )
        self.right_text.pack(fill="both", expand=True)

        # Bind scroll events for synchronization
        self.left_text.bind("<MouseWheel>", self.on_scroll)
        self.right_text.bind("<MouseWheel>", self.on_scroll)
        self.left_text.bind("<Button-4>", self.on_scroll)  # Linux scroll up
        self.left_text.bind("<Button-5>", self.on_scroll)  # Linux scroll down
        self.right_text.bind("<Button-4>", self.on_scroll)
        self.right_text.bind("<Button-5>", self.on_scroll)

        # Add context menus
        self.create_context_menu(self.left_text)
        self.create_context_menu(self.right_text)

        # Configure text tags with GitHub-style colors
        # Added lines - GitHub green
        self.results_text.tag_config(
            "added",
            foreground="#24292f",  # Dark text
            background="#d1f7d6",  # Light green background like GitHub
            selectbackground="#8cd99a",
            lmargin1=3,
            lmargin2=3,
        )

        # Removed lines - GitHub red
        self.results_text.tag_config(
            "removed",
            foreground="#24292f",  # Dark text
            background="#ffdce0",  # Light red background like GitHub
            selectbackground="#ffb3ba",
            lmargin1=3,
            lmargin2=3,
        )

        # Context lines - neutral
        self.results_text.tag_config(
            "context",
            foreground="#57606a",  # Gray text
            background="#ffffff",
            lmargin1=3,
            lmargin2=3,
        )

        # Header lines - GitHub gray header style
        self.results_text.tag_config(
            "header",
            foreground="#57606a",  # Gray text
            background="#f6f8fa",  # Light gray background
            font=("Consolas", 10, "bold"),
            lmargin1=3,
            lmargin2=3,
        )

        # File path headers - darker
        self.results_text.tag_config(
            "filepath",
            foreground="#0969da",  # GitHub blue
            background="#f6f8fa",
            font=("Consolas", 10, "bold"),
            lmargin1=3,
            lmargin2=3,
        )

        # Chunk headers (@@)
        self.results_text.tag_config(
            "chunk",
            foreground="#8250df",  # Purple like GitHub
            background="#fbf0ff",  # Light purple background
            font=("Consolas", 10, "bold"),
            lmargin1=3,
            lmargin2=3,
        )

        # Search highlight
        self.results_text.tag_config(
            "search_highlight", background="#fff8c5", foreground="#24292f"  # Soft yellow
        )
        self.results_text.tag_config(
            "word_added", background="#46d160", foreground="#24292f"
        )
        self.results_text.tag_config(
            "word_removed",
            background="#ff8182",
            foreground="#24292f",
            overstrike=True,
        )
        self.results_text.tag_config(
            "moved", background="#ddf4ff", foreground="#0969da"
        )

        # Side-by-side styling
        self.left_text.configure(font=("Consolas", 10), bg="#ffffff", padx=10, pady=5)
        self.right_text.configure(font=("Consolas", 10), bg="#ffffff", padx=10, pady=5)

        self.left_text.tag_config("removed", background="#ffdce0", foreground="#24292f")
        self.left_text.tag_config("context", foreground="#57606a", background="#ffffff")
        self.left_text.tag_config(
            "word_removed",
            background="#ff8182",
            foreground="#24292f",
            overstrike=True,
        )
        self.right_text.tag_config("added", background="#d1f7d6", foreground="#24292f")
        self.right_text.tag_config("context", foreground="#57606a", background="#ffffff")
        self.right_text.tag_config(
            "word_added", background="#46d160", foreground="#24292f"
        )
        self.left_text.tag_config("moved", background="#ddf4ff", foreground="#0969da")
        self.right_text.tag_config("moved", background="#ddf4ff", foreground="#0969da")

        # Status bar
        self.status_bar = ttk.Label(root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Keyboard shortcuts
        self.root.bind("<Control-n>", lambda e: self.new_comparison())
        self.root.bind("<Control-o>", lambda e: self.browse_file("a"))
        self.root.bind("<Control-Shift-O>", lambda e: self.browse_file("b"))
        self.root.bind("<Control-r>", lambda e: self.compare_files())
        self.root.bind("<Control-f>", lambda e: self.show_search_dialog())
        self.root.bind("<Control-s>", lambda e: self.export_txt())
        self.root.bind("<Control-e>", lambda e: self.export_html())
        self.root.bind("<Control-Key-1>", lambda e: self.select_view(0))
        self.root.bind("<Control-Key-2>", lambda e: self.select_view(1))
        self.root.bind("<Control-q>", lambda e: self.root.destroy())
        self.root.bind("<Control-plus>", lambda e: self.adjust_font(1))
        self.root.bind("<Control-minus>", lambda e: self.adjust_font(-1))
        self.root.bind("<Control-equal>", lambda e: self.adjust_font(1))  # For + without shift
        self.root.bind("<F3>", lambda e: self.next_diff())
        self.root.bind("<Shift-F3>", lambda e: self.prev_diff())

        # Store diff positions for navigation
        self.diff_positions = []
        self.current_diff_index = -1

        # Native application menu
        self.create_menu_bar()
        self.root.bind("<Configure>", self.on_root_resize, add="+")

    def on_root_resize(self, event):
        """Debounce root resize events before recalculating the control layout."""
        if event.widget is not self.root:
            return
        if self._layout_after_id is not None:
            self.root.after_cancel(self._layout_after_id)
        self._layout_after_id = self.root.after(
            75,
            lambda width=event.width: self.apply_responsive_layout(width),
        )

    def apply_responsive_layout(self, width=None):
        """Reflow controls into rows appropriate for the available window width."""
        if width is None:
            width = self.root.winfo_width()

        mode = self.responsive_mode_for_width(width)

        self._layout_after_id = None
        if mode == self._responsive_mode:
            return
        self._responsive_mode = mode

        for frame in (self.options_frame, self.stats_frame, self.actions_frame):
            frame.grid_forget()
        for frame in (
            self.primary_actions,
            self.export_actions,
            self.navigation_actions,
        ):
            frame.grid_forget()
        for widget in (
            self.context_label,
            self.context_spinbox,
            self.ignore_case_check,
            self.ignore_whitespace_check,
            self.font_label,
            self.font_spinbox,
            self.line_numbers_check,
        ):
            widget.grid_forget()

        for column in range(3):
            self.control_frame.columnconfigure(column, weight=0)
            self.actions_frame.columnconfigure(column, weight=0)
        for column in range(7):
            self.options_frame.columnconfigure(column, weight=0)

        if mode == "wide":
            self.control_frame.columnconfigure(1, weight=1)
            self.options_frame.grid(row=0, column=0, padx=2, sticky="ew")
            self.stats_frame.grid(row=0, column=1, padx=2, sticky="ew")
            self.actions_frame.grid(row=0, column=2, padx=2, sticky="ew")
        elif mode == "medium":
            self.control_frame.columnconfigure(0, weight=1)
            self.control_frame.columnconfigure(1, weight=1)
            self.options_frame.grid(row=0, column=0, padx=2, sticky="ew")
            self.stats_frame.grid(row=0, column=1, padx=2, sticky="ew")
            self.actions_frame.grid(
                row=1,
                column=0,
                columnspan=2,
                padx=2,
                pady=(3, 0),
                sticky="ew",
            )
        else:
            self.control_frame.columnconfigure(0, weight=1)
            self.options_frame.grid(row=0, column=0, padx=2, sticky="ew")
            self.stats_frame.grid(
                row=1,
                column=0,
                padx=2,
                pady=(3, 0),
                sticky="ew",
            )
            self.actions_frame.grid(
                row=2,
                column=0,
                padx=2,
                pady=(3, 0),
                sticky="ew",
            )

        self._layout_option_controls(mode)
        self._layout_action_groups(mode)
        stats_wrap = max(240, width - 40) if mode in ("compact", "narrow") else 420
        self.stats_label.configure(wraplength=stats_wrap)

    @staticmethod
    def responsive_mode_for_width(width):
        """Map a window width to a stable responsive layout mode."""
        if width >= Config.GUI_WIDE_BREAKPOINT:
            return "wide"
        if width >= Config.GUI_COMPACT_BREAKPOINT:
            return "medium"
        if width >= Config.GUI_NARROW_BREAKPOINT:
            return "compact"
        return "narrow"

    def _layout_option_controls(self, mode):
        """Arrange comparison and view options within the Options group."""
        if mode in ("wide", "medium"):
            placements = (
                (self.context_label, 0, 0, (2, 2)),
                (self.context_spinbox, 0, 1, (2, 6)),
                (self.ignore_case_check, 0, 2, (4, 4)),
                (self.ignore_whitespace_check, 0, 3, (4, 8)),
                (self.font_label, 0, 4, (2, 2)),
                (self.font_spinbox, 0, 5, (2, 6)),
                (self.line_numbers_check, 0, 6, (4, 4)),
            )
        elif mode == "compact":
            placements = (
                (self.context_label, 0, 0, (2, 2)),
                (self.context_spinbox, 0, 1, (2, 6)),
                (self.ignore_case_check, 0, 2, (4, 4)),
                (self.ignore_whitespace_check, 0, 3, (4, 4)),
                (self.font_label, 1, 0, (2, 2)),
                (self.font_spinbox, 1, 1, (2, 6)),
                (self.line_numbers_check, 1, 2, (4, 4)),
            )
        else:
            placements = (
                (self.context_label, 0, 0, (2, 2)),
                (self.context_spinbox, 0, 1, (2, 6)),
                (self.font_label, 0, 2, (8, 2)),
                (self.font_spinbox, 0, 3, (2, 6)),
                (self.ignore_case_check, 1, 0, (2, 4)),
                (self.ignore_whitespace_check, 1, 1, (2, 4)),
                (self.line_numbers_check, 1, 2, (2, 4)),
            )

        for widget, row, column, padx in placements:
            widget.grid(row=row, column=column, padx=padx, pady=2, sticky="w")

    def _layout_action_groups(self, mode):
        """Arrange primary, export, and navigation actions without clipping."""
        if mode in ("wide", "medium"):
            self.actions_frame.columnconfigure(2, weight=1)
            self.primary_actions.grid(row=0, column=0, sticky="w")
            self.export_actions.grid(row=0, column=1, padx=(8, 0), sticky="w")
            self.navigation_actions.grid(row=0, column=2, padx=(8, 0), sticky="e")
        elif mode == "compact":
            self.actions_frame.columnconfigure(1, weight=1)
            self.primary_actions.grid(row=0, column=0, sticky="w")
            self.export_actions.grid(row=0, column=1, padx=(8, 0), sticky="e")
            self.navigation_actions.grid(
                row=1,
                column=0,
                columnspan=2,
                pady=(4, 0),
                sticky="w",
            )
        else:
            self.actions_frame.columnconfigure(0, weight=1)
            self.primary_actions.grid(row=0, column=0, sticky="w")
            self.export_actions.grid(row=1, column=0, pady=(4, 0), sticky="w")
            self.navigation_actions.grid(row=2, column=0, pady=(4, 0), sticky="w")

    def create_menu_bar(self):
        """Create the native top menu and connect it to existing GUI actions."""
        menu_bar = tk.Menu(self.root, tearoff=0)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(
            label="New Comparison",
            command=self.new_comparison,
            accelerator="Ctrl+N",
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="Open File A...",
            command=lambda: self.browse_file("a"),
            accelerator="Ctrl+O",
        )
        file_menu.add_command(
            label="Open File B...",
            command=lambda: self.browse_file("b"),
            accelerator="Ctrl+Shift+O",
        )
        file_menu.add_command(label="Swap Files", command=self.swap_files)
        file_menu.add_separator()

        export_menu = tk.Menu(file_menu, tearoff=0)
        export_menu.add_command(
            label="Text Report...",
            command=self.export_txt,
            accelerator="Ctrl+S",
        )
        export_menu.add_command(
            label="HTML Report...",
            command=self.export_html,
            accelerator="Ctrl+E",
        )
        export_menu.add_command(label="JSON Report...", command=self.export_json)
        file_menu.add_cascade(label="Export", menu=export_menu)
        file_menu.add_separator()
        file_menu.add_command(
            label="Exit",
            command=self.root.destroy,
            accelerator="Ctrl+Q",
        )
        menu_bar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(
            label="Copy Selection",
            command=self.copy_selection,
            accelerator="Ctrl+C",
        )
        edit_menu.add_command(label="Copy All Results", command=self.copy_all_diff)
        edit_menu.add_command(
            label="Select All Results",
            command=lambda: self.select_all(self.results_text),
            accelerator="Ctrl+A",
        )
        edit_menu.add_separator()
        edit_menu.add_command(
            label="Find...",
            command=self.show_search_dialog,
            accelerator="Ctrl+F",
        )
        edit_menu.add_separator()
        edit_menu.add_command(label="Clear Results", command=self.clear_results)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)

        compare_menu = tk.Menu(menu_bar, tearoff=0)
        compare_menu.add_command(
            label="Compare Files",
            command=self.compare_files,
            accelerator="Ctrl+R",
        )
        compare_menu.add_separator()
        compare_menu.add_command(
            label="Previous Difference",
            command=self.prev_diff,
            accelerator="Shift+F3",
        )
        compare_menu.add_command(
            label="Next Difference",
            command=self.next_diff,
            accelerator="F3",
        )
        compare_menu.add_separator()
        compare_menu.add_checkbutton(
            label="Ignore Case",
            variable=self.ignore_case_var,
        )
        compare_menu.add_checkbutton(
            label="Ignore Whitespace",
            variable=self.ignore_whitespace_var,
        )
        compare_menu.add_checkbutton(
            label="Changes Only",
            variable=self.filter_var,
            command=self.refresh_structured_view,
        )
        compare_menu.add_separator()
        compare_menu.add_command(label="Reset Options to Defaults", command=self.reset_options)
        menu_bar.add_cascade(label="Compare", menu=compare_menu)

        view_menu = tk.Menu(menu_bar, tearoff=0)
        view_menu.add_command(
            label="Unified Diff",
            command=lambda: self.select_view(0),
            accelerator="Ctrl+1",
        )
        view_menu.add_command(
            label="Side-by-Side",
            command=lambda: self.select_view(1),
            accelerator="Ctrl+2",
        )
        view_menu.add_separator()
        view_menu.add_command(
            label="Zoom In",
            command=lambda: self.adjust_font(1),
            accelerator="Ctrl++",
        )
        view_menu.add_command(
            label="Zoom Out",
            command=lambda: self.adjust_font(-1),
            accelerator="Ctrl+-",
        )
        view_menu.add_command(label="Reset Zoom", command=self.reset_zoom)
        view_menu.add_separator()
        view_menu.add_checkbutton(
            label="Show Line Numbers",
            variable=self.show_line_numbers_var,
            command=self.toggle_line_numbers,
        )
        view_menu.add_checkbutton(
            label="Synchronize Side-by-Side Scrolling",
            variable=self.sync_scroll_var,
        )
        menu_bar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About DOCX Diff", command=self.show_about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.menu_bar = menu_bar
        self.root.config(menu=menu_bar)

    def new_comparison(self):
        """Reset the workspace for a new comparison."""
        self.cancel_comparison()
        self.file_a_var.set("")
        self.file_b_var.set("")
        self.reset_options()
        self.clear_results()
        self.file_a_combo.focus_set()
        self.status_bar.config(text="Ready for a new comparison")

    def reset_options(self):
        """Restore comparison and view options to application defaults."""
        self.context_var.set(Config.DEFAULT_CONTEXT_LINES)
        self.ignore_case_var.set(False)
        self.ignore_whitespace_var.set(False)
        self.show_line_numbers_var.set(False)
        self.filter_var.set(False)
        self.sync_scroll_var.set(True)
        self.reset_zoom(update_status=False)
        self.status_bar.config(text="Options reset to defaults")

    def reset_zoom(self, update_status=True):
        """Restore the default results font size."""
        self.font_size_var.set(Config.DEFAULT_FONT_SIZE)
        self.update_font_size()
        if update_status:
            self.status_bar.config(text=f"Font size reset to {Config.DEFAULT_FONT_SIZE}pt")

    def select_view(self, index):
        """Select a comparison result tab by index."""
        self.notebook.select(index)

    def show_shortcuts(self):
        """Display the keyboard shortcut reference."""
        shortcuts = (
            "Ctrl+N       New comparison\n"
            "Ctrl+O       Open File A\n"
            "Ctrl+Shift+O Open File B\n"
            "Ctrl+R       Compare files\n"
            "Ctrl+F       Find in results\n"
            "Ctrl+S       Export text report\n"
            "Ctrl+E       Export HTML report\n"
            "Ctrl+1       Unified diff view\n"
            "Ctrl+2       Side-by-side view\n"
            "F3           Next difference\n"
            "Shift+F3     Previous difference\n"
            "Ctrl++/-     Adjust font size\n"
            "Ctrl+Q       Exit"
        )
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)

    def show_about(self):
        """Display application version and capability information."""
        messagebox.showinfo(
            "About DOCX Diff",
            f"DOCX Diff v{__version__}\n\n"
            "Compare Microsoft Word DOCX files using unified and side-by-side views.\n"
            "Includes text, HTML, and JSON report exports.",
        )

    def browse_file(self, file_type):
        filename = filedialog.askopenfilename(
            title=f"Select DOCX File {file_type.upper()}",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if filename:
            if file_type == "a":
                self.file_a_var.set(filename)
            else:
                self.file_b_var.set(filename)
            self.add_to_history(filename)

    def create_context_menu(self, text_widget):
        """Create right-click context menu."""
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(label="Copy", command=lambda: self.copy_text(text_widget))
        context_menu.add_command(label="Select All", command=lambda: self.select_all(text_widget))
        context_menu.add_separator()
        context_menu.add_command(label="Clear", command=lambda: text_widget.delete(1.0, tk.END))

        def show_context_menu(event):
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

        text_widget.bind("<Button-3>", show_context_menu)  # Windows/Linux
        text_widget.bind("<Button-2>", show_context_menu)  # macOS

    def copy_text(self, text_widget):
        """Copy selected text to clipboard."""
        try:
            selected_text = text_widget.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.status_bar.config(text="Copied to clipboard")
        except tk.TclError:
            self.status_bar.config(text="No text selected")

    def select_all(self, text_widget):
        """Select all text in widget."""
        text_widget.tag_add(tk.SEL, "1.0", tk.END)
        text_widget.mark_set(tk.INSERT, "1.0")
        text_widget.see(tk.INSERT)

    def copy_all_diff(self):
        """Copy all diff content."""
        content = self.results_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_bar.config(text="All diff copied to clipboard")

    def copy_selection(self):
        """Copy selected text from results."""
        self.copy_text(self.results_text)

    def copy_pane(self, pane):
        """Copy content from side-by-side pane."""
        text_widget = self.left_text if pane == "left" else self.right_text
        content = text_widget.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_bar.config(text=f"{pane.capitalize()} pane copied to clipboard")

    def adjust_font(self, delta):
        """Adjust font size."""
        current_size = self.font_size_var.get()
        new_size = max(8, min(20, current_size + delta))
        self.font_size_var.set(new_size)
        self.update_font_size()

    def update_font_size(self):
        """Update font size for all text widgets."""
        size = self.font_size_var.get()
        font = ("Consolas", size)

        self.results_text.configure(font=font)
        self.left_text.configure(font=font)
        self.right_text.configure(font=font)

        # Update tag fonts
        for tag in ["added", "removed", "context", "chunk", "filepath"]:
            if tag in ["header", "chunk", "filepath"]:
                self.results_text.tag_config(tag, font=("Consolas", size, "bold"))

        self.status_bar.config(text=f"Font size: {size}pt")

    def toggle_line_numbers(self):
        """Toggle line number display."""
        self.refresh_structured_view()
        if self.show_line_numbers_var.get():
            self.status_bar.config(text="Line numbers enabled")
        else:
            self.status_bar.config(text="Line numbers disabled")

    def refresh_structured_view(self):
        """Re-render the current result after display-only option changes."""
        result = self.current_comparison_result
        if result is None:
            return
        self.render_structured_comparison(
            Path(self.file_a_var.get()),
            Path(self.file_b_var.get()),
            result,
        )

    def on_scroll(self, event):
        """Synchronize scrolling between side-by-side panes."""
        if not self.sync_scroll_var.get():
            return

        # Determine which widget triggered the event
        widget = event.widget
        other_widget = self.right_text if widget == self.left_text else self.left_text

        # Scroll the other widget
        if event.num == 5 or event.delta < 0:  # Scroll down
            other_widget.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:  # Scroll up
            other_widget.yview_scroll(-1, "units")

    def next_diff(self):
        """Jump to next difference."""
        if not self.diff_positions:
            self.status_bar.config(text="No differences to navigate")
            return

        self.current_diff_index = (self.current_diff_index + 1) % len(self.diff_positions)
        pos = self.diff_positions[self.current_diff_index]

        self.results_text.see(pos)
        self.results_text.mark_set(tk.INSERT, pos)
        self.status_bar.config(
            text=f"Difference {self.current_diff_index + 1} of {len(self.diff_positions)}"
        )

    def prev_diff(self):
        """Jump to previous difference."""
        if not self.diff_positions:
            self.status_bar.config(text="No differences to navigate")
            return

        self.current_diff_index = (self.current_diff_index - 1) % len(self.diff_positions)
        pos = self.diff_positions[self.current_diff_index]

        self.results_text.see(pos)
        self.results_text.mark_set(tk.INSERT, pos)
        self.status_bar.config(
            text=f"Difference {self.current_diff_index + 1} of {len(self.diff_positions)}"
        )

    def clear_results(self):
        self.results_text.delete(1.0, tk.END)
        self.left_text.delete(1.0, tk.END)
        self.right_text.delete(1.0, tk.END)
        self.current_diff_lines = []
        self.current_comparison_result = None
        self.stats = {"additions": 0, "deletions": 0, "changes": 0, "similarity": 0.0}
        self.stats_label.config(text="No comparison yet")
        self.status_bar.config(text="Ready")
        self.diff_positions = []
        self.current_diff_index = -1

    def compare_files(self):
        """Validate inputs and start a cancellable background comparison."""
        file_a = self.file_a_var.get()
        file_b = self.file_b_var.get()

        if not file_a or not file_b:
            messagebox.showerror("Error", "Please select both files to compare.")
            return

        path_a = Path(file_a)
        path_b = Path(file_b)

        if not path_a.exists() or not path_b.exists():
            messagebox.showerror("Error", "One or both input files do not exist.")
            return

        if path_a.suffix.lower() != ".docx" or path_b.suffix.lower() != ".docx":
            messagebox.showerror("Error", "This tool supports DOCX files only.")
            return

        if self._comparison_thread is not None and self._comparison_thread.is_alive():
            messagebox.showwarning("Comparison Running", "Cancel the current comparison first.")
            return

        self._comparison_cancel.clear()
        self.compare_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.status_bar.config(text="Loading and comparing document structure...")

        options = {
            "ignore_case": self.ignore_case_var.get(),
            "ignore_whitespace": self.ignore_whitespace_var.get(),
        }
        self._comparison_thread = threading.Thread(
            target=self._comparison_worker,
            args=(path_a, path_b, options),
            daemon=True,
        )
        self._comparison_thread.start()
        self.root.after(50, self._poll_comparison_queue)

    def cancel_comparison(self):
        """Request cancellation of the active comparison."""
        if self._comparison_thread is not None and self._comparison_thread.is_alive():
            self._comparison_cancel.set()
            self.status_bar.config(text="Cancelling comparison...")

    def _comparison_worker(self, path_a: Path, path_b: Path, options: dict):
        try:
            result = compare_documents(
                path_a,
                path_b,
                ignore_case=options["ignore_case"],
                ignore_whitespace=options["ignore_whitespace"],
                cancel_event=self._comparison_cancel,
            )
        except ComparisonCancelled:
            self._comparison_queue.put(("cancelled", None))
        except Exception as error:
            self._comparison_queue.put(("failed", error))
        else:
            self._comparison_queue.put(("complete", (path_a, path_b, result)))

    def _poll_comparison_queue(self):
        """Deliver worker results on Tk's main thread."""
        try:
            message, payload = self._comparison_queue.get_nowait()
        except queue.Empty:
            if self._comparison_thread is not None and self._comparison_thread.is_alive():
                self.root.after(50, self._poll_comparison_queue)
            return

        if message == "complete":
            path_a, path_b, result = payload
            self._comparison_complete(path_a, path_b, result)
        elif message == "failed":
            self._comparison_failed(payload)
        else:
            self._comparison_cancelled()

    def _finish_comparison_state(self):
        self.compare_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self._comparison_thread = None

    def _comparison_cancelled(self):
        self._finish_comparison_state()
        self.status_bar.config(text="Comparison cancelled")

    def _comparison_failed(self, error: Exception):
        self._finish_comparison_state()
        self.status_bar.config(text="Error occurred")
        messagebox.showerror("Error", f"An error occurred:\n{str(error)}")

    def _comparison_complete(
        self,
        path_a: Path,
        path_b: Path,
        result: ComparisonResult,
    ):
        self._finish_comparison_state()
        self.current_comparison_result = result
        self.add_to_history(str(path_a))
        self.add_to_history(str(path_b))

        a_lines = result.legacy_lines_a()
        b_lines = result.legacy_lines_b()
        self.current_diff_lines = list(
            difflib.unified_diff(
                a_lines,
                b_lines,
                fromfile=str(path_a),
                tofile=str(path_b),
                lineterm="",
                n=self.context_var.get(),
            )
        )
        self.stats = {
            "additions": result.statistics["additions"],
            "deletions": result.statistics["deletions"],
            "changes": result.statistics["replacements"],
            "moved": result.statistics["moved"],
            "similarity": result.statistics["similarity"],
        }
        self.stats_label.config(
            text=(
                f"Similarity: {result.statistics['similarity']:.2f}% | "
                f"Additions: {result.statistics['additions']} | "
                f"Deletions: {result.statistics['deletions']} | "
                f"Replaced: {result.statistics['replacements']} | "
                f"Moved: {result.statistics['moved']}"
            )
        )
        self.render_structured_comparison(path_a, path_b, result)
        changed = sum(change.change_type != "unchanged" for change in result.changes)
        self.status_bar.config(text=f"Comparison complete: {changed} structured changes")
        if not result.differences_found:
            messagebox.showinfo("Result", "No differences found between the files.")

    def _insert_text_with_spans(
        self,
        widget,
        prefix: str,
        text: str,
        base_tag: str,
        span_tag: str,
        spans: List[WordSpan],
    ):
        widget.insert(tk.END, prefix, base_tag)
        cursor = 0
        for span in sorted(spans, key=lambda item: item.start):
            widget.insert(tk.END, text[cursor : span.start], base_tag)
            widget.insert(tk.END, text[span.start : span.end], (base_tag, span_tag))
            cursor = span.end
        widget.insert(tk.END, text[cursor:] + "\n", base_tag)

    def render_structured_comparison(
        self,
        path_a: Path,
        path_b: Path,
        result: ComparisonResult,
    ):
        """Render one structured result in both GUI views."""
        self.results_text.delete(1.0, tk.END)
        self.add_diff_header(
            path_a,
            path_b,
            result.statistics["additions"],
            result.statistics["deletions"],
        )
        self.diff_positions = []
        show_only_changes = self.filter_var.get()
        display_number = 1
        for change in result.changes:
            if show_only_changes and change.change_type == "unchanged":
                continue
            if change.change_type == "moved":
                location = (
                    f"{block_location_label(change.old_block)} -> "
                    f"{block_location_label(change.new_block)}"
                )
            else:
                location = block_location_label(change.new_block or change.old_block)
            if change.change_type != "unchanged":
                self.diff_positions.append(self.results_text.index(tk.END))
                self.results_text.insert(
                    tk.END,
                    f"@@ {change.change_type.upper()} | {location} @@\n",
                    "chunk",
                )

            number = f"{display_number:4d} │ " if self.show_line_numbers_var.get() else ""
            if change.change_type == "unchanged":
                block = change.new_block or change.old_block
                self.results_text.insert(tk.END, f"   {number}{block.text}\n", "context")
            elif change.change_type == "deleted":
                self._insert_text_with_spans(
                    self.results_text,
                    f" - {number}",
                    change.old_block.text,
                    "removed",
                    "word_removed",
                    change.old_spans,
                )
            elif change.change_type == "added":
                self._insert_text_with_spans(
                    self.results_text,
                    f" + {number}",
                    change.new_block.text,
                    "added",
                    "word_added",
                    change.new_spans,
                )
            elif change.change_type == "moved":
                self.results_text.insert(
                    tk.END,
                    f" ~ {number}{change.new_block.text}\n",
                    "moved",
                )
            else:
                self._insert_text_with_spans(
                    self.results_text,
                    f" - {number}",
                    change.old_block.text,
                    "removed",
                    "word_removed",
                    change.old_spans,
                )
                self._insert_text_with_spans(
                    self.results_text,
                    f" + {number}",
                    change.new_block.text,
                    "added",
                    "word_added",
                    change.new_spans,
                )
            display_number += 1

        if not result.differences_found:
            self.results_text.insert(tk.END, "No differences found.\n", "context")
        self.current_diff_index = -1
        self.display_structured_side_by_side(result)

    def display_structured_side_by_side(self, result: ComparisonResult):
        """Render aligned structured blocks with word-level highlights."""
        self.left_text.delete(1.0, tk.END)
        self.right_text.delete(1.0, tk.END)
        left_number = 1
        right_number = 1
        show_numbers = self.show_line_numbers_var.get()
        for change in result.changes:
            old_text = change.old_block.text if change.old_block else ""
            new_text = change.new_block.text if change.new_block else ""
            left_prefix = f"{left_number:4d} │ " if show_numbers and old_text else ""
            right_prefix = f"{right_number:4d} │ " if show_numbers and new_text else ""
            if change.change_type == "unchanged":
                self.left_text.insert(tk.END, left_prefix + old_text + "\n", "context")
                self.right_text.insert(tk.END, right_prefix + new_text + "\n", "context")
            elif change.change_type == "moved":
                self.left_text.insert(tk.END, left_prefix + old_text + "\n", "moved")
                self.right_text.insert(tk.END, right_prefix + new_text + "\n", "moved")
            else:
                if old_text:
                    self._insert_text_with_spans(
                        self.left_text,
                        left_prefix,
                        old_text,
                        "removed",
                        "word_removed",
                        change.old_spans,
                    )
                else:
                    self.left_text.insert(tk.END, "\n", "context")
                if new_text:
                    self._insert_text_with_spans(
                        self.right_text,
                        right_prefix,
                        new_text,
                        "added",
                        "word_added",
                        change.new_spans,
                    )
                else:
                    self.right_text.insert(tk.END, "\n", "context")
            if old_text:
                left_number += 1
            if new_text:
                right_number += 1

    def add_diff_header(self, path_a, path_b, additions, deletions):
        """Add a GitHub-style header to the diff view."""
        # Add top separator
        separator = "═" * 100 + "\n"
        self.results_text.insert(tk.END, separator, "header")

        # File names header
        header_text = "  Comparing Files\n"
        self.results_text.insert(tk.END, header_text, "header")
        self.results_text.insert(tk.END, separator, "header")

        # File A info
        file_a_text = f"  📄 {path_a.name}\n"
        file_a_path = f"     {str(path_a)}\n"
        self.results_text.insert(tk.END, file_a_text, "filepath")
        self.results_text.insert(tk.END, file_a_path, "context")

        # VS text
        vs_text = "  ⚡ vs\n"
        self.results_text.insert(tk.END, vs_text, "chunk")

        # File B info
        file_b_text = f"  📄 {path_b.name}\n"
        file_b_path = f"     {str(path_b)}\n"
        self.results_text.insert(tk.END, file_b_text, "filepath")
        self.results_text.insert(tk.END, file_b_path, "context")

        # Statistics summary
        summary = f"\n  📊 Changes: +{additions} additions, -{deletions} deletions\n"
        self.results_text.insert(tk.END, summary, "chunk")

        # Bottom separator
        self.results_text.insert(tk.END, separator + "\n", "header")

    def display_side_by_side(self, a_lines, b_lines, matcher):
        """Display side-by-side comparison with optional line numbers."""
        self.left_text.delete(1.0, tk.END)
        self.right_text.delete(1.0, tk.END)

        opcodes = matcher.get_opcodes()
        show_line_nums = self.show_line_numbers_var.get()

        left_line_num = 1
        right_line_num = 1

        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                for line in a_lines[i1:i2]:
                    if show_line_nums:
                        display_line = f"{left_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.left_text.insert(tk.END, display_line, "context")
                    left_line_num += 1

                for line in b_lines[j1:j2]:
                    if show_line_nums:
                        display_line = f"{right_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.right_text.insert(tk.END, display_line, "context")
                    right_line_num += 1

            elif tag == "delete":
                for line in a_lines[i1:i2]:
                    if show_line_nums:
                        display_line = f"{left_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.left_text.insert(tk.END, display_line, "removed")
                    left_line_num += 1

                # Add empty lines to right pane
                for _ in range(i2 - i1):
                    if show_line_nums:
                        self.right_text.insert(tk.END, "     │ \n", "context")
                    else:
                        self.right_text.insert(tk.END, "\n")

            elif tag == "insert":
                # Add empty lines to left pane
                for _ in range(j2 - j1):
                    if show_line_nums:
                        self.left_text.insert(tk.END, "     │ \n", "context")
                    else:
                        self.left_text.insert(tk.END, "\n")

                for line in b_lines[j1:j2]:
                    if show_line_nums:
                        display_line = f"{right_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.right_text.insert(tk.END, display_line, "added")
                    right_line_num += 1

            elif tag == "replace":
                for line in a_lines[i1:i2]:
                    if show_line_nums:
                        display_line = f"{left_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.left_text.insert(tk.END, display_line, "removed")
                    left_line_num += 1

                for line in b_lines[j1:j2]:
                    if show_line_nums:
                        display_line = f"{right_line_num:4d} │ {line}\n"
                    else:
                        display_line = line + "\n"
                    self.right_text.insert(tk.END, display_line, "added")
                    right_line_num += 1

                # Align columns by padding the shorter one with empty spacer lines
                a_count = i2 - i1
                b_count = j2 - j1
                if a_count > b_count:
                    for _ in range(a_count - b_count):
                        if show_line_nums:
                            self.right_text.insert(tk.END, "     │ \n", "context")
                        else:
                            self.right_text.insert(tk.END, "\n")
                elif b_count > a_count:
                    for _ in range(b_count - a_count):
                        if show_line_nums:
                            self.left_text.insert(tk.END, "     │ \n", "context")
                        else:
                            self.left_text.insert(tk.END, "\n")

    def export_txt(self):
        """Export comparison results to a text file."""
        if self.current_comparison_result is None:
            messagebox.showwarning("Warning", "No comparison results to export.")
            return

        filename = filedialog.asksaveasfilename(
            title="Export to Text File",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        )

        if filename:
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write("DOCX Comparison Report\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Similarity: {self.stats['similarity']:.2f}%\n")
                    f.write(f"Additions: {self.stats['additions']}\n")
                    f.write(f"Deletions: {self.stats['deletions']}\n")
                    f.write(f"Changes: {self.stats['changes']}\n")
                    f.write("=" * 70 + "\n\n")
                    if self.current_diff_lines:
                        f.write("\n".join(self.current_diff_lines))
                    else:
                        f.write("No differences found.\n")

                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_bar.config(text=f"Exported to {Path(filename).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export:\n{str(e)}")

    def export_html(self):
        """Export comparison results to an HTML file."""
        if self.current_comparison_result is None:
            messagebox.showwarning("Warning", "No comparison results to export.")
            return

        filename = filedialog.asksaveasfilename(
            title="Export to HTML File",
            defaultextension=".html",
            filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")],
        )

        if filename:
            try:
                result = self.current_comparison_result
                export_to_html(
                    result.legacy_lines_a(),
                    result.legacy_lines_b(),
                    self.file_a_var.get(),
                    self.file_b_var.get(),
                    filename,
                    self.context_var.get(),
                    structured_result=result,
                )

                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_bar.config(text=f"Exported to {Path(filename).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export:\n{str(e)}")

    def export_json(self):
        """Export comparison results to a JSON file."""
        if self.current_comparison_result is None:
            messagebox.showwarning("Warning", "No comparison results to export.")
            return

        file_a = self.file_a_var.get()
        file_b = self.file_b_var.get()

        if not file_a or not file_b:
            messagebox.showerror("Error", "Missing file information for export.")
            return

        filename = filedialog.asksaveasfilename(
            title="Export to JSON File",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
        )

        if filename:
            try:
                result = self.current_comparison_result
                export_to_json(
                    result.legacy_lines_a(),
                    result.legacy_lines_b(),
                    file_a,
                    file_b,
                    filename,
                    self.context_var.get(),
                    structured_result=result,
                )

                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_bar.config(text=f"Exported to {Path(filename).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export:\n{str(e)}")

    def swap_files(self):
        """Swap File A and File B paths."""
        a = self.file_a_var.get()
        b = self.file_b_var.get()
        self.file_a_var.set(b)
        self.file_b_var.set(a)
        self.status_bar.config(text="Swapped File A and File B")

    def load_history(self):
        """Load recently used files history."""
        self.history = []
        try:
            if self.history_file.exists():
                with open(self.history_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        # Filter out empty strings and non-string values to prevent corruption
                        self.history = [item for item in data if isinstance(item, str) and item]
        except Exception as e:
            logger.warning(f"Failed to load history: {e}")

    def save_history(self):
        """Save recently used files history."""
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"Failed to save history: {e}")

    def add_to_history(self, file_path):
        """Add a file path to history."""
        if not file_path:
            return
        try:
            norm_path = str(Path(file_path).resolve())
        except (OSError, ValueError):
            norm_path = file_path

        if norm_path in self.history:
            self.history.remove(norm_path)
        self.history.insert(0, norm_path)
        self.history = self.history[:10]

        self.save_history()
        self.update_history_comboboxes()

    def update_history_comboboxes(self):
        """Update values of File A and File B comboboxes.

        Uses hasattr checks because load_history/history initialization
        occurs before the combobox widgets are constructed in __init__.
        """
        if hasattr(self, "file_a_combo") and hasattr(self, "file_b_combo"):
            self.file_a_combo["values"] = self.history
            self.file_b_combo["values"] = self.history

    def show_search_dialog(self):
        """Show search dialog."""
        search_window = tk.Toplevel(self.root)
        search_window.title("Search")
        search_window.geometry("400x100")

        ttk.Label(search_window, text="Search for:").pack(pady=10)
        search_entry = ttk.Entry(search_window, width=40)
        search_entry.pack(pady=5)
        search_entry.focus()

        def perform_search():
            query = search_entry.get()
            if not query:
                return

            # Remove previous highlights
            self.results_text.tag_remove("search_highlight", "1.0", tk.END)

            # Search and highlight
            start_pos = "1.0"
            count = 0
            while True:
                start_pos = self.results_text.search(query, start_pos, tk.END, nocase=True)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(query)}c"
                self.results_text.tag_add("search_highlight", start_pos, end_pos)
                start_pos = end_pos
                count += 1

            if count > 0:
                self.status_bar.config(text=f"Found {count} occurrence(s) of '{query}'")
                # Scroll to first match
                self.results_text.see("1.0")
                first_match = self.results_text.search(query, "1.0", tk.END, nocase=True)
                if first_match:
                    self.results_text.see(first_match)
            else:
                self.status_bar.config(text=f"No matches found for '{query}'")
                messagebox.showinfo("Search", f"No matches found for '{query}'")

            search_window.destroy()

        ttk.Button(search_window, text="Search", command=perform_search).pack(pady=10)
        search_entry.bind("<Return>", lambda e: perform_search())


def launch_gui():
    """Launch the GUI application.

    Handles GUI initialization with proper error handling for missing dependencies
    and system configuration issues.
    """
    try:
        root = tk.Tk()
        root.title(f"DOCX Diff v{__version__}")

        # Configure for high DPI displays
        try:
            root.tk.call("tk", "scaling", root.winfo_fpixels("1i") / 72.0)
        except Exception:
            pass

        # Enable visual styles if available
        try:
            root.tk.call("source", "azure.tcl")
            root.tk.call("set_theme", "light")
        except Exception:
            pass  # Azure theme not available, use default

        # Set window icon if available
        try:
            # Try to load icon if it exists
            icon_path = Path(__file__).parent / "icon.ico"
            if icon_path.exists():
                root.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Initialize GUI
        DocxDiffGUI(root)

        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")

        logger.info("GUI launched successfully")
        root.mainloop()

    except ImportError as e:
        error_msg = f"GUI dependencies missing: {e}\nPlease ensure tkinter is installed."
        logger.error(error_msg)
        print(error_msg, file=sys.stderr)
        sys.exit(2)
    except Exception as e:
        error_msg = f"Failed to launch GUI: {e}"
        logger.exception(error_msg)
        print(error_msg, file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    try:
        # If no arguments provided, launch GUI; otherwise use CLI
        if len(sys.argv) == 1:
            logger.info("Starting GUI mode")
            launch_gui()
        else:
            logger.info("Starting CLI mode")
            exit_code = main()
            raise SystemExit(exit_code)
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        print("\nOperation cancelled.", file=sys.stderr)
        sys.exit(130)  # Standard exit code for SIGINT
    except Exception as e:
        logger.exception("Unexpected error in main")
        print(f"\nUnexpected error: {e}", file=sys.stderr)
        sys.exit(2)
