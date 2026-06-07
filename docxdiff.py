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
import json
import logging
import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Tuple, Union

# Version information
__version__ = "1.0.0"
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
    SUPPORTED_EXTENSIONS = (".docx",)
    MAX_FILE_SIZE_MB = 100  # Maximum file size in MB
    ENCODING = "utf-8"


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
        differ = difflib.HtmlDiff(wrapcolumn=80)

        html_content = differ.make_file(
            a_lines, b_lines, fromdesc=a_name, todesc=b_name, context=True, numlines=context_lines
        )

        # Add custom styling and metadata
        custom_style = f"""
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; }}
            .diff {{ border: 1px solid #ddd; border-radius: 4px; }}
            .diff_header {{ background-color: #e0e0e0; padding: 10px; font-weight: bold; }}
            td.diff_header {{ text-align: right; color: #666; }}
            .diff_next {{ background-color: #c0c0c0; }}
            .diff_add {{ background-color: #d4edda; }}
            .diff_chg {{ background-color: #fff3cd; }}
            .diff_sub {{ background-color: #f8d7da; }}
            .footer {{ margin-top: 20px; padding: 10px; background: #f5f5f5;
                      border-radius: 4px; font-size: 0.9em; color: #666; }}
        </style>
        <div class="footer">
            Generated by DOCX Diff v{__version__} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </div>
        """

        html_content = html_content.replace("</head>", f"{custom_style}</head>")

        output_file = Path(output_path)
        output_file.write_text(html_content, encoding=Config.ENCODING)
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

        stats = calculate_diff_stats(a_lines, b_lines)

        output = {
            "metadata": {
                "version": __version__,
                "file_a": a_name,
                "file_b": b_name,
                "timestamp": datetime.now().isoformat(),
                "context_lines": context_lines,
            },
            "statistics": stats,
            "changes": changes,
        }

        output_file = Path(output_path)
        output_file.write_text(
            json.dumps(output, indent=2, ensure_ascii=False), encoding=Config.ENCODING
        )
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
        int: Exit code (0=no differences, 1=differences found, 2=error)
    """
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

    # Load and process files with error handling
    try:
        if args.verbose:
            print(f"Loading {path_a}...", file=sys.stderr)

        a_lines = load_docx_lines(path_a)

        if args.verbose:
            print(f"Loading {path_b}...", file=sys.stderr)

        b_lines = load_docx_lines(path_b)

    except (ValueError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"Unexpected error loading files: {e}", file=sys.stderr)
        logger.exception("Unexpected error in file loading")
        return 2

    # Apply transformations
    if args.ignore_whitespace:
        if args.verbose:
            print("Normalizing whitespace...", file=sys.stderr)
        a_lines = [" ".join(x.split()) for x in a_lines]
        b_lines = [" ".join(x.split()) for x in b_lines]

    if args.ignore_case:
        if args.verbose:
            print("Converting to lowercase...", file=sys.stderr)
        a_lines = [x.lower() for x in a_lines]
        b_lines = [x.lower() for x in b_lines]

    # Calculate statistics if needed
    stats = None
    if args.stats or args.verbose or args.json:
        stats = calculate_diff_stats(a_lines, b_lines)

    # Export to HTML if requested
    if args.html:
        try:
            if args.verbose:
                print(f"Exporting to HTML: {args.html}...", file=sys.stderr)
            export_to_html(a_lines, b_lines, str(path_a), str(path_b), args.html, args.context)
            print(f"HTML export saved to: {args.html}")
        except (IOError, RuntimeError) as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    # Export to JSON if requested
    if args.json:
        try:
            if args.verbose:
                print(f"Exporting to JSON: {args.json}...", file=sys.stderr)
            export_to_json(a_lines, b_lines, str(path_a), str(path_b), args.json, args.context)
            print(f"JSON export saved to: {args.json}")
        except (IOError, RuntimeError) as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    # If only exporting, we might skip normal output
    if args.quiet:
        # Just return exit code based on differences
        matcher = difflib.SequenceMatcher(None, a_lines, b_lines)
        return 0 if matcher.ratio() == 1.0 else 1

    # Redirect output if needed
    original_stdout = sys.stdout
    output_file = None
    if args.output:
        try:
            output_file = open(args.output, "w", encoding=Config.ENCODING)
            sys.stdout = output_file
        except IOError as e:
            print(f"Error: Cannot write to output file: {e}", file=sys.stderr)
            return 2

    try:
        # Display diff based on options
        any_diff = False

        if args.side_by_side:
            print_side_by_side(a_lines, b_lines, str(path_a), str(path_b))
            any_diff = True
        elif args.color:
            any_diff = print_colored_diff(
                a_lines,
                b_lines,
                a_name=str(path_a),
                b_name=str(path_b),
                context_lines=args.context,
            )
        else:
            any_diff = print_unified_diff(
                a_lines,
                b_lines,
                a_name=str(path_a),
                b_name=str(path_b),
                context_lines=args.context,
            )

        if not any_diff:
            print("No differences found.")

        # Show statistics if requested
        if args.stats and stats:
            print_statistics(stats, verbose=args.verbose)

    except Exception as e:
        print(f"Error during diff generation: {e}", file=sys.stderr)
        logger.exception("Unexpected error during diff generation")
        return 2
    finally:
        # Restore stdout
        if output_file:
            output_file.close()
            sys.stdout = original_stdout
            print(f"Output saved to: {args.output}")

    logger.debug(f"Comparison complete. Differences found: {any_diff}")
    return 1 if any_diff else 0


class DocxDiffGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Comparison Tool - Enhanced")
        self.root.geometry("1600x900")  # Large window for better diff viewing

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

        # Statistics tracking
        self.stats = {"additions": 0, "deletions": 0, "changes": 0, "similarity": 0.0}
        self.current_diff_lines = []

        # File selection frame - compact
        file_frame = ttk.LabelFrame(root, text="Files", padding=3)
        file_frame.pack(fill="x", padx=5, pady=(5, 2))

        # File A
        ttk.Label(file_frame, text="A:").grid(row=0, column=0, sticky="w", pady=2, padx=(5, 2))
        self.file_a_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_a_var, width=80).grid(
            row=0, column=1, padx=2, sticky="ew"
        )
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file("a"), width=8).grid(
            row=0, column=2, padx=2
        )

        # File B
        ttk.Label(file_frame, text="B:").grid(row=1, column=0, sticky="w", pady=2, padx=(5, 2))
        self.file_b_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_b_var, width=80).grid(
            row=1, column=1, padx=2, sticky="ew"
        )
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file("b"), width=8).grid(
            row=1, column=2, padx=2
        )

        # Make entry fields expand
        file_frame.columnconfigure(1, weight=1)

        # Compact options and controls frame - single row
        control_frame = ttk.Frame(root)
        control_frame.pack(fill="x", padx=5, pady=(2, 5))

        # Left side - Options in a compact frame
        options_frame = ttk.LabelFrame(control_frame, text="Options", padding=2)
        options_frame.pack(side="left", padx=2)

        ttk.Label(options_frame, text="Context:").grid(row=0, column=0, sticky="w", padx=2)
        self.context_var = tk.IntVar(value=3)
        ttk.Spinbox(options_frame, from_=0, to=20, textvariable=self.context_var, width=6).grid(
            row=0, column=1, padx=2
        )

        self.ignore_case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Ignore Case", variable=self.ignore_case_var).grid(
            row=0, column=2, padx=5
        )

        self.ignore_whitespace_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, text="Ignore Space", variable=self.ignore_whitespace_var
        ).grid(row=0, column=3, padx=5)

        # View options
        ttk.Label(options_frame, text="Font:").grid(row=0, column=4, sticky="w", padx=(10, 2))
        self.font_size_var = tk.IntVar(value=10)
        ttk.Spinbox(
            options_frame,
            from_=8,
            to=20,
            textvariable=self.font_size_var,
            width=6,
            command=self.update_font_size,
        ).grid(row=0, column=5, padx=2)

        self.show_line_numbers_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame,
            text="Line #",
            variable=self.show_line_numbers_var,
            command=self.toggle_line_numbers,
        ).grid(row=0, column=6, padx=5)

        # Statistics inline - compact display
        stats_frame = ttk.LabelFrame(control_frame, text="Stats", padding=2)
        stats_frame.pack(side="left", padx=2, fill="x", expand=True)

        self.stats_label = ttk.Label(stats_frame, text="No comparison yet", font=("Segoe UI", 9))
        self.stats_label.pack(padx=5)

        # Compact action buttons on right side
        button_frame = ttk.LabelFrame(control_frame, text="Actions", padding=2)
        button_frame.pack(side="right", padx=2)

        ttk.Button(
            button_frame, text="Compare", command=self.compare_files, width=8, style="Green.TButton"
        ).pack(side="left", padx=2)
        ttk.Button(button_frame, text="Clear", command=self.clear_results, width=6).pack(
            side="left", padx=2
        )
        ttk.Button(button_frame, text="Export TXT", command=self.export_txt, width=9).pack(
            side="left", padx=2
        )
        ttk.Button(button_frame, text="Export HTML", command=self.export_html, width=10).pack(
            side="left", padx=2
        )
        ttk.Button(button_frame, text="Search", command=self.show_search_dialog, width=7).pack(
            side="left", padx=2
        )

        ttk.Separator(button_frame, orient=tk.VERTICAL).pack(side="left", fill="y", padx=5)

        ttk.Button(button_frame, text="◄", command=self.prev_diff, width=3).pack(
            side="left", padx=1
        )
        ttk.Button(button_frame, text="►", command=self.next_diff, width=3).pack(
            side="left", padx=1
        )

        self.filter_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(button_frame, text="Changes Only", variable=self.filter_var).pack(
            side="left", padx=5
        )

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

        # Side-by-side styling
        self.left_text.configure(font=("Consolas", 10), bg="#ffffff", padx=10, pady=5)
        self.right_text.configure(font=("Consolas", 10), bg="#ffffff", padx=10, pady=5)

        self.left_text.tag_config("removed", background="#ffdce0", foreground="#24292f")
        self.left_text.tag_config("context", foreground="#57606a", background="#ffffff")
        self.right_text.tag_config("added", background="#d1f7d6", foreground="#24292f")
        self.right_text.tag_config("context", foreground="#57606a", background="#ffffff")

        # Status bar
        self.status_bar = ttk.Label(root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Keyboard shortcuts
        self.root.bind("<Control-o>", lambda e: self.browse_file("a"))
        self.root.bind("<Control-Shift-O>", lambda e: self.browse_file("b"))
        self.root.bind("<Control-r>", lambda e: self.compare_files())
        self.root.bind("<Control-f>", lambda e: self.show_search_dialog())
        self.root.bind("<Control-s>", lambda e: self.export_txt())
        self.root.bind("<Control-e>", lambda e: self.export_html())
        self.root.bind("<Control-plus>", lambda e: self.adjust_font(1))
        self.root.bind("<Control-minus>", lambda e: self.adjust_font(-1))
        self.root.bind("<Control-equal>", lambda e: self.adjust_font(1))  # For + without shift
        self.root.bind("<F3>", lambda e: self.next_diff())
        self.root.bind("<Shift-F3>", lambda e: self.prev_diff())

        # Store diff positions for navigation
        self.diff_positions = []
        self.current_diff_index = -1

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
        # This would require more complex implementation with a separate line number column
        # For now, just show a message
        if self.show_line_numbers_var.get():
            self.status_bar.config(text="Line numbers enabled (recompare to see effect)")
        else:
            self.status_bar.config(text="Line numbers disabled")

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
        self.stats = {"additions": 0, "deletions": 0, "changes": 0, "similarity": 0.0}
        self.stats_label.config(text="No comparison yet")
        self.status_bar.config(text="Ready")
        self.diff_positions = []
        self.current_diff_index = -1

    def compare_files(self):
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

        try:
            self.status_bar.config(text="Loading files...")
            self.root.update_idletasks()

            # Load and process files
            a_lines = load_docx_lines(path_a)
            b_lines = load_docx_lines(path_b)

            self.status_bar.config(text="Processing comparison...")
            self.root.update_idletasks()

            if self.ignore_whitespace_var.get():
                a_lines = [" ".join(x.split()) for x in a_lines]
                b_lines = [" ".join(x.split()) for x in b_lines]

            if self.ignore_case_var.get():
                a_lines = [x.lower() for x in a_lines]
                b_lines = [x.lower() for x in b_lines]

            # Calculate similarity ratio
            sequence_matcher = difflib.SequenceMatcher(None, a_lines, b_lines)
            similarity_ratio = sequence_matcher.ratio() * 100

            # Generate diff
            diff = list(
                difflib.unified_diff(
                    a_lines,
                    b_lines,
                    fromfile=str(path_a),
                    tofile=str(path_b),
                    lineterm="",
                    n=self.context_var.get(),
                )
            )

            # Store diff for export
            self.current_diff_lines = diff

            # Calculate statistics
            additions = sum(
                1 for line in diff if line.startswith("+") and not line.startswith("+++")
            )
            deletions = sum(
                1 for line in diff if line.startswith("-") and not line.startswith("---")
            )
            changes = len([op for op in sequence_matcher.get_opcodes() if op[0] == "replace"])

            self.stats = {
                "additions": additions,
                "deletions": deletions,
                "changes": changes,
                "similarity": similarity_ratio,
            }

            # Update statistics display
            stats_text = (
                f"Similarity: {similarity_ratio:.2f}% | "
                f"Additions: {additions} | "
                f"Deletions: {deletions} | "
                f"Changes: {changes}"
            )
            self.stats_label.config(text=stats_text)

            # Display results in unified view with GitHub-style formatting
            self.results_text.delete(1.0, tk.END)

            # Add GitHub-style header
            self.add_diff_header(path_a, path_b, additions, deletions)

            # Track diff positions for navigation
            self.diff_positions = []
            line_num = 1
            any_output = False

            # Apply filter if enabled
            display_lines = []
            if self.filter_var.get():
                # Show only changes (no context lines)
                for line in diff:
                    if (
                        line.startswith("+")
                        or line.startswith("-")
                        or line.startswith("@")
                        or line.startswith("---")
                        or line.startswith("+++")
                    ):
                        display_lines.append(line)
            else:
                display_lines = diff

            for line in display_lines:
                any_output = True
                # File paths (--- and +++)
                if line.startswith("---") or line.startswith("+++"):
                    self.results_text.insert(tk.END, line + "\n", "filepath")
                # Chunk headers (@@ ... @@)
                elif line.startswith("@@"):
                    # Add extra visual separation before chunk
                    self.results_text.insert(tk.END, "\n", "context")
                    pos = f"{line_num}.0"
                    self.results_text.insert(tk.END, line + "\n", "chunk")
                    self.diff_positions.append(pos)  # Track chunk position
                    line_num += 2
                # Added lines - with visual indicator
                elif line.startswith("+") and not line.startswith("+++"):
                    # Add visual indicator
                    if self.show_line_numbers_var.get():
                        display_line = f"  + {line_num:4d} │ " + line[1:] + "\n"
                    else:
                        display_line = "  +" + line[1:] + "\n"
                    pos = f"{line_num}.0"
                    self.results_text.insert(tk.END, display_line, "added")
                    self.diff_positions.append(pos)
                    line_num += 1
                # Removed lines - with visual indicator
                elif line.startswith("-") and not line.startswith("---"):
                    # Add visual indicator
                    if self.show_line_numbers_var.get():
                        display_line = f"  - {line_num:4d} │ " + line[1:] + "\n"
                    else:
                        display_line = "  -" + line[1:] + "\n"
                    pos = f"{line_num}.0"
                    self.results_text.insert(tk.END, display_line, "removed")
                    self.diff_positions.append(pos)
                    line_num += 1
                # Context lines
                else:
                    # Add spacing for context to align with changed lines
                    if self.show_line_numbers_var.get():
                        display_line = (
                            f"    {line_num:4d} │ " + line + "\n" if line else line + "\n"
                        )
                    else:
                        display_line = "   " + line + "\n" if line else line + "\n"
                    self.results_text.insert(tk.END, display_line, "context")
                    line_num += 1

            if not any_output:
                self.results_text.insert(tk.END, "No differences found.\n", "context")
                messagebox.showinfo("Result", "No differences found between the files.")
            else:
                # Add separator at the end
                self.results_text.insert(tk.END, "\n" + "―" * 80 + "\n", "header")
                end_text = f"End of comparison - {len(self.diff_positions)} differences found\n"
                self.results_text.insert(tk.END, end_text, "header")
                self.current_diff_index = -1

            # Display side-by-side view
            self.display_side_by_side(a_lines, b_lines, sequence_matcher)

            self.status_bar.config(text=f"Comparison complete: {len(diff)} difference lines")

        except Exception as e:
            self.status_bar.config(text="Error occurred")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

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
        if not self.current_diff_lines:
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
                    f.write("\n".join(self.current_diff_lines))

                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_bar.config(text=f"Exported to {Path(filename).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export:\n{str(e)}")

    def export_html(self):
        """Export comparison results to an HTML file."""
        if not self.current_diff_lines:
            messagebox.showwarning("Warning", "No comparison results to export.")
            return

        filename = filedialog.asksaveasfilename(
            title="Export to HTML File",
            defaultextension=".html",
            filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")],
        )

        if filename:
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write("<!DOCTYPE html>\n<html>\n<head>\n")
                    f.write("<meta charset='utf-8'>\n")
                    f.write("<title>DOCX Comparison Report</title>\n")
                    f.write("<style>\n")
                    f.write(
                        "body { font-family: 'Courier New', monospace; "
                        "margin: 20px; }\n"
                    )
                    f.write(".header { color: blue; font-weight: bold; }\n")
                    f.write(".added { background-color: #e6ffe6; color: green; }\n")
                    f.write(".removed { background-color: #ffe6e6; color: red; }\n")
                    f.write(
                        ".stats { background-color: #f0f0f0; padding: 10px; "
                        "margin-bottom: 20px; border-radius: 5px; }\n"
                    )
                    f.write("pre { white-space: pre-wrap; }\n")
                    f.write("</style>\n</head>\n<body>\n")
                    f.write("<h1>DOCX Comparison Report</h1>\n")
                    f.write("<div class='stats'>\n")
                    f.write(
                        f"<p><strong>Generated:</strong> "
                        f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>\n"
                    )
                    f.write(
                        f"<p><strong>Similarity:</strong> {self.stats['similarity']:.2f}%</p>\n"
                    )
                    f.write(f"<p><strong>Additions:</strong> {self.stats['additions']} | ")
                    f.write(f"<strong>Deletions:</strong> {self.stats['deletions']} | ")
                    f.write(f"<strong>Changes:</strong> {self.stats['changes']}</p>\n")
                    f.write("</div>\n")
                    f.write("<pre>\n")

                    for line in self.current_diff_lines:
                        line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        if (
                            line.startswith("+++")
                            or line.startswith("---")
                            or line.startswith("@@")
                        ):
                            f.write(f"<span class='header'>{line}</span>\n")
                        elif line.startswith("+"):
                            f.write(f"<span class='added'>{line}</span>\n")
                        elif line.startswith("-"):
                            f.write(f"<span class='removed'>{line}</span>\n")
                        else:
                            f.write(f"{line}\n")

                    f.write("</pre>\n</body>\n</html>")

                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_bar.config(text=f"Exported to {Path(filename).name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export:\n{str(e)}")

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
