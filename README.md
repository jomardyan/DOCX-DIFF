# DOCX DIFF

A Python tool to compare two DOCX files and display text differences with multiple output formats. Perfect for tracking changes in Word documents and identifying modifications between versions.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Mode Comparison](#mode-comparison)
- [Usage](#usage)
  - [GUI Mode (Interactive)](#️-gui-mode-interactive)
  - [CLI Mode (Command Line)](#️-cli-mode-command-line)
- [Visual Guide](#visual-guide)
- [Use Cases](#use-cases)
- [Additional Resources](#-additional-resources)
- [Contributing](#contributing)
- [License](#license)
- [Troubleshooting](#troubleshooting)
- [Future Enhancements](#future-enhancements)

## Features

### 🎯 Core Features
- **Dual Mode Operation**: CLI for automation, GUI for interactive use
- **Text Extraction**: Extracts text from paragraphs and table cells in DOCX files
- **Multiple Output Formats**: Unified diff, side-by-side, colored terminal output
- **Smart Comparison**: Case-insensitive and whitespace normalization options

### 📊 Analysis & Statistics
- **Detailed Statistics**: Similarity percentage, line counts, change breakdown
- **Real-time Updates**: Live statistics in GUI mode
- **Change Navigation**: Jump between differences with keyboard shortcuts

### 💾 Export Options
- **HTML Export**: Beautiful formatted reports with syntax highlighting
- **JSON Export**: Structured data for programmatic analysis
- **Text Export**: Plain text diff output to files
- **Clipboard Support**: Copy entire diff or selections

### 🎨 Visual Features (GUI)
- **Tabbed Interface**: Switch between unified and side-by-side views
- **Color Coding**: Green for additions, red for deletions, white for context
- **Synchronized Scrolling**: Side-by-side panes scroll together
- **Adjustable Font**: Zoom in/out for better readability
- **Line Numbers**: Optional line number display
- **Search Function**: Find text within diff results

### ⚙️ Comparison Options
- **Context Lines**: Configurable (0-20 lines)
- **Case-Insensitive**: Optional case matching
- **Whitespace Normalization**: Ignore spacing differences
- **Filter Changes**: Show only modified lines

### 🖥️ Cross-Platform
- Works on Windows, macOS, and Linux
- Modern UI with native look and feel
- High-DPI aware on Windows

## Requirements

- Python 3.6+
- `python-docx` library

## Installation

### From Source

1. Clone the repository:
```bash
git clone https://github.com/yourusername/docx-diff.git
cd docx-diff
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Quick Start

### For Interactive Use (GUI)
```bash
# Launch GUI - no arguments needed
python docxdiff.py
```

Then use the Browse buttons to select your files and click Compare.

### For Command Line
```bash
# Basic comparison
python docxdiff.py file1.docx file2.docx

# With colored output and statistics
python docxdiff.py file1.docx file2.docx --color --stats

# Export to HTML
python docxdiff.py file1.docx file2.docx --html report.html
```

## Mode Comparison

| Feature | CLI Mode | GUI Mode |
|---------|----------|----------|
| **Launch** | `python docxdiff.py file1 file2` | `python docxdiff.py` |
| **File Selection** | Command arguments | Browse buttons / Drag & drop |
| **Visual Diff** | Colored terminal (optional) | Always color-coded |
| **Side-by-Side** | Text-based columns | Split-pane with sync scroll |
| **Statistics** | Optional flag `--stats` | Always visible |
| **Navigation** | Manual scrolling | ◄/► buttons between changes |
| **Search** | Pipe to grep/findstr | Built-in search (Ctrl+F) |
| **Export HTML** | `--html file.html` | Export HTML button |
| **Export JSON** | `--json file.json` | Not available |
| **Export TXT** | `--output file.txt` | Export TXT button |
| **Copy to Clipboard** | Terminal selection | Copy All / Copy Selection |
| **Font Adjustment** | N/A | +/- buttons, Ctrl+±  |
| **Line Numbers** | N/A | Toggle checkbox |
| **Best For** | Scripting, automation, CI/CD | Interactive review, manual comparison |

## Usage

DOCX DIFF supports two modes of operation: **Command Line Interface (CLI)** for automation and scripting, and **Graphical User Interface (GUI)** for interactive use.

---

## 🖥️ GUI Mode (Interactive)

Launch the GUI by running the script without arguments:

```bash
python docxdiff.py
```

### GUI Features

The GUI provides a user-friendly interface with the following capabilities:

#### **File Selection**
- Browse button for easy file selection
- Drag-and-drop support (on supported platforms)
- Recently used files history

#### **Comparison Options**
- **Context Lines**: Adjustable spinner (0-20 lines)
- **Ignore Case**: Checkbox for case-insensitive comparison
- **Ignore Whitespace**: Checkbox to normalize whitespace
- **Font Size**: Adjustable (8-20pt) for better readability
- **Line Numbers**: Toggle to show/hide line numbers

#### **View Modes**
1. **Unified Diff Tab**
   - GitHub-style color-coded diff display
   - Additions in green background
   - Deletions in red background
   - Context lines in white
   - Horizontal scrolling for long lines

2. **Side-by-Side Tab**
   - Split-pane view with synchronized scrolling
   - Original document (File A) on the left with red highlighting
   - Modified document (File B) on the right with green highlighting
   - Toggle sync scroll on/off
   - Resizable panes

#### **Statistics Panel**
- Real-time similarity percentage
- Line counts (added, deleted, unchanged)
- Total lines in each file
- Change breakdown percentages

#### **Navigation Controls**
- **◄ Previous Diff**: Jump to previous change
- **► Next Diff**: Jump to next change
- **Changes Only**: Filter to show only changed lines
- **Search**: Find text within diff results (Ctrl+F)

#### **Export Options**
- **Export TXT**: Save diff as plain text file
- **Export HTML**: Generate formatted HTML report
- Copy All: Copy entire diff to clipboard
- Copy Selection: Copy selected text

#### **Keyboard Shortcuts**
- `Ctrl+O`: Open File A
- `Ctrl+Shift+O`: Open File B
- `Ctrl+R`: Compare files
- `Ctrl+F`: Search in results
- `F3`: Find next
- `Shift+F3`: Find previous
- `Ctrl+C`: Copy selection
- `Ctrl++`: Increase font size
- `Ctrl+-`: Decrease font size

### GUI Workflow

1. **Launch**: Run `python docxdiff.py` without arguments
2. **Select Files**: Click "Browse" buttons or use Ctrl+O to select files
3. **Configure Options**: Set context lines, case sensitivity, etc.
4. **Compare**: Click "Compare" button (or Ctrl+R)
5. **Review**: Switch between Unified and Side-by-Side views
6. **Navigate**: Use ◄/► buttons to jump between changes
7. **Export**: Save results as TXT or HTML if needed

---

## ⌨️ CLI Mode (Command Line)

Use CLI mode for automation, scripting, and integration with other tools.

### Basic Comparison

Compare two DOCX files:
```bash
python docxdiff.py file1.docx file2.docx
```

### Comparison Options

**Adjust context lines** (default: 3):
```bash
python docxdiff.py file1.docx file2.docx --context 5
python docxdiff.py file1.docx file2.docx -c 5
```

**Case-insensitive comparison**:
```bash
python docxdiff.py file1.docx file2.docx --ignore-case
python docxdiff.py file1.docx file2.docx -i
```

**Normalize whitespace**:
```bash
python docxdiff.py file1.docx file2.docx --ignore-whitespace
python docxdiff.py file1.docx file2.docx -w
```

### Output Format Options

**Colored terminal output**:
```bash
python docxdiff.py file1.docx file2.docx --color
```

**Side-by-side comparison**:
```bash
python docxdiff.py file1.docx file2.docx --side-by-side
python docxdiff.py file1.docx file2.docx -y
```

**Show statistics**:
```bash
python docxdiff.py file1.docx file2.docx --stats
python docxdiff.py file1.docx file2.docx -s
```

**Verbose output**:
```bash
python docxdiff.py file1.docx file2.docx --verbose --stats
python docxdiff.py file1.docx file2.docx -v -s
```

**Quiet mode** (only exit code):
```bash
python docxdiff.py file1.docx file2.docx --quiet
python docxdiff.py file1.docx file2.docx -q
```

### Export Options

**Export to HTML**:
```bash
python docxdiff.py file1.docx file2.docx --html report.html
```

**Export to JSON**:
```bash
python docxdiff.py file1.docx file2.docx --json data.json
```

**Save output to file**:
```bash
python docxdiff.py file1.docx file2.docx --output diff.txt
python docxdiff.py file1.docx file2.docx -o diff.txt
```

### Combined Examples

**Colored diff with statistics**:
```bash
python docxdiff.py file1.docx file2.docx --color --stats
```

**Full report with HTML and JSON export**:
```bash
python docxdiff.py file1.docx file2.docx --html report.html --json data.json --stats
```

**Case-insensitive with whitespace normalization and verbose output**:
```bash
python docxdiff.py file1.docx file2.docx -i -w -v --stats
```

**Quiet mode for scripting**:
```bash
python docxdiff.py file1.docx file2.docx -q
if [ $? -eq 0 ]; then
    echo "Files are identical"
else
    echo "Files differ"
fi
```

### Command Line Options Reference

```
positional arguments:
  file_a                First DOCX file path
  file_b                Second DOCX file path

Comparison Options:
  --context N, -c N     Number of context lines (default: 3)
  --ignore-case, -i     Compare case-insensitively
  --ignore-whitespace, -w
                        Normalize whitespace before comparing

Output Format Options:
  --color               Display colorized output in terminal
  --side-by-side, -y    Display differences side by side
  --stats, -s           Show statistics summary
  --quiet, -q           Suppress output, only show exit code
  --verbose, -v         Show detailed information

Export Options:
  --html FILE           Export diff to HTML file
  --json FILE           Export diff to JSON file
  --output FILE, -o FILE
                        Write output to file instead of stdout
```

### Output Formats

#### Unified Diff (Default)
```
--- file1.docx
+++ file2.docx
@@ -5,3 +5,4 @@
 P|This is a paragraph that didn't change
-P|This line was removed
+P|This line was added instead
 P|Another unchanged line
```

#### Statistics Output
```
============================================================
DIFF STATISTICS
============================================================
Similarity:        87.5%
Lines Added:       5
Lines Deleted:     3
Lines Unchanged:   40
Total Lines (A):   43
Total Lines (B):   45
============================================================
```

### Exit Codes

- `0`: Files are identical (no differences)
- `1`: Files differ
- `2`: Error occurred (file not found, invalid format, etc.)

## 📖 Additional Resources

- **[EXAMPLES.md](EXAMPLES.md)**: Detailed usage examples and workflows
- **[CONTRIBUTING.md](.github/CONTRIBUTING.md)**: Guidelines for contributors
- **[CHANGELOG.md](config/CHANGELOG.md)**: Version history and release notes
- **[CODE_OF_CONDUCT.md](.github/CODE_OF_CONDUCT.md)**: Community standards

## Visual Guide

### GUI Screenshots

#### Unified Diff View
The unified view displays changes in a GitHub-style format:
- **Green background**: Added lines (+)
- **Red background**: Deleted lines (-)
- **White background**: Context (unchanged lines)
- Toolbar with font controls, copy functions, and navigation
- Statistics panel showing similarity percentage and change counts

#### Side-by-Side View
The side-by-side view shows both documents simultaneously:
- **Left pane**: Original document (File A) with red highlighting for removed content
- **Right pane**: Modified document (File B) with green highlighting for added content
- Synchronized scrolling keeps both panes aligned
- Resizable panes to adjust view

#### Control Panel
- **File Selection**: Browse buttons for each file with full path display
- **Options**: Context lines spinner, ignore case/whitespace checkboxes
- **View Options**: Font size control, line number toggle
- **Statistics**: Real-time similarity and change metrics
- **Navigation**: Previous/Next diff buttons, changes-only filter

### CLI Output Examples

#### Default Unified Diff
```
--- document_v1.docx
+++ document_v2.docx
@@ -12,7 +12,8 @@
 P|Introduction
 P|This document outlines the project requirements.
-P|The deadline is June 2024.
+P|The deadline is July 2024.
+P|Budget has been increased by 15%.
 P|Project team consists of 5 members.
```

#### Colored Output
When using `--color`, the terminal displays:
- Red text for deletions (-)
- Green text for additions (+)
- Cyan text for file headers (---)
- Yellow text for line markers (@@)

#### Statistics Display
```
============================================================
DIFF STATISTICS
============================================================
Similarity:        92.5%
Lines Added:       8
Lines Deleted:     3
Lines Unchanged:   147
Total Lines (A):   150
Total Lines (B):   155

Change Breakdown:
  Additions:       72.73%
  Deletions:       27.27%
============================================================
```

## Use Cases

### Document Review
- Compare contract versions before signing
- Track changes in legal documents
- Review edited manuscripts or reports
- Verify document revisions

### Quality Assurance
- Validate document transformations
- Ensure accuracy in document migrations
- Verify template applications
- Check automated document generation

### Version Control
- Compare document versions in repositories
- Generate change reports for stakeholders
- Document evolution tracking
- Audit document modifications

### Automation
- CI/CD pipeline validation
- Automated document testing
- Batch document comparison
- Schedule periodic document checks

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Troubleshooting

### Common Issues

#### Python/Module Issues
- **Import Error for `docx`**: Install python-docx (not docx): `pip install python-docx`
- **Python version error**: Ensure Python 3.6+ is installed: `python --version`
- **Module not found**: Verify installation: `pip list | grep python-docx`

#### File Issues
- **File Not Found**: 
  - Verify file paths are correct and files exist
  - Use absolute paths or ensure correct working directory
  - Check file permissions (read access required)
- **Not a DOCX file error**: Only `.docx` files are supported (not `.doc`, `.odt`, etc.)
- **Corrupted file**: Try opening the file in Word to verify it's valid

#### Display Issues (GUI)
- **GUI doesn't launch**: 
  - Ensure tkinter is installed: `python -m tkinter`
  - On Linux: `sudo apt-get install python3-tk`
- **DPI Issues on Windows**: Automatically handled, but try running as administrator if issues persist
- **Fonts look wrong**: The GUI uses Consolas for code view - ensure it's installed
- **Window too small**: The GUI starts at 1600x900 - resize or maximize the window

#### Terminal Issues (CLI)
- **Colors not showing**: 
  - Use Windows Terminal on Windows (not Command Prompt)
  - On Linux/Mac, ensure terminal supports ANSI colors
  - Alternative: Use `--html` export for colored output
- **Output truncated**: Redirect to file: `python docxdiff.py file1 file2 > output.txt`
- **Unicode errors**: Ensure terminal encoding is UTF-8

#### Export Issues
- **Permission Error**: 
  - Ensure write permissions for output directory
  - Try saving to a different location (e.g., Desktop)
  - Close any open files with the same name
- **HTML opens incorrectly**: Right-click → Open With → Browser
- **JSON format error**: Use a JSON validator or viewer

### Performance Tips

#### For Large Documents
```bash
# Use quiet mode for quick checks
python docxdiff.py large1.docx large2.docx --quiet

# Export to file instead of terminal
python docxdiff.py large1.docx large2.docx -o output.txt

# Use fewer context lines
python docxdiff.py large1.docx large2.docx -c 0
```

#### For Many Comparisons
- Use JSON export for programmatic processing
- Consider batch scripts (see [EXAMPLES.md](EXAMPLES.md))
- Use `--quiet` mode with exit code checking

### Getting Help

- Check [EXAMPLES.md](EXAMPLES.md) for detailed usage examples
- Review [CONTRIBUTING.md](.github/CONTRIBUTING.md) for development setup
- Open an issue on GitHub for bugs or feature requests
- Use `python docxdiff.py --help` for CLI reference

## Future Enhancements

- Support for comparing PDF and other document formats
- Word-level diff highlighting
- Configuration file support
- Batch comparison of multiple document pairs
- Integration with version control systems
- Advanced filtering options
