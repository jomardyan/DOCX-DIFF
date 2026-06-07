import sys
import json
from pathlib import Path
from unittest.mock import MagicMock, patch, mock_open
import pytest
import difflib
import tkinter as tk

# Import functions to test from docxdiff
from docxdiff import (
    Config,
    validate_file,
    iter_block_text,
    load_docx_lines,
    print_unified_diff,
    print_colored_diff,
    calculate_diff_stats,
    export_to_html,
    export_to_json,
    print_side_by_side,
    print_statistics,
    main,
    DocxDiffGUI,
)


def test_validate_file_not_found(tmp_path):
    non_existent = tmp_path / "does_not_exist.docx"
    is_valid, msg = validate_file(non_existent)
    assert not is_valid
    assert "File not found" in msg


def test_validate_file_not_a_file(tmp_path):
    is_valid, msg = validate_file(tmp_path)
    assert not is_valid
    assert "Not a file" in msg


def test_validate_file_unsupported_extension(tmp_path):
    invalid_file = tmp_path / "test.txt"
    invalid_file.write_text("hello")
    is_valid, msg = validate_file(invalid_file)
    assert not is_valid
    assert "Unsupported file type" in msg


def test_validate_file_too_large(tmp_path, monkeypatch):
    large_file = tmp_path / "large.docx"
    large_file.write_text("fake docx")

    # Mock Config.MAX_FILE_SIZE_MB to be extremely small for testing
    monkeypatch.setattr(Config, "MAX_FILE_SIZE_MB", 0.000001)

    is_valid, msg = validate_file(large_file)
    assert not is_valid
    assert "File too large" in msg


def test_validate_file_valid(tmp_path):
    valid_file = tmp_path / "valid.docx"
    valid_file.write_text("fake docx")
    is_valid, msg = validate_file(valid_file)
    assert is_valid
    assert msg == ""


def test_iter_block_text():
    # Mock a docx Document
    mock_doc = MagicMock()

    # Mock paragraphs
    mock_p1 = MagicMock()
    mock_p1.text = "  Paragraph 1  "
    mock_p2 = MagicMock()
    mock_p2.text = None
    mock_doc.paragraphs = [mock_p1, mock_p2]

    # Mock tables
    mock_table = MagicMock()
    mock_row1 = MagicMock()
    mock_cell1 = MagicMock()
    mock_cell1.text = "Cell 1\nwith space"
    mock_row1.cells = [mock_cell1]
    mock_table.rows = [mock_row1]
    mock_doc.tables = [mock_table]

    blocks = list(iter_block_text(mock_doc))

    assert len(blocks) == 3
    assert blocks[0] == "P|Paragraph 1"
    assert blocks[1] == "P|"
    assert blocks[2] == "T1R1C1|Cell 1 with space"


def test_load_docx_lines_validation_failure(tmp_path):
    non_existent = tmp_path / "no.docx"
    with pytest.raises(ValueError, match="File not found"):
        load_docx_lines(non_existent)


@patch("docxdiff.Document")
def test_load_docx_lines_success(mock_document_class, tmp_path):
    valid_file = tmp_path / "valid.docx"
    valid_file.write_text("fake zip")

    # Mock the Document object returned by Document(path)
    mock_doc = MagicMock()
    mock_p = MagicMock()
    mock_p.text = "Hello World"
    mock_doc.paragraphs = [mock_p]
    mock_doc.tables = []
    mock_document_class.return_value = mock_doc

    lines = load_docx_lines(valid_file)
    assert lines == ["P|Hello World"]


@patch("docxdiff.Document")
def test_load_docx_lines_runtime_error(mock_document_class, tmp_path):
    mock_document_class.side_effect = Exception("Corrupted file")
    valid_file = tmp_path / "corrupt.docx"
    valid_file.write_text("fake corrupt zip")

    with pytest.raises(RuntimeError, match="Failed to load document"):
        load_docx_lines(valid_file)


def test_print_unified_diff(capsys):
    a_lines = ["line 1", "line 2"]
    b_lines = ["line 1", "line 2 changed"]

    any_diff = print_unified_diff(a_lines, b_lines, "file_a.docx", "file_b.docx")
    assert any_diff

    captured = capsys.readouterr()
    assert "--- file_a.docx" in captured.out
    assert "+++ file_b.docx" in captured.out
    assert "-line 2" in captured.out
    assert "+line 2 changed" in captured.out


def test_print_colored_diff(capsys):
    a_lines = ["line 1", "line 2"]
    b_lines = ["line 1", "line 2 changed"]

    any_diff = print_colored_diff(a_lines, b_lines, "file_a.docx", "file_b.docx")
    assert any_diff

    captured = capsys.readouterr()
    # Check that ANSI color codes or correct diff lines are printed
    assert "file_a.docx" in captured.out
    assert "file_b.docx" in captured.out
    assert "line 2" in captured.out


def test_calculate_diff_stats():
    a_lines = ["line 1", "line 2", "line 3"]
    b_lines = ["line 1", "line 2 changed", "line 4"]

    stats = calculate_diff_stats(a_lines, b_lines)
    assert stats["additions"] == 2
    assert stats["deletions"] == 2
    assert stats["unchanged"] == 1
    assert stats["total_lines_a"] == 3
    assert stats["total_lines_b"] == 3
    assert isinstance(stats["similarity"], float)


def test_export_to_html(tmp_path):
    a_lines = ["line 1"]
    b_lines = ["line 2"]
    out_html = tmp_path / "diff.html"

    success = export_to_html(a_lines, b_lines, "a.docx", "b.docx", str(out_html))
    assert success
    assert out_html.exists()
    content = out_html.read_text(encoding="utf-8")
    assert "DOCX Diff v" in content
    assert "a.docx" in content
    assert "b.docx" in content


def test_export_to_json(tmp_path):
    a_lines = ["line 1"]
    b_lines = ["line 2"]
    out_json = tmp_path / "diff.json"

    success = export_to_json(a_lines, b_lines, "a.docx", "b.docx", str(out_json))
    assert success
    assert out_json.exists()

    data = json.loads(out_json.read_text(encoding="utf-8"))
    assert "metadata" in data
    assert "statistics" in data
    assert "changes" in data
    assert data["statistics"]["similarity"] == 0.0


def test_print_side_by_side(capsys):
    a_lines = ["line 1", "line 2"]
    b_lines = ["line 1", "line 3"]

    print_side_by_side(a_lines, b_lines, "a.docx", "b.docx", width=40)
    captured = capsys.readouterr()
    assert "a.docx" in captured.out
    assert "b.docx" in captured.out
    assert "line 1" in captured.out


def test_print_statistics(capsys):
    stats = {
        "similarity": 80.0,
        "additions": 1,
        "deletions": 1,
        "unchanged": 4,
        "total_lines_a": 5,
        "total_lines_b": 5,
    }
    print_statistics(stats, verbose=True)
    captured = capsys.readouterr()
    assert "DIFF STATISTICS" in captured.out
    assert "Similarity:        80.0%" in captured.out
    assert "Change Breakdown:" in captured.out


@patch("docxdiff.load_docx_lines")
def test_main_cli_success(mock_load, capsys, tmp_path):
    mock_load.side_effect = [["line 1", "line 2"], ["line 1", "line 2 changed"]]

    file_a = tmp_path / "a.docx"
    file_b = tmp_path / "b.docx"
    file_a.write_text("a")
    file_b.write_text("b")

    test_args = ["docxdiff.py", str(file_a), str(file_b)]
    with patch.object(sys, "argv", test_args):
        exit_code = main()
        assert exit_code == 1  # Differences found

    captured = capsys.readouterr()
    assert "line 2" in captured.out


@patch("docxdiff.load_docx_lines")
def test_main_cli_same_files(mock_load, capsys, tmp_path):
    file_a = tmp_path / "a.docx"
    file_a.write_text("a")

    test_args = ["docxdiff.py", str(file_a), str(file_a)]
    with patch.object(sys, "argv", test_args):
        exit_code = main()
        assert exit_code == 2

    captured = capsys.readouterr()
    assert "Error: Both files point to the same location" in captured.err


def test_display_side_by_side_alignment():
    # Create a mock instance of DocxDiffGUI without calling __init__
    gui = MagicMock(spec=DocxDiffGUI)

    # Mock text widgets and their insert methods
    gui.left_text = MagicMock()
    gui.right_text = MagicMock()

    # Mock show_line_numbers_var
    gui.show_line_numbers_var = MagicMock()
    gui.show_line_numbers_var.get.return_value = True

    # Prepare inputs: 'replace' scenario
    # a_lines has 2 lines, b_lines has 1 line
    a_lines = ["old 1", "old 2"]
    b_lines = ["new 1"]

    matcher = difflib.SequenceMatcher(None, a_lines, b_lines)

    # Call display_side_by_side using the mocked gui instance
    DocxDiffGUI.display_side_by_side(gui, a_lines, b_lines, matcher)

    # Check that delete was called on both text widgets
    gui.left_text.delete.assert_any_call(1.0, tk.END)
    gui.right_text.delete.assert_any_call(1.0, tk.END)

    # In the replace block, 2 lines are removed from left, and 1 line is added to right
    # Because of our alignment fix, 1 empty spacer line should have been added to the right side!
    # Let's count how many times "     │ \n" was inserted into the right text widget.
    # It should be exactly 1 time!
    insert_calls = gui.right_text.insert.call_args_list
    spacer_calls = [c for c in insert_calls if c[0][1] == "     │ \n"]
    assert len(spacer_calls) == 1


def test_gui_swap_files():
    # Mock DocxDiffGUI instance
    gui = MagicMock(spec=DocxDiffGUI)
    gui.file_a_var = MagicMock()
    gui.file_b_var = MagicMock()
    gui.status_bar = MagicMock()

    gui.file_a_var.get.return_value = "path/to/a.docx"
    gui.file_b_var.get.return_value = "path/to/b.docx"

    DocxDiffGUI.swap_files(gui)

    gui.file_a_var.set.assert_called_once_with("path/to/b.docx")
    gui.file_b_var.set.assert_called_once_with("path/to/a.docx")
    gui.status_bar.config.assert_called_once_with(text="Swapped File A and File B")


def test_gui_add_to_history():
    gui = MagicMock(spec=DocxDiffGUI)
    gui.history = []
    gui.save_history = MagicMock()
    gui.update_history_comboboxes = MagicMock()

    # Test adding a path
    DocxDiffGUI.add_to_history(gui, "test1.docx")
    assert str(Path("test1.docx").resolve()) in gui.history

    # Test limiting history
    gui.history = [f"file_{i}.docx" for i in range(15)]
    DocxDiffGUI.add_to_history(gui, "new.docx")
    assert len(gui.history) == 10
    assert str(Path("new.docx").resolve()) in gui.history


@patch("docxdiff.messagebox")
@patch("docxdiff.filedialog.asksaveasfilename")
@patch("docxdiff.load_docx_lines")
@patch("docxdiff.export_to_json")
def test_gui_export_json_success(mock_export, mock_load, mock_ask, mock_msg):
    gui = MagicMock(spec=DocxDiffGUI)
    gui.current_diff_lines = ["some diff"]
    gui.file_a_var = MagicMock()
    gui.file_b_var = MagicMock()
    gui.file_a_var.get.return_value = "a.docx"
    gui.file_b_var.get.return_value = "b.docx"
    gui.ignore_whitespace_var = MagicMock()
    gui.ignore_whitespace_var.get.return_value = False
    gui.ignore_case_var = MagicMock()
    gui.ignore_case_var.get.return_value = False
    gui.context_var = MagicMock()
    gui.context_var.get.return_value = 3
    gui.status_bar = MagicMock()

    mock_ask.return_value = "output.json"
    mock_load.side_effect = [["line 1"], ["line 2"]]

    DocxDiffGUI.export_json(gui)

    mock_export.assert_called_once()
    gui.status_bar.config.assert_called_once_with(text="Exported to output.json")
    mock_msg.showinfo.assert_called_once_with(
        "Success", "Results exported to output.json"
    )
