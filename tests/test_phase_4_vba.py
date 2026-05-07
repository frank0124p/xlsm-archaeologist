"""Phase 4 VBA analysis tests (unit tests using inline source code)."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsm_archaeologist.analyzers.vba_call_graph import extract_calls
from xlsm_archaeologist.analyzers.vba_procedure_splitter import split_procedures
from xlsm_archaeologist.analyzers.vba_range_detector import detect_range_accesses, detect_triggers
from xlsm_archaeologist.extractors.vba_extractor import extract_vba_modules


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


# ---------------------------------------------------------------------------
# Procedure splitter tests
# ---------------------------------------------------------------------------

_SIMPLE_MODULE = """\
Option Explicit

Public Sub HelloWorld()
    MsgBox "Hello"
End Sub

Private Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
"""

_NESTED_IF_MODULE = """\
Sub CheckValue()
    If x > 0 Then
        If x > 10 Then
            MsgBox "big"
        End If
    End If
End Sub
"""

_PROPERTY_MODULE = """\
Public Property Get Name() As String
    Name = "test"
End Property

Public Property Let Name(val As String)
    mName = val
End Property
"""


def test_split_finds_sub() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    names = [c.name for c in chunks]
    assert "HelloWorld" in names


def test_split_finds_function() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    names = [c.name for c in chunks]
    assert "Add" in names


def test_split_procedure_types() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    types = {c.name: c.procedure_type for c in chunks}
    assert types["HelloWorld"] == "sub"
    assert types["Add"] == "function"


def test_split_public_private() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    pub = {c.name: c.is_public for c in chunks}
    assert pub["HelloWorld"] is True
    assert pub["Add"] is False


def test_split_parameters() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    add_chunk = next(c for c in chunks if c.name == "Add")
    assert len(add_chunk.parameters) == 2
    assert add_chunk.parameters[0].name == "a"
    assert add_chunk.parameters[0].type_hint == "Integer"


def test_split_property() -> None:
    chunks = split_procedures(_PROPERTY_MODULE)
    types = [c.procedure_type for c in chunks]
    assert "property_get" in types
    assert "property_let" in types


def test_split_empty() -> None:
    assert split_procedures("") == []


def test_split_source_lines_preserved() -> None:
    chunks = split_procedures(_SIMPLE_MODULE)
    hello = next(c for c in chunks if c.name == "HelloWorld")
    src = "\n".join(hello.source_lines)
    assert "MsgBox" in src
    assert "End Sub" in src


# ---------------------------------------------------------------------------
# Range detector tests
# ---------------------------------------------------------------------------

_WRITE_CODE = '''\
Sub WriteCell()
    Range("A1").Value = 42
    Sheets("Data").Range("B2").Value = "hello"
End Sub
'''

_READ_CODE = '''\
Sub ReadCell()
    Dim x As Variant
    x = Range("C3").Value
End Sub
'''

_DYNAMIC_CODE = '''\
Sub DynRange()
    Dim lastRow As Long
    lastRow = 10
    Range("A" & lastRow).Value = "X"
End Sub
'''

_CELLS_CODE = '''\
Sub UseCells()
    Cells(1, 1).Value = 100
End Sub
'''


def test_detect_write() -> None:
    reads, writes, notes = detect_range_accesses(_WRITE_CODE)
    assert len(writes) >= 1
    refs = [w.range_ref for w in writes]
    assert "A1" in refs


def test_detect_read() -> None:
    reads, writes, notes = detect_range_accesses(_READ_CODE)
    assert any(r.range_ref == "C3" for r in reads)


def test_detect_dynamic_range() -> None:
    reads, writes, notes = detect_range_accesses(_DYNAMIC_CODE)
    assert len(notes) >= 1
    assert any("lastRow" in n or "Dynamic" in n for n in notes)


def test_detect_cells_write() -> None:
    reads, writes, notes = detect_range_accesses(_CELLS_CODE)
    assert any(w.via == "cells_method" for w in writes)


def test_detect_no_false_positives_in_comments() -> None:
    code = "' Range(\"A1\").Value = 99\n"
    reads, writes, notes = detect_range_accesses(code)
    assert len(writes) == 0


# ---------------------------------------------------------------------------
# Event trigger detection
# ---------------------------------------------------------------------------


def test_detect_worksheet_change() -> None:
    code = 'Sub Worksheet_Change(ByVal Target As Range)\n  MsgBox "changed"\nEnd Sub'
    triggers = detect_triggers("Worksheet_Change", code)
    assert len(triggers) == 1
    assert triggers[0].event == "Worksheet_Change"


def test_detect_workbook_open() -> None:
    code = "Sub Workbook_Open()\n  MsgBox \"opened\"\nEnd Sub"
    triggers = detect_triggers("Workbook_Open", code)
    assert len(triggers) == 1


def test_detect_intersect_target() -> None:
    code = (
        'Sub Worksheet_Change(ByVal Target As Range)\n'
        '  If Not Intersect(Target, Range("B2:B10")) Is Nothing Then\n'
        '    MsgBox "hit"\n  End If\nEnd Sub'
    )
    triggers = detect_triggers("Worksheet_Change", code)
    assert triggers[0].target == "B2:B10"


def test_no_trigger_for_normal_sub() -> None:
    triggers = detect_triggers("MyHelper", "Sub MyHelper()\nEnd Sub")
    assert triggers == []


# ---------------------------------------------------------------------------
# Call graph extraction
# ---------------------------------------------------------------------------

_CALL_CODE = """\
Sub Main()
    Call Helper1
    Helper2
End Sub

Sub Helper1()
End Sub

Sub Helper2()
End Sub

Sub NotCalled()
End Sub
"""


def test_extract_calls_direct() -> None:
    all_names = {"Helper1", "Helper2", "NotCalled", "Main"}
    calls = extract_calls(_CALL_CODE, all_names)
    assert "Helper1" in calls
    assert "Helper2" in calls


def test_extract_calls_excludes_not_called() -> None:
    all_names = {"Helper1", "Helper2", "NotCalled"}
    calls = extract_calls("Sub Main()\n    Call Helper1\nEnd Sub", all_names)
    assert "NotCalled" not in calls


def test_extract_calls_sorted() -> None:
    all_names = {"Zzz", "Aaa", "Mmm"}
    calls = extract_calls("Sub X()\n    Call Zzz\n    Call Aaa\nEnd Sub", all_names)
    assert calls == sorted(calls)


# ---------------------------------------------------------------------------
# VBA extractor (smoke test on fixture without real VBA binary)
# ---------------------------------------------------------------------------


def test_vba_extractor_no_macros(fixtures_dir: Path) -> None:
    """simple.xlsm has no VBA — extractor should yield nothing."""
    warnings: list[str] = []
    modules = list(extract_vba_modules(fixtures_dir / "simple.xlsm", warnings))
    assert modules == []
