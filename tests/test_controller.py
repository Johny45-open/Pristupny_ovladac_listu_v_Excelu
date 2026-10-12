"""Testy controlleru s FakeExcelAdapter (bez Excelu)."""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from controller.sheet_controller import SheetController, validate_sheet_name
from core.errors import (
    CannotHideAllSheets,
    ExcelNotRunning,
    NoActiveWorkbook,
    SheetNotFound,
    ValidationError,
    WorkbookProtected,
    user_message,
)
from core.models import SheetInfo, SheetType, Visibility
from tests.fakes import FakeExcelAdapter


def make():
    adapter = FakeExcelAdapter()
    c = SheetController(adapter)
    c.refresh()
    return c, adapter


def test_refresh_lists_sheets():
    c, _ = make()
    assert c.workbook is not None and c.workbook.name == "Test.xlsx"
    assert [s.name for s in c.sheets] == ["List1", "List2", "Graf1"]
    assert c.sheets[0].is_active


def test_activate_changes_active():
    c, _ = make()
    c.activate("List2")
    assert [s.is_active for s in c.sheets] == [False, True, False]


def test_add_and_rename_and_focus_data():
    c, _ = make()
    created = c.add("Nový")
    assert created.name == "Nový"
    assert "Nový" in [s.name for s in c.sheets]
    renamed = c.rename("Nový", "Přejmenovaný")
    assert renamed.name == "Přejmenovaný"


def test_delete_requires_confirmation():
    c, _ = make()
    try:
        c.delete("List2", confirmed=False)
        assert False, "musí vyžadovat confirmed=True"
    except ValidationError:
        pass
    c.delete("List2", confirmed=True)
    assert "List2" not in [s.name for s in c.sheets]


def test_index_to_select_after_delete():
    c, _ = make()
    assert c.index_to_select_after_delete("List2") == 1
    assert c.index_to_select_after_delete("Graf1") == 1  # poslední -> předchozí


def test_cannot_hide_last_visible():
    adapter = FakeExcelAdapter(sheets=[
        SheetInfo("Jediný", 1, SheetType.WORKSHEET, Visibility.VISIBLE, True),
    ])
    c = SheetController(adapter)
    c.refresh()
    try:
        c.set_visible("Jediný", False)
        assert False, "musí odmítnout skrýt poslední viditelný"
    except CannotHideAllSheets:
        pass
    # Adaptér vůbec nebyl volán pro skrytí:
    assert adapter.list_sheets()[0].visibility == Visibility.VISIBLE


def test_hide_and_show():
    c, _ = make()
    c.set_visible("List2", False)
    assert c.find("List2").visibility == Visibility.HIDDEN
    c.set_visible("List2", True)
    assert c.find("List2").visibility == Visibility.VISIBLE


def test_move_left_right_and_to_position():
    c, _ = make()
    c.move_left("List2")
    assert [s.name for s in c.sheets][0] == "List2"
    c.move_right("List2")
    assert [s.name for s in c.sheets][1] == "List2"
    c.move("Graf1", 1)
    assert [s.name for s in c.sheets][0] == "Graf1"


def test_move_bounds():
    c, _ = make()
    try:
        c.move_left("List1")
        assert False
    except ValidationError:
        pass
    try:
        c.move("List1", 99)
        assert False
    except ValidationError:
        pass


def test_rename_validation():
    for bad in ["", "   ", "A" * 32, "List[1]", "a/b", "x:y", "a*b", "a?b"]:
        try:
            validate_sheet_name(bad, ["List1"])
            assert False, f"měl odmítnout: {bad!r}"
        except ValidationError:
            pass
    try:
        validate_sheet_name("list1", ["List1"], old_name="List2")
        assert False, "duplicita (case-insensitive) musí selhat"
    except ValidationError:
        pass
    assert validate_sheet_name("  OK  ", []) == "OK"


def test_protected_workbook_blocks_mutations():
    adapter = FakeExcelAdapter(protected=True)
    c = SheetController(adapter)
    c.refresh()
    assert c.workbook is not None and c.workbook.protect_structure
    for fn in (lambda: c.add("X"), lambda: c.rename("List1", "Y"),
               lambda: c.delete("List1", confirmed=True), lambda: c.move("List1", 2),
               lambda: c.set_visible("List1", False)):
        try:
            fn()
            assert False, "chráněný sešit musí blokovat"
        except WorkbookProtected:
            pass


def test_no_excel_no_workbook_sheet_not_found():
    c = SheetController(FakeExcelAdapter(running=False))
    try:
        c.refresh()
        assert False
    except ExcelNotRunning:
        pass
    c2 = SheetController(FakeExcelAdapter(has_workbook=False))
    try:
        c2.refresh()
        assert False
    except NoActiveWorkbook:
        pass
    c3, _ = make()
    try:
        c3.activate("Neexistuje")
        assert False
    except SheetNotFound:
        pass


def test_no_autosave_call():
    _, adapter = make()
    assert adapter.save_calls == 0


def test_user_messages_have_specific_titles():
    for exc, title_part in [
        (ExcelNotRunning(), "Excel"),
        (NoActiveWorkbook(), "sešit"),
        (WorkbookProtected(), "chráněn"),
        (CannotHideAllSheets(), "skrýt"),
    ]:
        title, msg = user_message(exc)
        assert title_part.lower() in title.lower()
        assert "Upozornění" != title  # nikdy generický titulek
        assert len(msg) > 10
