"""Validační a bezpečnostní testy: nesmí se zjednodušovat."""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from controller.sheet_controller import SheetController
from core.errors import SheetOperationDenied, ValidationError
from core.models import SheetInfo, SheetType, Visibility
from excel_backend.base import ExcelAdapter
from tests.fakes import FakeExcelAdapter


class ExplodingAdapter(ExcelAdapter):
    """Simuluje pád COM vrstvy — controller nesmí spadnout, jen předat výjimku."""

    def connect(self):
        raise SheetOperationDenied("simulovaná chyba COM")

    def get_workbook_info(self):
        raise SheetOperationDenied("simulovaná chyba COM")

    def list_sheets(self):
        raise SheetOperationDenied("simulovaná chyba COM")

    def activate_sheet(self, name):
        raise SheetOperationDenied("simulovaná chyba COM")

    def add_sheet(self, name=None):
        raise SheetOperationDenied("simulovaná chyba COM")

    def rename_sheet(self, old_name, new_name):
        raise SheetOperationDenied("simulovaná chyba COM")

    def delete_sheet(self, name):
        raise SheetOperationDenied("simulovaná chyba COM")

    def move_sheet(self, name, new_index_1based):
        raise SheetOperationDenied("simulovaná chyba COM")

    def set_sheet_visible(self, name, visible):
        raise SheetOperationDenied("simulovaná chyba COM")


def test_com_error_does_not_crash_controller():
    c = SheetController(ExplodingAdapter())
    try:
        c.refresh()
        assert False
    except SheetOperationDenied:
        pass  # očekáváno, žádný pád


def test_hidden_sheet_activation_denied_in_fake():
    adapter = FakeExcelAdapter(sheets=[
        SheetInfo("A", 1, SheetType.WORKSHEET, Visibility.VISIBLE, True),
        SheetInfo("B", 2, SheetType.WORKSHEET, Visibility.HIDDEN, False),
    ])
    c = SheetController(adapter)
    c.refresh()
    try:
        c.activate("B")
        assert False, "aktivace skrytého listu musí selhat"
    except SheetOperationDenied:
        pass


def test_very_hidden_roundtrip_logic():
    adapter = FakeExcelAdapter(sheets=[
        SheetInfo("A", 1, SheetType.WORKSHEET, Visibility.VISIBLE, True),
        SheetInfo("V", 2, SheetType.WORKSHEET, Visibility.VERY_HIDDEN, False),
    ])
    c = SheetController(adapter)
    c.refresh()
    # Zobrazení velmi skrytého listu je povoleno:
    c.set_visible("V", True)
    assert c.find("V").visibility == Visibility.VISIBLE


def test_chart_sheet_type_preserved():
    c = SheetController(FakeExcelAdapter())
    c.refresh()
    graf = c.find("Graf1")
    assert graf.sheet_type == SheetType.CHART
