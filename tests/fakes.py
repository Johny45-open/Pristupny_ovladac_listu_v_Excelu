"""Falešný adaptér Excelu pro testy bez instalovaného Excelu."""

from __future__ import annotations

from core.errors import (
    ExcelNotRunning,
    NoActiveWorkbook,
    SheetNotFound,
    SheetOperationDenied,
    WorkbookProtected,
)
from core.models import SheetInfo, SheetType, Visibility, WorkbookInfo
from excel_backend.base import ExcelAdapter


class FakeExcelAdapter(ExcelAdapter):
    def __init__(
        self,
        sheets: list[SheetInfo] | None = None,
        workbook_name: str = "Test.xlsx",
        protected: bool = False,
        running: bool = True,
        has_workbook: bool = True,
    ):
        if sheets is None:
            sheets = [
                SheetInfo("List1", 1, SheetType.WORKSHEET, Visibility.VISIBLE, True),
                SheetInfo("List2", 2, SheetType.WORKSHEET, Visibility.VISIBLE, False),
                SheetInfo("Graf1", 3, SheetType.CHART, Visibility.VISIBLE, False),
            ]
        self._sheets: list[SheetInfo] = list(sheets)
        self._wb_name = workbook_name
        self._protected = protected
        self.running = running
        self.has_workbook = has_workbook
        self.save_calls = 0  # nikdy se nesmí zvýšit — test zákazu ukládání
        self.calls: list[str] = []

    # -- pomocné ---------------------------------------------------------
    def _ensure(self) -> None:
        if not self.running:
            raise ExcelNotRunning()
        if not self.has_workbook:
            raise NoActiveWorkbook()

    def _guard_write(self) -> None:
        self._ensure()
        if self._protected:
            raise WorkbookProtected()

    def _reindex(self) -> None:
        self._sheets = [
            SheetInfo(s.name, i + 1, s.sheet_type, s.visibility, s.is_active)
            for i, s in enumerate(self._sheets)
        ]

    def _find_pos(self, name: str) -> int:
        for i, s in enumerate(self._sheets):
            if s.name.lower() == name.lower():
                return i
        raise SheetNotFound(name)

    # -- interface -------------------------------------------------------
    def connect(self) -> WorkbookInfo:
        self.calls.append("connect")
        self._ensure()
        return self.get_workbook_info()

    def get_workbook_info(self) -> WorkbookInfo:
        self._ensure()
        return WorkbookInfo(
            name=self._wb_name,
            full_name=f"C:\\Fake\\{self._wb_name}",
            protect_structure=self._protected,
            sheet_count=len(self._sheets),
        )

    def list_sheets(self) -> list[SheetInfo]:
        self._ensure()
        return list(self._sheets)

    def activate_sheet(self, name: str) -> None:
        self._ensure()
        pos = self._find_pos(name)
        if self._sheets[pos].visibility != Visibility.VISIBLE:
            raise SheetOperationDenied(f"List „{name}“ je skrytý a nelze ho aktivovat.")
        self._sheets = [
            SheetInfo(s.name, s.index, s.sheet_type, s.visibility, i == pos)
            for i, s in enumerate(self._sheets)
        ]

    def add_sheet(self, name: str | None = None) -> SheetInfo:
        self._guard_write()
        new_name = name or f"List{len(self._sheets) + 1}"
        if any(s.name.lower() == new_name.lower() for s in self._sheets):
            raise SheetOperationDenied(f"List „{new_name}“ již existuje.")
        created = SheetInfo(new_name, len(self._sheets) + 1, SheetType.WORKSHEET, Visibility.VISIBLE, False)
        self._sheets.append(created)
        self._reindex()
        return created

    def rename_sheet(self, old_name: str, new_name: str) -> SheetInfo:
        self._guard_write()
        pos = self._find_pos(old_name)
        s = self._sheets[pos]
        renamed = SheetInfo(new_name, s.index, s.sheet_type, s.visibility, s.is_active)
        self._sheets[pos] = renamed
        return renamed

    def delete_sheet(self, name: str) -> None:
        self._guard_write()
        pos = self._find_pos(name)
        del self._sheets[pos]
        self._reindex()
        if self._sheets and not any(s.is_active for s in self._sheets):
            s0 = self._sheets[0]
            self._sheets[0] = SheetInfo(s0.name, s0.index, s0.sheet_type, s0.visibility, True)

    def move_sheet(self, name: str, new_index_1based: int) -> SheetInfo:
        self._guard_write()
        pos = self._find_pos(name)
        target = max(1, min(int(new_index_1based), len(self._sheets))) - 1
        item = self._sheets.pop(pos)
        self._sheets.insert(target, item)
        self._reindex()
        for s in self._sheets:
            if s.name.lower() == name.lower():
                return s
        raise SheetNotFound(name)  # pragma: no cover

    def set_sheet_visible(self, name: str, visible: bool) -> SheetInfo:
        self._guard_write()
        pos = self._find_pos(name)
        s = self._sheets[pos]
        updated = SheetInfo(
            s.name, s.index, s.sheet_type,
            Visibility.VISIBLE if visible else Visibility.HIDDEN, s.is_active,
        )
        self._sheets[pos] = updated
        return updated
