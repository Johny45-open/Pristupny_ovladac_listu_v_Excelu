"""Abstraktní kontrakt adaptéru Excelu.

Umožňuje testovat controller s FakeExcelAdapter bez instalovaného Excelu.
COM implementace je pouze v com_adapter.py.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

from core.models import SheetInfo, WorkbookInfo


class ExcelAdapter(ABC):
    @abstractmethod
    def connect(self) -> WorkbookInfo:
        """Připojí se k běžící instanci a vrátí info o aktivním sešitě."""
        raise NotImplementedError

    @abstractmethod
    def get_workbook_info(self) -> WorkbookInfo:
        raise NotImplementedError

    @abstractmethod
    def list_sheets(self) -> list[SheetInfo]:
        raise NotImplementedError

    @abstractmethod
    def activate_sheet(self, name: str) -> None:
        raise NotImplementedError

    @abstractmethod
    def add_sheet(self, name: str | None = None) -> SheetInfo:
        raise NotImplementedError

    @abstractmethod
    def rename_sheet(self, old_name: str, new_name: str) -> SheetInfo:
        raise NotImplementedError

    @abstractmethod
    def delete_sheet(self, name: str) -> None:
        raise NotImplementedError

    @abstractmethod
    def move_sheet(self, name: str, new_index_1based: int) -> SheetInfo:
        raise NotImplementedError

    @abstractmethod
    def set_sheet_visible(self, name: str, visible: bool) -> SheetInfo:
        raise NotImplementedError
