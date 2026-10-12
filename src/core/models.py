"""Datové modely sdílené mezi backendem, controllerem a GUI.

Obsahuje pouze čistá data (dataclasses + enumy). Nikdy zde nedržíme
živé COM objekty — ty zůstávají výhradně v excel_backend/com_adapter.py
a v příslušném worker vlákně.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class SheetType(str, Enum):
    WORKSHEET = "pracovní list"
    CHART = "graf"
    DIALOG = "dialogový list"
    MACRO = "makro list"
    UNKNOWN = "jiný list"


class Visibility(str, Enum):
    VISIBLE = "viditelný"
    HIDDEN = "skrytý"
    VERY_HIDDEN = "velmi skrytý"


@dataclass(frozen=True)
class SheetInfo:
    """Jeden list v sešitě. index je 1-based pozice v Excelu."""

    name: str
    index: int
    sheet_type: SheetType = SheetType.WORKSHEET
    visibility: Visibility = Visibility.VISIBLE
    is_active: bool = False

    def display_text(self) -> str:
        """Jeden řádek pro QListWidget, čitelný pro NVDA bez spoléhání na barvu."""
        parts = [self.name, self.sheet_type.value, self.visibility.value]
        parts.append("aktivní" if self.is_active else "neaktivní")
        return ", ".join(parts)


@dataclass(frozen=True)
class WorkbookInfo:
    name: str
    full_name: str = ""
    protect_structure: bool = False
    sheet_count: int = 0
