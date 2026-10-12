"""Aplikační logika: validace + stav. Žádné Qt, žádné COM přímo.

SheetController je synchronní a testovatelný s FakeExcelAdapter.
V produkci ho volá ExcelWorker v dedikovaném vlákně.
Nikdy nevolá Save/SaveAs.
"""

from __future__ import annotations

from core.errors import (
    CannotHideAllSheets,
    NoActiveWorkbook,
    SheetNotFound,
    ValidationError,
    WorkbookProtected,
)
from core.models import SheetInfo, Visibility, WorkbookInfo
from excel_backend.base import ExcelAdapter

FORBIDDEN_NAME_CHARS = set('[]:*?/\\')
MAX_SHEET_NAME_LEN = 31


def validate_sheet_name(new_name: str, existing: list[str] | None = None, old_name: str = "") -> str:
    """Ověří název listu podle limitů Excelu. Vrací očištěný název."""
    cleaned = (new_name or "").strip()
    if not cleaned:
        raise ValidationError("Název listu nesmí být prázdný.")
    if len(cleaned) > MAX_SHEET_NAME_LEN:
        raise ValidationError(f"Název listu smí mít nejvýše {MAX_SHEET_NAME_LEN} znaků.")
    bad = sorted(c for c in cleaned if c in FORBIDDEN_NAME_CHARS)
    if bad:
        raise ValidationError("Název listu nesmí obsahovat znaky: [ ] : * ? / \\")
    if cleaned.startswith("'") or cleaned.endswith("'"):
        raise ValidationError("Název listu nesmí začínat ani končit apostrofem.")
    if existing is not None:
        lowered = {n.lower() for n in existing}
        if cleaned.lower() != old_name.lower() and cleaned.lower() in lowered:
            raise ValidationError(f"List s názvem „{cleaned}“ již existuje.")
    return cleaned


class SheetController:
    def __init__(self, adapter: ExcelAdapter):
        self._adapter = adapter
        self.workbook: WorkbookInfo | None = None
        self.sheets: list[SheetInfo] = []

    # -- čtení ---------------------------------------------------------
    def refresh(self) -> tuple[WorkbookInfo, list[SheetInfo]]:
        self.workbook = self._adapter.get_workbook_info()
        self.sheets = self._adapter.list_sheets()
        return self.workbook, self.sheets

    def connect(self) -> tuple[WorkbookInfo, list[SheetInfo]]:
        self.workbook = self._adapter.connect()
        self.sheets = self._adapter.list_sheets()
        return self.workbook, self.sheets

    def _require_sheets(self) -> list[SheetInfo]:
        if not self.sheets:
            # Zkusit obnovit, aby chyba byla aktuální (žádný sešit vs. prázdno).
            self.refresh()
        return self.sheets

    def find(self, name: str) -> SheetInfo:
        for s in self._require_sheets():
            if s.name.lower() == name.lower():
                return s
        raise SheetNotFound(name)

    # -- operace -------------------------------------------------------
    def activate(self, name: str) -> None:
        self.find(name)  # včasná srozumitelná chyba
        self._adapter.activate_sheet(name)
        # Aktualizovat příznak is_active lokálně bez nutnosti refresh:
        self.sheets = [SheetInfo(s.name, s.index, s.sheet_type, s.visibility, s.name.lower() == name.lower()) for s in self.sheets]

    def add(self, name: str | None = None) -> SheetInfo:
        existing = [s.name for s in self._require_sheets()]
        cleaned: str | None = None
        if name is not None and name.strip() != "":
            cleaned = validate_sheet_name(name, existing)
        self._guard_not_protected()
        created = self._adapter.add_sheet(cleaned)
        self.refresh()
        return created

    def rename(self, old_name: str, new_name: str) -> SheetInfo:
        existing = [s.name for s in self._require_sheets()]
        self.find(old_name)
        cleaned = validate_sheet_name(new_name, existing, old_name=old_name)
        self._guard_not_protected()
        renamed = self._adapter.rename_sheet(old_name, cleaned)
        self.refresh()
        return renamed

    def delete(self, name: str, confirmed: bool = False) -> None:
        if not confirmed:
            raise ValidationError("Odstranění listu musí být výslovně potvrzeno v dialogu.")
        self.find(name)
        self._guard_not_protected()
        self._adapter.delete_sheet(name)
        self.refresh()

    def move(self, name: str, new_index_1based: int) -> SheetInfo:
        sheets = self._require_sheets()
        self.find(name)
        if not 1 <= int(new_index_1based) <= len(sheets):
            raise ValidationError(f"Cílová pozice musí být v rozsahu 1–{len(sheets)}.")
        self._guard_not_protected()
        moved = self._adapter.move_sheet(name, int(new_index_1based))
        self.refresh()
        return moved

    def move_left(self, name: str) -> SheetInfo:
        sheets = self._require_sheets()
        cur = self.find(name)
        if cur.index <= 1:
            raise ValidationError("List je již první a nelze ho přesunout doleva.")
        return self.move(name, cur.index - 1)

    def move_right(self, name: str) -> SheetInfo:
        sheets = self._require_sheets()
        cur = self.find(name)
        if cur.index >= len(sheets):
            raise ValidationError("List je již poslední a nelze ho přesunout doprava.")
        return self.move(name, cur.index + 1)

    def set_visible(self, name: str, visible: bool) -> SheetInfo:
        sheets = self._require_sheets()
        cur = self.find(name)
        if not visible and cur.visibility == Visibility.VISIBLE:
            visible_count = sum(1 for s in sheets if s.visibility == Visibility.VISIBLE)
            if visible_count <= 1:
                raise CannotHideAllSheets()
        self._guard_not_protected()
        updated = self._adapter.set_sheet_visible(name, visible)
        self.refresh()
        return updated

    def _guard_not_protected(self) -> None:
        if self.workbook is not None and self.workbook.protect_structure:
            raise WorkbookProtected()
        # Pokud stav neznáme, necháme adaptér, aby ochranu detekoval z COM.

    # -- pomocné pro UI -------------------------------------------------
    def index_to_select_after_delete(self, deleted_name: str) -> int:
        """0-based řádek k vybrání po smazání (soused). Volat před refreshem."""
        names = [s.name for s in self.sheets]
        try:
            pos = next(i for i, n in enumerate(names) if n.lower() == deleted_name.lower())
        except StopIteration:
            return 0
        # Po smazání vyber stejný index, nebo předchozí pokud byl poslední.
        return min(pos, max(0, len(names) - 2))
