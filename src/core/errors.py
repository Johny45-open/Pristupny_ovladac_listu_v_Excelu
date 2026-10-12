"""Doménové výjimky a překlad na české texty pro dialogy.

Důvod: COM chyby (pywintypes.com_error) jsou anglicky a kryptické.
GUI nikdy nezobrazuje text COM výjimky přímo, vždy (titulek, zpráva)
z user_message().
"""

from __future__ import annotations


class ExcelError(Exception):
    """Základ všech doménových chyb ovladače."""


class ExcelNotRunning(ExcelError):
    pass


class NoActiveWorkbook(ExcelError):
    pass


class WorkbookProtected(ExcelError):
    pass


class SheetNotFound(ExcelError):
    def __init__(self, name: str = ""):
        super().__init__(name)
        self.sheet_name = name


class SheetOperationDenied(ExcelError):
    pass


class CannotHideAllSheets(ExcelError):
    pass


class ValidationError(ExcelError):
    pass


class ExcelBusy(ExcelError):
    pass


def user_message(exc: BaseException) -> tuple[str, str]:
    """Vrací (titulek_dialogu, srozumitelná_zpráva) v češtině."""
    if isinstance(exc, ExcelNotRunning):
        return (
            "Excel není spuštěn",
            "Nebyla nalezena spuštěná instance aplikace Microsoft Excel.\n\n"
            "Spusťte Excel a otevřete sešit, poté stiskněte Obnovit.",
        )
    if isinstance(exc, NoActiveWorkbook):
        return (
            "Žádný aktivní sešit",
            "Excel běží, ale není otevřen žádný aktivní sešit.\n\n"
            "Otevřete sešit v Excelu a stiskněte Obnovit.",
        )
    if isinstance(exc, WorkbookProtected):
        return (
            "Sešit je chráněn",
            "Sešit má chráněnou strukturu, proto Excel operaci odmítl.\n\n"
            "Zrušte ochranu v Excelu (Revize > Zrušit ochranu sešitu) a akci opakujte.",
        )
    if isinstance(exc, CannotHideAllSheets):
        return (
            "List nelze skrýt",
            "Nelze skrýt poslední viditelný list. Sešit musí mít alespoň jeden viditelný list.",
        )
    if isinstance(exc, SheetNotFound):
        name = getattr(exc, "sheet_name", "")
        return (
            "List nebyl nalezen",
            f"List „{name}“ již v sešitě neexistuje.\n\nStiskněte Obnovit pro načtení aktuálního seznamu.",
        )
    if isinstance(exc, ValidationError):
        return ("Neplatný název", str(exc) or "Zadaný název není platný.")
    if isinstance(exc, ExcelBusy):
        return (
            "Excel nereaguje",
            "Excel právě nereaguje — pravděpodobně upravujete buňku nebo je otevřen dialog v Excelu.\n\n"
            "Dokončete úpravu (Enter/Esc) nebo zavřete dialog v Excelu a akci opakujte.",
        )
    if isinstance(exc, SheetOperationDenied):
        detail = str(exc).strip()
        base = (
            "Excel operaci odmítl. Možná je sešit chráněn, list má nepovolený typ, "
            "nebo je aktivní zvláštní zobrazení."
        )
        return ("Operace se nezdařila", f"{base}" + (f"\n\nDetail: {detail}" if detail else ""))
    if isinstance(exc, ExcelError):
        return ("Chyba Excelu", str(exc) or "Nastala chyba při komunikaci s Excelem.")
    return ("Neočekávaná chyba", f"Nastala neočekávaná chyba: {exc}")
