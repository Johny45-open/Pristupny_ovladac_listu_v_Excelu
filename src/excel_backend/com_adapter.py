"""COM adaptér pro desktopový Microsoft Excel (Windows, pywin32).

Pravidla:
- Připojuje se VÝHRADNĚ k běžící instanci přes GetActiveObject.
  Nikdy nevytváří novou instanci (Dispatch) bez vědomí uživatele.
- Živé COM objekty nevynáší ven — vrací pouze dataclasses z core.models.
- Všechny pywintypes.com_error překládá na doménové výjimky z core.errors.
- Objekt adaptéru smí být používán pouze ve vlákně, které ho vytvořilo
  (viz workers/excel_worker.py — CoInitialize/CoUninitialize).
"""

from __future__ import annotations

from core.errors import (
    ExcelBusy,
    ExcelNotRunning,
    NoActiveWorkbook,
    SheetNotFound,
    SheetOperationDenied,
    WorkbookProtected,
)
from core.models import SheetInfo, SheetType, Visibility, WorkbookInfo
from excel_backend.base import ExcelAdapter

# Excel konstanty (nezávislé na instalovaném Excelu, aby modul šel importovat i bez něj):
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2

# XlSheetType:
XL_WORKSHEET = -4167
XL_CHART = -4109
XL_DIALOG_SHEET = -4116
XL_MACRO_SHEET = 3  # xlExcel4MacroSheet
XL_INTL_MACRO = 4  # xlExcel4IntlMacroSheet


def _map_type(xl_type: int) -> SheetType:
    if xl_type == XL_WORKSHEET:
        return SheetType.WORKSHEET
    if xl_type == XL_CHART:
        return SheetType.CHART
    if xl_type == XL_DIALOG_SHEET:
        return SheetType.DIALOG
    if xl_type in (XL_MACRO_SHEET, XL_INTL_MACRO):
        return SheetType.MACRO
    return SheetType.UNKNOWN


def _map_visibility(xl_visible) -> Visibility:
    try:
        v = int(xl_visible)
    except (TypeError, ValueError):
        return Visibility.VISIBLE
    if v == XL_SHEET_VISIBLE:
        return Visibility.VISIBLE
    if v == XL_SHEET_VERY_HIDDEN:
        return Visibility.VERY_HIDDEN
    return Visibility.HIDDEN


class ComExcelAdapter(ExcelAdapter):
    def __init__(self) -> None:
        self._app = None

    # -- připojení -----------------------------------------------------
    def connect(self) -> WorkbookInfo:
        try:
            import win32com.client  # type: ignore
        except ImportError as e:  # pragma: no cover - bez pywin32 na Linuxu
            raise ExcelNotRunning("Knihovna pywin32 není k dispozici.") from e
        try:
            self._app = win32com.client.GetActiveObject("Excel.Application")
        except Exception as e:
            self._app = None
            raise ExcelNotRunning(str(e)) from e
        return self.get_workbook_info()

    def _workbook(self):
        if self._app is None:
            self.connect()
        assert self._app is not None
        try:
            import pywintypes  # type: ignore
        except ImportError:  # pragma: no cover
            pywintypes = None  # type: ignore
        try:
            wb = self._app.ActiveWorkbook
        except Exception as e:
            if pywintypes is not None and isinstance(e, pywintypes.com_error):
                raise ExcelBusy(str(e)) from e
            raise SheetOperationDenied(str(e)) from e
        if wb is None:
            raise NoActiveWorkbook()
        return wb

    def _wrap_com_error(self, e: Exception, context: str = "") -> Exception:
        msg = f"{context}: {e}".strip() if context else str(e)
        text = msg.lower()
        if "protect" in text or "chrán" in text or "0x800a03ec" in text and "protect" in text:
            # Hrubá detekce; přesnou ochranu řeší _ensure_not_protected().
            pass
        return SheetOperationDenied(msg)

    def _ensure_not_protected(self, wb) -> None:
        try:
            if bool(wb.ProtectStructure):
                raise WorkbookProtected()
        except WorkbookProtected:
            raise
        except Exception as e:
            raise SheetOperationDenied(str(e)) from e

    def _sheet_by_name(self, wb, name: str):
        try:
            return wb.Sheets(name)
        except Exception as e:
            raise SheetNotFound(name) from e

    # -- čtení ---------------------------------------------------------
    def get_workbook_info(self) -> WorkbookInfo:
        wb = self._workbook()
        try:
            name = str(wb.Name)
            try:
                full = str(wb.FullName)
            except Exception:
                full = name
            try:
                protected = bool(wb.ProtectStructure)
            except Exception:
                protected = False
            try:
                count = int(wb.Sheets.Count)
            except Exception:
                count = 0
            return WorkbookInfo(name=name, full_name=full, protect_structure=protected, sheet_count=count)
        except (ExcelNotRunning, NoActiveWorkbook):
            raise
        except Exception as e:
            raise SheetOperationDenied(str(e)) from e

    def list_sheets(self) -> list[SheetInfo]:
        wb = self._workbook()
        try:
            try:
                active_name = str(wb.ActiveSheet.Name)
            except Exception:
                active_name = ""
            result: list[SheetInfo] = []
            count = int(wb.Sheets.Count)
            for i in range(1, count + 1):
                sh = wb.Sheets(i)
                try:
                    name = str(sh.Name)
                except Exception:
                    continue
                try:
                    xl_type = int(sh.Type)
                except Exception:
                    xl_type = XL_WORKSHEET
                try:
                    xl_vis = int(sh.Visible)
                except Exception:
                    xl_vis = XL_SHEET_VISIBLE
                result.append(
                    SheetInfo(
                        name=name,
                        index=i,
                        sheet_type=_map_type(xl_type),
                        visibility=_map_visibility(xl_vis),
                        is_active=(name == active_name),
                    )
                )
            return result
        except (ExcelNotRunning, NoActiveWorkbook):
            raise
        except Exception as e:
            raise SheetOperationDenied(str(e)) from e

    # -- zápis ---------------------------------------------------------
    def activate_sheet(self, name: str) -> None:
        wb = self._workbook()
        sh = self._sheet_by_name(wb, name)
        try:
            # Skrytý list nelze aktivovat — Excel vyhodí chybu, přeložíme ji.
            sh.Activate()
        except Exception as e:
            raise SheetOperationDenied(f"List „{name}“ nelze aktivovat (možná je skrytý): {e}") from e

    def add_sheet(self, name: str | None = None) -> SheetInfo:
        wb = self._workbook()
        self._ensure_not_protected(wb)
        try:
            # Přidat na konec pro předvídatelnost.
            last = wb.Sheets(wb.Sheets.Count)
            sh = wb.Sheets.Add(After=last)
            if name:
                sh.Name = name
            return SheetInfo(
                name=str(sh.Name),
                index=int(wb.Sheets(str(sh.Name)).Index)
                if hasattr(wb.Sheets(str(sh.Name)), "Index")
                else int(wb.Sheets.Count),
                sheet_type=_map_type(int(sh.Type)),
                visibility=_map_visibility(int(sh.Visible)),
                is_active=True,
            )
        except WorkbookProtected:
            raise
        except Exception as e:
            raise self._wrap_com_error(e, "Nový list se nepodařilo vytvořit") from e

    def rename_sheet(self, old_name: str, new_name: str) -> SheetInfo:
        wb = self._workbook()
        self._ensure_not_protected(wb)
        sh = self._sheet_by_name(wb, old_name)
        try:
            sh.Name = new_name
            return SheetInfo(
                name=str(sh.Name),
                index=int(sh.Index),
                sheet_type=_map_type(int(sh.Type)),
                visibility=_map_visibility(int(sh.Visible)),
                is_active=False,
            )
        except Exception as e:
            raise self._wrap_com_error(e, f"Přejmenování listu „{old_name}“ se nezdařilo") from e

    def delete_sheet(self, name: str) -> None:
        wb = self._workbook()
        self._ensure_not_protected(wb)
        sh = self._sheet_by_name(wb, name)
        try:
            app = self._app
            prev_alerts = None
            try:
                prev_alerts = bool(app.DisplayAlerts)
                app.DisplayAlerts = False
            except Exception:
                prev_alerts = None
            try:
                sh.Delete()
            finally:
                if prev_alerts is not None:
                    try:
                        app.DisplayAlerts = prev_alerts
                    except Exception:
                        pass
        except Exception as e:
            raise self._wrap_com_error(e, f"Odstranění listu „{name}“ se nezdařilo") from e

    def move_sheet(self, name: str, new_index_1based: int) -> SheetInfo:
        wb = self._workbook()
        self._ensure_not_protected(wb)
        sh = self._sheet_by_name(wb, name)
        try:
            count = int(wb.Sheets.Count)
            target = max(1, min(int(new_index_1based), count))
            current = int(sh.Index)
            if target == current:
                return SheetInfo(
                    name=str(sh.Name),
                    index=current,
                    sheet_type=_map_type(int(sh.Type)),
                    visibility=_map_visibility(int(sh.Visible)),
                    is_active=False,
                )
            if target < current:
                sh.Move(Before=wb.Sheets(target))
            else:
                sh.Move(After=wb.Sheets(target))
            moved = wb.Sheets(name)
            return SheetInfo(
                name=str(moved.Name),
                index=int(moved.Index),
                sheet_type=_map_type(int(moved.Type)),
                visibility=_map_visibility(int(moved.Visible)),
                is_active=False,
            )
        except WorkbookProtected:
            raise
        except Exception as e:
            raise self._wrap_com_error(e, f"Přesunutí listu „{name}“ se nezdařilo") from e

    def set_sheet_visible(self, name: str, visible: bool) -> SheetInfo:
        wb = self._workbook()
        self._ensure_not_protected(wb)
        sh = self._sheet_by_name(wb, name)
        try:
            sh.Visible = XL_SHEET_VISIBLE if visible else XL_SHEET_HIDDEN
            return SheetInfo(
                name=str(sh.Name),
                index=int(sh.Index),
                sheet_type=_map_type(int(sh.Type)),
                visibility=_map_visibility(int(sh.Visible)),
                is_active=False,
            )
        except Exception as e:
            raise self._wrap_com_error(e, f"Změna viditelnosti listu „{name}“ se nezdařila") from e
