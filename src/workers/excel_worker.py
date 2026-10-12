"""Qt worker pro komunikaci s Excelem bez blokování GUI.

- Žije v dedikovaném QThread.
- Při startu vlákna volá pythoncom.CoInitialize(), při konci CoUninitialize().
- COM adaptér vytváří až uvnitř worker vlákna (adapter_factory).
- GUI volá sloty přes signály; výsledky se vrací signály.
- Živé COM objekty nikdy neopouští toto vlákno — ven jdou pouze dataclasses.
"""

from __future__ import annotations

from typing import Callable

from PyQt6.QtCore import QObject, pyqtSignal, pyqtSlot

from controller.sheet_controller import SheetController
from core.errors import user_message
from core.models import SheetInfo, WorkbookInfo
from excel_backend.base import ExcelAdapter


class ExcelWorker(QObject):
    sheetsReady = pyqtSignal(object, object)  # (WorkbookInfo, list[SheetInfo])
    activateFinished = pyqtSignal(str)  # name
    sheetCreated = pyqtSignal(object)  # SheetInfo
    sheetRenamed = pyqtSignal(object)  # SheetInfo
    sheetDeleted = pyqtSignal(str)  # deleted name
    sheetMoved = pyqtSignal(object)  # SheetInfo
    visibilityChanged = pyqtSignal(object)  # SheetInfo
    operationFailed = pyqtSignal(str, str)  # (titulek, zpráva)
    busyChanged = pyqtSignal(bool)

    def __init__(self, adapter_factory: Callable[[], ExcelAdapter], parent=None):
        super().__init__(parent)
        self._factory = adapter_factory
        self._controller: SheetController | None = None
        self._com_initialized = False

    # -- životní cyklus vlákna ------------------------------------------
    @pyqtSlot()
    def onThreadStarted(self) -> None:
        try:
            import pythoncom  # type: ignore

            pythoncom.CoInitialize()
            self._com_initialized = True
        except ImportError:
            # Ne-Windows / testy bez pywin32: pokračuj bez COM init.
            self._com_initialized = False
        try:
            adapter = self._factory()
        except Exception as e:  # pragma: no cover - fatální
            title, msg = user_message(e)
            self.operationFailed.emit(title, msg)
            return
        self._controller = SheetController(adapter)

    @pyqtSlot()
    def onThreadFinished(self) -> None:
        self._controller = None
        if self._com_initialized:
            try:
                import pythoncom  # type: ignore

                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False

    def _guard(self) -> SheetController | None:
        if self._controller is None:
            self.operationFailed.emit("Operace se nezdařila", "Pracovní vlákno ještě není připraveno.")
            return None
        return self._controller

    def _run(self, fn):
        self.busyChanged.emit(True)
        try:
            fn()
        except Exception as e:
            title, msg = user_message(e)
            self.operationFailed.emit(title, msg)
        finally:
            self.busyChanged.emit(False)

    # -- sloty volané z GUI ----------------------------------------------
    @pyqtSlot()
    def refresh(self) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            wb, sheets = c.connect() if c.workbook is None else c.refresh()
            self.sheetsReady.emit(wb, sheets)

        self._run(_do)

    @pyqtSlot(str)
    def activateSheet(self, name: str) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            c.activate(name)
            self.activateFinished.emit(name)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(object)
    def addSheet(self, name) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            created = c.add(name if isinstance(name, str) and name.strip() else None)
            self.sheetCreated.emit(created)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str, str)
    def renameSheet(self, old_name: str, new_name: str) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            renamed = c.rename(old_name, new_name)
            self.sheetRenamed.emit(renamed)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str)
    def deleteSheet(self, name: str) -> None:
        """GUI musí potvrzení řešit PŘED voláním (dialog s defaultem Ne)."""
        c = self._guard()
        if c is None:
            return

        def _do():
            c.delete(name, confirmed=True)
            self.sheetDeleted.emit(name)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str, int)
    def moveSheet(self, name: str, new_index: int) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            moved = c.move(name, int(new_index))
            self.sheetMoved.emit(moved)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str)
    def moveLeft(self, name: str) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            moved = c.move_left(name)
            self.sheetMoved.emit(moved)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str)
    def moveRight(self, name: str) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            moved = c.move_right(name)
            self.sheetMoved.emit(moved)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    @pyqtSlot(str, bool)
    def setVisible(self, name: str, visible: bool) -> None:
        c = self._guard()
        if c is None:
            return

        def _do():
            updated = c.set_visible(name, bool(visible))
            self.visibilityChanged.emit(updated)
            if c.workbook is not None:
                self.sheetsReady.emit(c.workbook, c.sheets)

        self._run(_do)

    # -- synchronní pomocné (pro testy bez Qt event-loopu) ----------------
    def controllerForTests(self) -> SheetController | None:
        return self._controller

    def ensureControllerForTests(self) -> SheetController:
        if self._controller is None:
            self._controller = SheetController(self._factory())
        return self._controller
