"""Hlavní okno: pouze standardní widgety Qt, plná obsluha klávesnicí.

- Seznam listů: QListWidget, jedna položka = "název, typ, viditelnost, aktivní".
- Žádná informace není nesena pouze barvou.
- Fokus: start na seznamu; po dialozích návrat na volající prvek;
  po operacích výběr přesunutého/vytvořeného/sousedního listu.
"""

from __future__ import annotations

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.models import SheetInfo, Visibility, WorkbookInfo
from ui.dialogs import (
    AddSheetDialog,
    MoveSheetDialog,
    RenameDialog,
    confirm_delete,
    show_error,
)


class MainWindow(QMainWindow):
    # Signály ven do ExcelWorker (přes frontu Qt, neblokují GUI):
    requestRefresh = pyqtSignal()
    requestActivate = pyqtSignal(str)
    requestAdd = pyqtSignal(object)
    requestRename = pyqtSignal(str, str)
    requestDelete = pyqtSignal(str)
    requestMoveLeft = pyqtSignal(str)
    requestMoveRight = pyqtSignal(str)
    requestMoveTo = pyqtSignal(str, int)
    requestSetVisible = pyqtSignal(str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ovladač listů Excelu")
        self._sheets: list[SheetInfo] = []
        self._last_focus_before_dialog: QWidget | None = None

        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        self.workbookLabel = QLabel("Není připojeno k Excelu.", self)
        self.workbookLabel.setAccessibleName("Aktuální sešit")
        self.workbookLabel.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByKeyboard)
        layout.addWidget(self.workbookLabel)

        self.statusLabel = QLabel("", self)
        self.statusLabel.setAccessibleName("Stav")
        layout.addWidget(self.statusLabel)

        self.sheetList = QListWidget(self)
        self.sheetList.setAccessibleName("Seznam listů")
        self.sheetList.setAccessibleDescription("Šipkami vyberte list. Enter přejde na vybraný list.")
        self.sheetList.itemActivated.connect(self._onItemActivated)
        layout.addWidget(self.sheetList)

        btn_row1 = QHBoxLayout()
        self.btnRefresh = QPushButton("&Obnovit (F5)", self)
        self.btnRefresh.setAccessibleName("Obnovit seznam listů")
        self.btnRefresh.clicked.connect(lambda: self.requestRefresh.emit())
        btn_row1.addWidget(self.btnRefresh)

        self.btnGo = QPushButton("&Přejít na list", self)
        self.btnGo.setAccessibleName("Přejít na vybraný list")
        self.btnGo.clicked.connect(self._onGo)
        btn_row1.addWidget(self.btnGo)

        self.btnAdd = QPushButton("&Nový list", self)
        self.btnAdd.setAccessibleName("Vytvořit nový list")
        self.btnAdd.clicked.connect(self._onAdd)
        btn_row1.addWidget(self.btnAdd)

        self.btnRename = QPushButton("Pře&jmenovat", self)
        self.btnRename.setAccessibleName("Přejmenovat vybraný list")
        self.btnRename.clicked.connect(self._onRename)
        btn_row1.addWidget(self.btnRename)
        layout.addLayout(btn_row1)

        btn_row2 = QHBoxLayout()
        self.btnDelete = QPushButton("O&dstranit", self)
        self.btnDelete.setAccessibleName("Odstranit vybraný list")
        self.btnDelete.clicked.connect(self._onDelete)
        btn_row2.addWidget(self.btnDelete)

        self.btnLeft = QPushButton("Přesunout &vlevo", self)
        self.btnLeft.setAccessibleName("Přesunout list doleva")
        self.btnLeft.clicked.connect(self._onMoveLeft)
        btn_row2.addWidget(self.btnLeft)

        self.btnRight = QPushButton("Přesunout v&pravo", self)
        self.btnRight.setAccessibleName("Přesunout list doprava")
        self.btnRight.clicked.connect(self._onMoveRight)
        btn_row2.addWidget(self.btnRight)

        self.btnMoveTo = QPushButton("Přesunout &na pozici…", self)
        self.btnMoveTo.setAccessibleName("Přesunout list na konkrétní pozici")
        self.btnMoveTo.clicked.connect(self._onMoveTo)
        btn_row2.addWidget(self.btnMoveTo)
        layout.addLayout(btn_row2)

        btn_row3 = QHBoxLayout()
        self.btnHide = QPushButton("&Skrýt list", self)
        self.btnHide.setAccessibleName("Skrýt vybraný list")
        self.btnHide.clicked.connect(self._onHide)
        btn_row3.addWidget(self.btnHide)

        self.btnShow = QPushButton("&Zobrazit list", self)
        self.btnShow.setAccessibleName("Zobrazit vybraný skrytý list")
        self.btnShow.clicked.connect(self._onShow)
        btn_row3.addWidget(self.btnShow)
        layout.addLayout(btn_row3)

        # Pořadí Tab:
        QWidget.setTabOrder(self.sheetList, self.btnRefresh)
        QWidget.setTabOrder(self.btnRefresh, self.btnGo)
        QWidget.setTabOrder(self.btnGo, self.btnAdd)
        QWidget.setTabOrder(self.btnAdd, self.btnRename)
        QWidget.setTabOrder(self.btnRename, self.btnDelete)
        QWidget.setTabOrder(self.btnDelete, self.btnLeft)
        QWidget.setTabOrder(self.btnLeft, self.btnRight)
        QWidget.setTabOrder(self.btnRight, self.btnMoveTo)
        QWidget.setTabOrder(self.btnMoveTo, self.btnHide)
        QWidget.setTabOrder(self.btnHide, self.btnShow)

        # F5 jako standardní Obnovit:
        act_refresh = QAction(self)
        act_refresh.setShortcut(QKeySequence.StandardKey.Refresh)
        act_refresh.triggered.connect(lambda: self.requestRefresh.emit())
        self.addAction(act_refresh)

        self.sheetList.setFocus()
        self._setBusy(False)

    # -- veřejné sloty z workeru -----------------------------------------
    def onSheetsReady(self, wb: WorkbookInfo, sheets: list[SheetInfo], select_name: str | None = None) -> None:
        self._sheets = list(sheets)
        if wb is not None:
            prot = " (chráněná struktura)" if wb.protect_structure else ""
            self.workbookLabel.setText(f"Aktuální sešit: {wb.name}{prot}")
        self.sheetList.clear()
        select_row = 0
        for i, s in enumerate(self._sheets):
            item = QListWidgetItem(s.display_text())
            item.setData(Qt.ItemDataRole.UserRole, s.name)
            self.sheetList.addItem(item)
            if select_name is not None and s.name.lower() == select_name.lower():
                select_row = i
            elif select_name is None and s.is_active:
                select_row = i
        if self._sheets:
            self.sheetList.setCurrentRow(select_row)
        self._updateButtons()
        self.statusLabel.setText(f"Načteno {len(self._sheets)} listů.")

    def onOperationFailed(self, title: str, message: str) -> None:
        show_error(self, title, message)
        self.statusLabel.setText(title)
        # Fokus zůstává kde byl; dialog ho po zavření vrátí.

    def onBusy(self, busy: bool) -> None:
        self._setBusy(busy)

    # -- fokus / výběr ----------------------------------------------------
    def selectedSheetName(self) -> str | None:
        item = self.sheetList.currentItem()
        if item is None:
            return None
        data = item.data(Qt.ItemDataRole.UserRole)
        return str(data) if data else item.text().split(",")[0].strip()

    def selectedSheet(self) -> SheetInfo | None:
        name = self.selectedSheetName()
        if not name:
            return None
        for s in self._sheets:
            if s.name.lower() == name.lower():
                return s
        return None

    def selectByName(self, name: str) -> None:
        for i in range(self.sheetList.count()):
            item = self.sheetList.item(i)
            if str(item.data(Qt.ItemDataRole.UserRole)).lower() == name.lower():
                self.sheetList.setCurrentRow(i)
                break
        self.sheetList.setFocus()

    def _setBusy(self, busy: bool) -> None:
        for b in (self.btnRefresh, self.btnGo, self.btnAdd, self.btnRename, self.btnDelete,
                  self.btnLeft, self.btnRight, self.btnMoveTo, self.btnHide, self.btnShow):
            b.setEnabled(not busy)
        if busy:
            self.statusLabel.setText("Pracuji…")
        self._updateButtons()
        if busy:
            for b in (self.btnRefresh, self.btnGo, self.btnAdd, self.btnRename, self.btnDelete,
                      self.btnLeft, self.btnRight, self.btnMoveTo, self.btnHide, self.btnShow):
                b.setEnabled(False)

    def _updateButtons(self) -> None:
        has = bool(self._sheets)
        sel = self.selectedSheet()
        self.btnGo.setEnabled(has and sel is not None)
        self.btnRename.setEnabled(has and sel is not None)
        self.btnDelete.setEnabled(has and sel is not None)
        self.btnLeft.setEnabled(has and sel is not None and sel.index > 1)
        self.btnRight.setEnabled(has and sel is not None and sel.index < len(self._sheets))
        self.btnMoveTo.setEnabled(has and sel is not None)
        if sel is None:
            self.btnHide.setEnabled(False)
            self.btnShow.setEnabled(False)
        else:
            self.btnHide.setEnabled(sel.visibility == Visibility.VISIBLE)
            self.btnShow.setEnabled(sel.visibility != Visibility.VISIBLE)

    # -- handlery ----------------------------------------------------------
    def _onItemActivated(self, item: QListWidgetItem) -> None:
        self._onGo()

    def _requireSelection(self) -> str | None:
        name = self.selectedSheetName()
        if not name:
            show_error(self, "Není vybrán list", "Nejprve šipkami vyberte list v seznamu.")
            self.sheetList.setFocus()
        return name

    def _onGo(self) -> None:
        name = self._requireSelection()
        if name:
            self.statusLabel.setText(f"Přecházím na list „{name}“…")
            self.requestActivate.emit(name)

    def _onAdd(self) -> None:
        self._last_focus_before_dialog = self.btnAdd
        dlg = AddSheetDialog(self)
        dlg.edit.setFocus()
        if dlg.exec() == dlg.DialogCode.Accepted:
            self.requestAdd.emit(dlg.value().strip() or None)
        self.btnAdd.setFocus()

    def _onRename(self) -> None:
        name = self._requireSelection()
        if not name:
            return
        self._last_focus_before_dialog = self.btnRename
        dlg = RenameDialog(self, name)
        dlg.edit.setFocus()
        if dlg.exec() == dlg.DialogCode.Accepted:
            self.requestRename.emit(name, dlg.value())
        self.sheetList.setFocus()

    def _onDelete(self) -> None:
        name = self._requireSelection()
        if not name:
            return
        # Výchozí volba v dialogu je Ne (viz dialogs.confirm_delete).
        if confirm_delete(self, name):
            self.requestDelete.emit(name)
        self.sheetList.setFocus()

    def _onMoveLeft(self) -> None:
        name = self._requireSelection()
        if name:
            self.requestMoveLeft.emit(name)

    def _onMoveRight(self) -> None:
        name = self._requireSelection()
        if name:
            self.requestMoveRight.emit(name)

    def _onMoveTo(self) -> None:
        sel = self.selectedSheet()
        if sel is None:
            self._requireSelection()
            return
        dlg = MoveSheetDialog(self, sel.name, len(self._sheets), sel.index)
        dlg.spin.setFocus()
        if dlg.exec() == dlg.DialogCode.Accepted:
            self.requestMoveTo.emit(sel.name, int(dlg.value()))
        self.sheetList.setFocus()

    def _onHide(self) -> None:
        name = self._requireSelection()
        if name:
            self.requestSetVisible.emit(name, False)

    def _onShow(self) -> None:
        name = self._requireSelection()
        if name:
            self.requestSetVisible.emit(name, True)
