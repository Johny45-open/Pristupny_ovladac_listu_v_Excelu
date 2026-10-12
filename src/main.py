"""Vstupní bod aplikace."""

from __future__ import annotations

import sys

sys.path.insert(0, __import__("os").path.join(__import__("os").path.dirname(__file__)))

from PyQt6.QtCore import QThread
from PyQt6.QtWidgets import QApplication

from excel_backend.com_adapter import ComExcelAdapter
from ui.main_window import MainWindow
from workers.excel_worker import ExcelWorker


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Ovladač listů Excelu")
    app.setOrganizationName("PristupnyOvladac")

    window = MainWindow()

    thread = QThread(window)
    worker = ExcelWorker(lambda: ComExcelAdapter())
    worker.moveToThread(thread)
    thread.started.connect(worker.onThreadStarted)
    thread.finished.connect(worker.onThreadFinished)

    # GUI -> worker (queued, neblokuje):
    window.requestRefresh.connect(worker.refresh)
    window.requestActivate.connect(worker.activateSheet)
    window.requestAdd.connect(worker.addSheet)
    window.requestRename.connect(worker.renameSheet)
    window.requestDelete.connect(worker.deleteSheet)
    window.requestMoveLeft.connect(worker.moveLeft)
    window.requestMoveRight.connect(worker.moveRight)
    window.requestMoveTo.connect(worker.moveSheet)
    window.requestSetVisible.connect(worker.setVisible)

    # worker -> GUI:
    worker.sheetsReady.connect(lambda wb, sheets: window.onSheetsReady(wb, sheets))
    worker.operationFailed.connect(window.onOperationFailed)
    worker.busyChanged.connect(window.onBusy)
    # Po jednorázových operacích držíme fokus na seznamu a vybíráme dotčený list:
    worker.sheetCreated.connect(lambda info: window.selectByName(info.name))
    worker.sheetRenamed.connect(lambda info: window.selectByName(info.name))
    worker.sheetMoved.connect(lambda info: window.selectByName(info.name))
    worker.visibilityChanged.connect(lambda info: window.selectByName(info.name))
    worker.activateFinished.connect(lambda name: window.selectByName(name))

    thread.start()
    window.show()
    window.sheetList.setFocus()
    window.requestRefresh.emit()

    code = app.exec()
    thread.quit()
    thread.wait(3000)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
