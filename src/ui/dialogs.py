"""Přístupné dialogy: skutečné titulky, QLabel+buddy, správné defaulty.

- Potvrzení odstranění: QMessageBox s defaultem Ne.
- Chybové dialogy: specifické titulky z core.errors.user_message, nikdy obecné.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)


class _LabeledLineDialog(QDialog):
    def __init__(self, parent: QWidget | None, title: str, label_text: str, initial: str = ""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        layout = QVBoxLayout(self)
        self.label = QLabel(label_text, self)
        self.edit = QLineEdit(initial, self)
        self.edit.setAccessibleName(label_text.rstrip(":"))
        self.label.setBuddy(self.edit)
        layout.addWidget(self.label)
        layout.addWidget(self.edit)
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Potvrdit")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Zrušit")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.edit.selectAll()

    def value(self) -> str:
        return self.edit.text()


class RenameDialog(_LabeledLineDialog):
    def __init__(self, parent: QWidget | None, old_name: str):
        super().__init__(parent, f"Přejmenovat list „{old_name}“", "Nový název listu:", old_name)


class AddSheetDialog(_LabeledLineDialog):
    def __init__(self, parent: QWidget | None):
        super().__init__(
            parent,
            "Vytvořit nový list",
            "Název nového listu (prázdné = výchozí název Excelu):",
            "",
        )


class MoveSheetDialog(QDialog):
    """Přesun na konkrétní pozici 1..N."""

    def __init__(self, parent: QWidget | None, sheet_name: str, count: int, current: int):
        super().__init__(parent)
        self.setWindowTitle(f"Přesunout list „{sheet_name}“")
        self.setModal(True)
        layout = QVBoxLayout(self)
        self.label = QLabel(f"Nová pozice (1–{count}):", self)
        self.spin = QSpinBox(self)
        self.spin.setMinimum(1)
        self.spin.setMaximum(max(1, count))
        self.spin.setValue(current)
        self.spin.setAccessibleName("Nová pozice listu")
        self.label.setBuddy(self.spin)
        layout.addWidget(self.label)
        layout.addWidget(self.spin)
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Přesunout")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Zrušit")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def value(self) -> int:
        return int(self.spin.value())


def confirm_delete(parent: QWidget | None, sheet_name: str) -> bool:
    """Potvrzovací dialog; defaultní volba je vždy Ne. Vrací True jen při Ano."""
    box = QMessageBox(parent)
    box.setWindowTitle(f"Odstranit list „{sheet_name}“?")
    box.setText(f"Opravdu chcete odstranit list „{sheet_name}“?")
    box.setInformativeText("Tuto akci nelze vrátit zpět. Pokračovat?")
    box.setIcon(QMessageBox.Icon.Warning)
    yes = box.addButton("Ano, odstranit", QMessageBox.ButtonRole.YesRole)
    no = box.addButton("Ne", QMessageBox.ButtonRole.NoRole)
    box.setDefaultButton(no)  # povinné: výchozí Ne
    box.setEscapeButton(no)
    box.exec()
    return box.clickedButton() == yes


def show_error(parent: QWidget | None, title: str, message: str) -> None:
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(message)
    box.setIcon(QMessageBox.Icon.Critical)
    box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
    box.exec()
