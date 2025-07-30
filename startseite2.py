import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QStackedWidget
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from uebersicht import Uebersicht
from rotationsplan import Rotationsplan
from Vorstellung import Vorstellung
from bereich_anlegen import BereichAnlegen

class Startseite(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        lbl = QLabel("UTeam Rotations‑Tool")
        lbl.setFont(QFont("Helvetica", 24, QFont.Weight.Bold))
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)
        layout.addSpacing(30)

        for txt, slot in [
            ("Mitarbeiterübersicht",  self.main_window.zeige_uebersicht),
            ("Rotationsplan",         self.main_window.zeige_rotationsplan),
            ("Vorstellungsgespräch",  self.main_window.zeige_vorstellung),
            ("Neuen Bereich anlegen", self.main_window.zeige_bereich_anlegen),
        ]:
            b = QPushButton(txt)
            b.clicked.connect(slot)
            layout.addWidget(b)

        layout.addStretch()
        be = QPushButton("Beenden")
        be.clicked.connect(self.main_window.close)
        layout.addWidget(be)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("UTeam Rotationsprogramm")
        self.resize(1000, 700)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # Alle Seiten instanziieren
        self.start      = Startseite(self)
        self.uebersicht = Uebersicht(self)
        self.rplan      = Rotationsplan(self)
        self.vorstellung= Vorstellung(self)
        self.bereich    = BereichAnlegen(self)

        for w in (self.start, self.uebersicht, self.rplan, self.vorstellung, self.bereich):
            self.stack.addWidget(w)

        self.zeige_startseite()

    def zeige_startseite(self):       self.stack.setCurrentWidget(self.start)
    def zeige_uebersicht(self):      self.stack.setCurrentWidget(self.uebersicht)
    def zeige_rotationsplan(self):   self.stack.setCurrentWidget(self.rplan)
    def zeige_vorstellung(self):     self.stack.setCurrentWidget(self.vorstellung)
    def zeige_bereich_anlegen(self): self.stack.setCurrentWidget(self.bereich)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    app.setStyleSheet("""
        QWidget { background-color: #f9f9f9; color: #222; font-family: Arial, sans-serif; font-size: 11pt; }
        QPushButton { background-color: #007ACC; color: white; border-radius: 5px; padding: 6px 12px; }
        QPushButton:hover { background-color: #005F99; }
        QTableWidget { gridline-color: #DDD; background-color: white; }
        QHeaderView::section { background-color: white; color: #003366; font-weight: bold; border: 1px solid #CCC; }
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
