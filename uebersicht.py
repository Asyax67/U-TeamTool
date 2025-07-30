# uebersicht.py

import pandas as pd
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox,
    QTableWidget, QTableWidgetItem, QLabel
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from modul_rotation import lade_daten, finde_aktuellen_bereich

class Uebersicht(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        # Titel
        lbl = QLabel("Mitarbeiterübersicht")
        lbl.setFont(QFont("Helvetica", 20, QFont.Weight.Bold))
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)

        # Buttons oben
        btn_layout = QHBoxLayout()
        layout.addLayout(btn_layout)

        btn_back = QPushButton("Zurück")
        btn_back.clicked.connect(self.main_window.zeige_startseite)
        btn_layout.addWidget(btn_back)
        btn_layout.addStretch()
        btn_quit = QPushButton("Beenden")
        btn_quit.clicked.connect(self.main_window.close)
        btn_layout.addWidget(btn_quit)

        # Tabelle
        self.table = QTableWidget()
        layout.addWidget(self.table, stretch=1)

        self.lade_und_zeige_daten()

    def lade_und_zeige_daten(self):
        df = lade_daten("Masterlist")
        if df is None:
            QMessageBox.critical(self, "Fehler", "Masterlist konnte nicht geladen werden.")
            return

        # Formatieren
        df["Aktuelles Austrittsdatum"] = (
            pd.to_datetime(df["Aktuelles Austrittsdatum"], errors="coerce")
            .dt.strftime("%d.%m.%Y")
        )
        df["Aktueller Bereich"] = df.apply(finde_aktuellen_bereich, axis=1)

        cols = ["Vorname", "Nachname", "Aktueller Bereich", "Aktuelles Austrittsdatum"]
        daten = df[cols]

        # Setup TableWidget
        self.table.setColumnCount(len(cols))
        self.table.setRowCount(len(daten))
        self.table.setHorizontalHeaderLabels(cols)

        # Header Styling
        font = QFont()
        font.setBold(True)
        self.table.horizontalHeader().setFont(font)
        for i in range(len(cols)):
            self.table.horizontalHeaderItem(i).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        # Daten füllen
        for row_idx, (_, row) in enumerate(daten.iterrows()):
            for col_idx, col in enumerate(cols):
                item = QTableWidgetItem(str(row[col]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)

        # Spalten automatisch breit genug
        self.table.resizeColumnsToContents()
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(self.table.SelectionMode.SingleSelection)
