# rotationsplan.py

import pandas as pd
import textwrap
import re

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor

# modul_rotation.lade_daten() sollte nun auch sheet_name annehmen!
from modul_rotation import lade_daten  

class Rotationsplan(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        # Layout & Titel
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        lbl = QLabel("Rotationsplan")
        lbl.setFont(QFont("Helvetica", 24, QFont.Weight.Bold))
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)

        # Buttons
        btn_layout = QHBoxLayout()
        layout.addLayout(btn_layout)
        btn_z = QPushButton("Zurück")
        btn_z.clicked.connect(self.main_window.zeige_startseite)
        btn_layout.addWidget(btn_z)
        btn_layout.addStretch()
        btn_b = QPushButton("Beenden")
        btn_b.clicked.connect(self.main_window.close)
        btn_layout.addWidget(btn_b)

        # Tabelle
        self.table = QTableWidget()
        layout.addWidget(self.table, stretch=1)
        # fette Linie unter Headern
        self.table.setStyleSheet("QHeaderView::section { border-bottom: 3px solid #666; }")

        self.lade_und_zeige_rotationsplan()

    def lade_und_zeige_rotationsplan(self):
        # wir erwarten jetzt, dass lade_daten sheet_name akzeptiert
        try:
            df = lade_daten(sheet_name="Rotationsplan")
        except TypeError:
            # fallback, falls lade_daten noch keine Übergabe erlaubt:
            from pandas import read_excel
            from modul_rotation import EXCEL_PFAD
            df = read_excel(EXCEL_PFAD, sheet_name="Rotationsplan")
        if df is None:
            QMessageBox.critical(self, "Fehler", "Excel-Datei konnte nicht geladen werden.")
            return

        # Spaltenliste
        cols = list(df.columns)

        # 1) Wenn es schon eine Spalte "Name" gibt, nehmen wir die
        if "Mitarbeiter" in df.columns:
            name_col = "Mitarbeiter"
        # 2) Sonst, falls "Vorname" + "Nachname" vorhanden, kombinieren wir
        elif {"Vorname", "Nachname"}.issubset(df.columns):
            df["Mitarbeiter"] = df["Vorname"].fillna("") + " " + df["Nachname"].fillna("")
            name_col = "Mitarbeiter"
        else:
            # letzter Ausweg: erste Spalte
            name_col = cols[0]

        # Welche Spalten wir anzeigen wollen:
        keep = [name_col] + [
            c for c in cols
            if c not in ("Vorname", "Nachname")
               and "lfd" not in str(c).lower()
               and not str(c).startswith("Unnamed")
        ]
        # letzte Spalte (falls Leer) entfernen
        if len(keep) > 2 and df[keep[-1]].isna().all():
            keep.pop()

        daten = df[keep]

        # Tabelle vorbereiten
        self.table.setColumnCount(len(keep))
        self.table.setRowCount(len(daten))

        # Header umbrechen bei max. 10 Zeichen
        wrapped = [
            "\n".join(textwrap.wrap(str(c), width=10, break_long_words=True))
            for c in keep
        ]
        self.table.setHorizontalHeaderLabels(wrapped)

        # Header stylen
        hdr = self.table.horizontalHeader()
        hf = QFont("Helvetica", 10, QFont.Weight.Bold)
        hdr.setFont(hf)
        hdr.setFixedHeight(80)
        for i in range(len(keep)):
            self.table.horizontalHeaderItem(i).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        # RegEx für Datumserkennung
        date_re = re.compile(r"\d{2}\.\d{2}\.\d{4}")

        # Daten einfüllen
        for r, row in daten.iterrows():
            max_h = 0
            for c, col in enumerate(keep):
                val = row[col]
                txt = "" if pd.isna(val) else str(val).strip()

                # Name-Spalte: fetter, blau, word-wrap
                if col == name_col:
                    w = QLabel(txt)
                    w.setWordWrap(True)
                    w.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    w.setStyleSheet("color:#003366; font-weight:bold;")
                    lines = textwrap.wrap(txt, width=15)
                    h = max(30, len(lines)*20)
                    self.table.setRowHeight(r, h)
                    self.table.setCellWidget(r, c, w)

                # X-Zellen: rote X
                elif txt.lower() == "x":
                    itm = QTableWidgetItem("X")
                    itm.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    itm.setForeground(QColor("red"))
                    itm.setFont(QFont("Helvetica", 11, QFont.Weight.Bold))
                    self.table.setItem(r, c, itm)

                # Datumszellen: hellgrau + erste Zeile grün, zweite rot
                elif date_re.search(txt):
                    parts = txt.split("\n")
                    html = ""
                    if parts:
                        html += f"<span style='color:darkgreen'>{parts[0]}</span>"
                    if len(parts) > 1:
                        html += "<br>" + f"<span style='color:red'>{parts[1]}</span>"
                    lbl = QLabel(html)
                    lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    lbl.setTextFormat(Qt.TextFormat.RichText)
                    lbl.setStyleSheet("background-color:#f0f0f0;")
                    self.table.setCellWidget(r, c, lbl)
                    h = max(30, len(parts)*20)
                    max_h = max(max_h, h)

                # Alle anderen Zellen:
                else:
                    itm = QTableWidgetItem(txt)
                    itm.setTextAlignment(Qt.AlignmentFlag.AlignCenter|Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(r, c, itm)
                    h = max(30, (txt.count("\n")+1)*20)
                    max_h = max(max_h, h)

            # Zeilenhöhe anwenden
            if max_h > self.table.rowHeight(r):
                self.table.setRowHeight(r, max_h)

        # Spalten strecken, Scrollbars:
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setSelectionBehavior(self.table.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(self.table.SelectionMode.SingleSelection)
