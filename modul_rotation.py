# modul_rotation.py

import pandas as pd
from datetime import datetime

# Pfad zur Masterliste (mit forward‐slashes!)
from pathlib import Path
import sys

# Basisordner bestimmen (funktioniert im Script UND in der EXE)
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

EXCEL_PATH = BASE_DIR / "Masterliste_UTeam.xlsx"
# falls du noch andere Dateien brauchst:
# WORD_VORLAGEN_DIR = BASE_DIR / "vorlagen"



def lade_daten(sheet_name: str = "Masterlist") -> pd.DataFrame | None:
    """
    Lädt das Sheet `sheet_name` als DataFrame.
    """
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"Fehler beim Laden der Excel-Datei ({sheet_name}): {e}")
        return None

def finde_aktuellen_bereich(row: pd.Series) -> str:
    """
    Liest 'Einsatz Station 1' bis 'Einsatz Station 8' und gibt die erste nicht-leere
    Station zurück.
    """
    for i in range(1, 9):
        key = f"Einsatz Station {i}"
        wert = row.get(key, "")
        if isinstance(wert, str) and wert.strip():
            return wert.strip()
    return "Nicht eingesetzt"
