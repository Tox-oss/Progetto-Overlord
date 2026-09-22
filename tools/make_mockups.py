from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
FILES = ROOT / "files"
SOURCE = FILES / "source"
DEST = FILES / "dest"


def _scrivi_blocchi(ws, titolo: str, blocchi: list, riga_inizio: int = 1):
    ws.cell(row=riga_inizio, column=1, value=titolo).font = Font(bold=True, size=14)
    riga = riga_inizio + 1
    for b in blocchi:
        ws.cell(row=riga, column=1, value=b["etichetta"]).font = Font(bold=True)
        riga += 1
        for v in b["valori"]:
            if v is not None:
                ws.cell(row=riga, column=1, value=v)
            riga += 1
        riga += 2  # doppia riga vuota = blocco terminato


def make_sorgenti() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPIANTO"
    _scrivi_blocchi(
        ws,
        "IMPIANTO MOCK",
        [
            {"etichetta": "VELOCITA'", "valori": [100, 102, 101, None, 105]},
            {"etichetta": "PRESSIONE", "valori": [3.2, 3.4]},
            {"etichetta": "TEMPERATURA", "valori": [22, 23, 21, 22]},
            {"etichetta": "LIVELLO", "valori": [55]},
        ],
    )
    wb.save(SOURCE / "impianto_mock.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "AMBIENTE"
    _scrivi_blocchi(
        ws,
        "AMBIENTE MOCK",
        [
            {"etichetta": "UMIDITA'", "valori": [51, 49, 53]},
            {"etichetta": "CO2", "valori": [420, 415, 418]},
            {"etichetta": "ILLUMINAZIONE", "valori": [300]},
        ],
    )
    wb.save(SOURCE / "ambiente_mock.xlsx")


def make_scheda() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "SCHEDA"
    title = ws["B1"]
    title.value = "SCHEDA DI MISURA MOCK"
    title.font = Font(bold=True, size=13)

    # Struttura comoda: blocco di etichette statico (non grassetto) + voci sparse.
    intestazioni = {
        "B3": "VELOCITA'",
        "E7": "PRESSIONE",
        "C12": "TEMPERATURA",
        "F15": "LIVELLO",
        "B19": "UMIDITA'",
        "E21": "CO2",
    }
    for coord, testo in intestazioni.items():
        cell = ws[coord]
        cell.value = testo
        cell.font = Font(bold=True)

    # Valore vecchio sotto TEMPERATURA: la compilazione deve svuotarlo e riscriverlo.
    ws["C13"] = 99
    ws["C14"] = 98

    wb.save(DEST / "scheda_mock.xlsx")


if __name__ == "__main__":
    SOURCE.mkdir(parents=True, exist_ok=True)
    DEST.mkdir(parents=True, exist_ok=True)
    make_sorgenti()
    make_scheda()
    print("Mockup creati:")
    for p in sorted(SOURCE.glob("*.xlsx")) + sorted(DEST.glob("*.xlsx")):
        print(f"  {p}")