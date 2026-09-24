"""Logica di lettura/scrittura dei file Excel.

Regola degli argomenti (sorgente e scheda):
- etichette scritte in grassetto (stile bold) = argomenti;
- sotto ogni etichetta, valori consecutivi nella stessa colonna;
- una singola riga vuota dentro la sequenza = buco dati (si salta);
- due righe vuote consecutive = la colonna/argomento è completata.

Nel sorgente gli argomenti stanno in una colonna verticale (colonna 1).
Nella scheda destinazione le etichette in grassetto possono stare in
qualunque posizione del foglio; si compila trovando l'etichetta con lo
stesso testo e scrivendo i valori nelle celle sotto di essa.
"""
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font


@dataclass
class Argomento:
    etichetta: str
    riga: int
    colonna: int = 1


def _cell_is_bold(cell) -> bool:
    font: Font | None = cell.font
    if font is None:
        return False
    return bool(font.bold)


def carica_workbook(path: Path):
    return load_workbook(path, data_only=False)


def trova_argomenti(path: Path, colonna: int = 1) -> list[Argomento]:
    """Etichette in grassetto nella colonna verticale del sorgente.

    Un grassetto è un argomento solo se la cella nella riga sotto non è a
    sua volta in grassetto (evita di scambiare i titoli per argomenti).
    """
    wb = carica_workbook(path)
    ws = wb.active
    risultato: list[Argomento] = []
    for row in ws.iter_rows(min_row=1, min_col=colonna, max_col=colonna):
        for cell in row:
            if cell.value is None or not _cell_is_bold(cell):
                continue
            sotto = ws.cell(row=cell.row + 1, column=colonna)
            if sotto.value is not None and _cell_is_bold(sotto):
                continue
            etichetta = str(cell.value).strip()
            if etichetta:
                risultato.append(Argomento(etichetta, cell.row, colonna))
    wb.close()
    return risultato


def trova_argomenti_ovunque(path: Path) -> list[Argomento]:
    """Tutte le etichette in grassetto del foglio attivo (qualsiasi colonna).

    Usata per elencare le voci sparse della scheda destinazione.
    """
    wb = carica_workbook(path)
    ws = wb.active
    risultato: list[Argomento] = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or not _cell_is_bold(cell):
                continue
            sz = cell.font.size if cell.font else None
            if sz is not None and sz > 11.5:
                continue
            sotto = ws.cell(row=cell.row + 1, column=cell.column)
            if sotto.value is not None and _cell_is_bold(sotto):
                continue
            etichetta = str(cell.value).strip()
            if etichetta:
                risultato.append(
                    Argomento(etichetta, cell.row, cell.column)
                )
    wb.close()
    return risultato


def estrai_colonna(path: Path, riga_etichetta: int, colonna: int = 1) -> list:
    """Valori sotto l'etichetta, fino alle due righe vuote consecutive."""
    wb = carica_workbook(path)
    ws = wb.active
    valori: list = []
    vuote = 0
    for row in ws.iter_rows(
        min_row=riga_etichetta + 1, min_col=colonna, max_col=colonna,
        values_only=True,
    ):
        valore = row[0]
        if valore is None or (isinstance(valore, str) and not valore.strip()):
            vuote += 1
            if vuote >= 2:
                break
            continue
        vuote = 0
        valori.append(valore)
    wb.close()
    return valori


def trova_cella_etichetta(path: Path, testo: str):
    """Posizione (riga, colonna) del grassetto col testo dato, o None."""
    wb = carica_workbook(path)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if (
                cell.value is not None
                and _cell_is_bold(cell)
                and str(cell.value).strip() == str(testo).strip()
            ):
                wb.close()
                return cell.row, cell.column
    wb.close()
    return None


def scrivi_sotto(path: Path, riga: int, colonna: int, valori: list) -> int:
    """Svuota il blocco sotto l'etichetta e ci scrive i nuovi valori.

    Il vecchio blocco viene individuato con la stessa regola della lettura:
    valori consecutivi fino alle due righe vuote consecutive.
    La scrittura si ferma prima di un'eventuale etichetta in grassetto
    sottostante (non la sovrascrive mai). Ritorna il numero di valori scritti.
    """
    wb = carica_workbook(path)
    ws = wb.active
    r = riga + 1
    vuote = 0
    cleared = 0
    while r <= ws.max_row:
        cell = ws.cell(row=r, column=colonna)
        v = cell.value
        if v is None or (isinstance(v, str) and not v.strip()):
            vuote += 1
            if vuote >= 2:
                break
        else:
            vuote = 0
            cell.value = None
            cleared += 1
        r += 1
    scritti = 0
    for i, v in enumerate(valori, start=riga + 1):
        cell = ws.cell(row=i, column=colonna)
        if _cell_is_bold(cell):
            break
        cell.value = v
        scritti += 1
    wb.close()
    wb.save(path)
    return scritti


def crea_voce(path: Path, testo: str, riga: int, colonna: int) -> None:
    """Crea (o aggiorna) una etichetta in grassetto vuota nella scheda."""
    wb = carica_workbook(path)
    ws = wb.active
    cell = ws.cell(row=riga, column=colonna, value=testo)
    cell.font = Font(bold=True)
    wb.close()
    wb.save(path)


def colonna_a_numero(colonna) -> int:
    """Accetta 'B' o 2 e restituisce un numero di colonna."""
    if isinstance(colonna, int):
        return colonna
    testo = str(colonna).strip().upper()
    n = 0
    for ch in testo:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Colonna non valida: {colonna}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n