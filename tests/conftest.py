import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
APP_DIR = ROOT / "app"


@pytest.fixture()
def repo_root() -> Path:
    return ROOT


def _make_source(path: Path, blocchi: list[tuple[str, list]]) -> None:
    """Costruisce un file sorgente a 1 colonna con il formato Overlord:
    etichetta in grassetto, valori sotto, 2 righe vuote consecutive = fine blocco.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    riga = 1
    for etichetta, valori in blocchi:
        ws.cell(row=riga, column=1, value=etichetta).font = Font(bold=True)
        riga += 1
        for v in valori:
            ws.cell(row=riga, column=1, value=v)
            riga += 1
        riga += 2  # due righe vuote = fine blocco
    wb.save(path)


@pytest.fixture()
def sorgente(tmp_path: Path) -> Path:
    p = tmp_path / "sorgente.xlsx"
    _make_source(
        p,
        [
            ("VELOCITA'", [100, 102, None, 105]),
            ("TEMPERATURA", [22, 23]),
        ],
    )
    return p


def _make_scheda(path: Path, voci: dict[str, str]) -> None:
    """Costruisce una scheda con etichette sparse in grassetto."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    for coord, testo in voci.items():
        cell = ws[coord]
        cell.value = testo
        cell.font = Font(bold=True)
    wb.save(path)


@pytest.fixture()
def scheda(tmp_path: Path) -> Path:
    p = tmp_path / "scheda.xlsx"
    _make_scheda(p, {"B3": "VELOCITA'", "C12": "TEMPERATURA"})
    return p