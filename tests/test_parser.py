"""Test unit del parser Overlord (isolati su file temporanei)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from parser import (  # noqa: E402
    _cell_is_bold,
    colonna_a_numero,
    crea_voce,
    estrai_colonna,
    trova_argomenti,
    trova_argomenti_ovunque,
    trova_cella_etichetta,
    scrivi_sotto,
)


def test_trova_argomenti_individua_etichette(sorgente):
    args = trova_argomenti(sorgente)
    assert [(a.etichetta, a.riga) for a in args] == [("VELOCITA'", 1), ("TEMPERATURA", 8)]


def test_estrai_colonna_salta_buchi_e_ferma_a_due_vuote(sorgente):
    # VELOCITA' sta alla riga 1: i valori 100,102,(buco),105
    valori = estrai_colonna(sorgente, 1)
    assert valori == [100, 102, 105]


def test_estrai_colonna_ferma_al_blocco_successivo(sorgente):
    valori = estrai_colonna(sorgente, 1)
    assert valori == [100, 102, 105]  # non include TEMPERATURA


def test_trova_cella_etichetta(scheda):
    pos = trova_cella_etichetta(scheda, "TEMPERATURA")
    assert pos == (12, 3)
    assert trova_cella_etichetta(scheda, "INESISTENTE") is None


def test_trova_argomenti_ovunque_coglie_etichette_sparse(scheda):
    voci = trova_argomenti_ovunque(scheda)
    assert [(v.etichetta, v.riga, v.colonna) for v in voci] == [
        ("VELOCITA'", 3, 2),
        ("TEMPERATURA", 12, 3),
    ]


def test_crea_voce_aggiunge_bold(scheda):
    crea_voce(scheda, "FOTOMETRIA", 28, 2)
    voci = trova_argomenti_ovunque(scheda)
    assert ("FOTOMETRIA", 28, 2) in [(v.etichetta, v.riga, v.colonna) for v in voci]


def test_colonna_a_numero():
    assert colonna_a_numero("A") == 1
    assert colonna_a_numero("B") == 2
    assert colonna_a_numero("AA") == 27
    assert colonna_a_numero(2) == 2
    with pytest.raises(ValueError):
        colonna_a_numero("1")


def test_scrivi_sotto_svuota_e_riscrive(tmp_path, scheda):
    # La scheda ha VELOCITA' in B3: sotto (B4) abbiamo dei valori vecchi.
    from openpyxl import load_workbook

    wb = load_workbook(scheda)
    ws = wb.active
    ws["B4"], ws["B5"], ws["B6"] = 99, 98, 97
    wb.save(scheda)
    wb.close()

    scritti = scrivi_sotto(scheda, 3, 2, [100, 101])
    assert scritti == 2

    wb = load_workbook(scheda)
    ws = wb.active
    assert ws["B4"].value == 100
    assert ws["B5"].value == 101
    assert ws["B6"].value is None  # vecchio valore svuotato
    assert ws["B3"].value == "VELOCITA'"  # etichetta intatta
    wb.close()


def test_scrivi_sotto_piu_lungo_definisce_nuovi_valori(tmp_path, scheda):
    import openpyxl

    wb = openpyxl.load_workbook(scheda)
    ws = wb.active
    ws["B4"], ws["B5"] = 99, 98
    wb.save(scheda)
    wb.close()

    scritti = scrivi_sotto(scheda, 3, 2, [100, 101, 102, 103])
    assert scritti == 4
    wb = openpyxl.load_workbook(scheda)
    ws = wb.active
    assert [ws["B4"].value, ws["B5"].value, ws["B6"].value, ws["B7"].value] == [
        100, 101, 102, 103,
    ]
    wb.close()


def test_cell_is_bold_con_stile_assente():
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    assert _cell_is_bold(ws["Z99"]) is False