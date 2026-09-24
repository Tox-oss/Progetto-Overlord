"""Test unit del parser Overlord (isolati su file temporanei)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from parser import (  # noqa: E402
    _cell_is_bold,
    _spazio_libero,
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


def test_scrivi_sotto_non_cancella_etichetta_sotto(scheda_stessa_colonna):
    """Finding 1: lo svuotamento NON deve cancellare la successiva etichetta
    in grassetto (ne' i suoi valori) se non ci sono due righe vuote in mezzo."""
    from openpyxl import load_workbook

    scritti = scrivi_sotto(scheda_stessa_colonna, 3, 2, [100, 101, 103])
    assert scritti == 3

    wb = load_workbook(scheda_stessa_colonna)
    ws = wb.active
    assert ws["B3"].value == "VELOCITA'"          # etichetta originale intatta
    assert ws["B4"].value == 100
    assert ws["B5"].value == 101
    assert ws["B6"].value == 103                  # si estende finche' non urta il bold
    assert ws["B7"].value == "PRESSIONE"          # LA SECONDA ETICHETTA RIMANE
    assert _cell_is_bold(ws["B7"]) is True
    assert ws["B8"].value == 9                    # e il suo valore resta
    wb.close()


def test_scrivi_sotto_salta_celle_unite(scheda_mergata):
    """Finding 9: le celle unite non devono far esplodere la scrittura."""
    from openpyxl import load_workbook

    scritti = scrivi_sotto(scheda_mergata, 3, 2, [10, 11])
    assert scritti == 1  # B4 libero, B5:B6 unito -> si ferma prima dell'unica cella

    wb = load_workbook(scheda_mergata)
    ws = wb.active
    assert ws["B3"].value == "VELOCITA'"
    assert ws["B4"].value == 10
    wb.close()


def test_scrivi_sotto_non_sovrascrive_zona_unita(tmp_path):
    """Finding 9: una cella unita sotto l'etichetta non fa esplodere la
    scrittura (AttributeError su MergedCell) e ne' viene sovrascritta."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "scheda_unita.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "V"
    ws["A1"].font = Font(bold=True)
    ws.merge_cells("A2:A3")
    wb.save(p)
    wb.close()

    scritti = scrivi_sotto(p, 1, 1, [5, 6])
    assert scritti == 0  # subito sotto c'e' una cella unita: niente scrittura

    from openpyxl import load_workbook

    wb = load_workbook(p)
    ws = wb.active
    assert ws["A1"].value == "V"
    wb.close()


def test_trova_argomenti_scarta_titolo_font_grande(sorgente_titolo_grande):
    """Finding 8: l'euristica font>11.5 vale anche nei sorgenti."""
    args = trova_argomenti(sorgente_titolo_grande)
    assert [(a.etichetta, a.riga) for a in args] == [("VELOCITA'", 4)]


def test_estrai_colonna_non_restituisce_calcolo_come_testo(tmp_path):
    """Finding 10: la stringa della formula (=10+1) non deve comparire come
    valore: con data_only=True le celle formula senza valore in cache vengono
    lette come vuote, mai come testo."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "formula.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "CALCOLO"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = "=10+1"
    wb.save(p)
    wb.close()

    valori = estrai_colonna(p, 1)
    assert "=10+1" not in [str(x) for x in valori]


def test_spazio_libero_limita_davanti_a_bold(scheda_stessa_colonna):
    """Lo spazio libero sotto un'etichetta termina prima del grassetto."""
    from openpyxl import load_workbook

    wb = load_workbook(scheda_stessa_colonna)
    ws = wb.active
    assert _spazio_libero(ws, 3, 2) == 3  # B4,B5,B6 (fermo a B7 bold)
    wb.close()