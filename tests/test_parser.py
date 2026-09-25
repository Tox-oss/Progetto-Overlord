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
    assert colonna_a_numero(" b ") == 2
    with pytest.raises(ValueError):
        colonna_a_numero("1")
    with pytest.raises(ValueError):
        colonna_a_numero("")
    with pytest.raises(ValueError):
        colonna_a_numero("  ")


def test_trova_cella_etichetta_case_sensitive_e_duplicati(tmp_path):
    """La ricerca è case-sensitive e in caso di etichette duplicate restituisce
    la PRIMA nel foglio."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "dupl.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "CO2"
    ws["B3"].font = Font(bold=True)
    ws["E7"] = "co2"
    ws["E7"].font = Font(bold=True)
    ws["E9"] = "CO2"
    ws["E9"].font = Font(bold=True)
    wb.save(p)
    wb.close()

    assert trova_cella_etichetta(p, "CO2") == (3, 2)   # prima occorrenza
    assert trova_cella_etichetta(p, "co2") == (7, 5)   # case-sensitive
    assert trova_cella_etichetta(p, "Co2") is None


def test_trova_cella_etichetta_ignora_titoli(tmp_path):
    """A1: come trova_argomenti, la ricerca nella scheda scarta i TITOLI:
    font > 11.5 o cella sottostante in grassetto (un argomento vero ha valori
    sotto). Prima compilava sotto un header grande al posto dell'argomento."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "titoli.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B2"] = "VELOCITA'"
    ws["B2"].font = Font(bold=True, size=16)   # titolo grande
    ws["B3"] = 999
    ws["B6"] = "VELOCITA'"
    ws["B6"].font = Font(bold=True)            # vero argomento
    ws["B7"] = 1
    ws["B10"] = "X"                            # grassetto sopra un altro grassetto
    ws["B10"].font = Font(bold=True)
    ws["B11"] = "Y"
    ws["B11"].font = Font(bold=True)
    wb.save(p)
    wb.close()

    assert trova_cella_etichetta(p, "VELOCITA'") == (6, 2)   # non (2, 2)
    assert trova_cella_etichetta(p, "X") is None             # sotto Y in grassetto
    assert trova_cella_etichetta(p, "Y") == (11, 2)


def test_colonna_a_numero_rifiuta_zero_e_negativi():
    """A2: le colonne iniziano da 1; 0 e negativi non sono mai validi."""
    with pytest.raises(ValueError):
        colonna_a_numero(0)
    with pytest.raises(ValueError):
        colonna_a_numero(-3)
    with pytest.raises(ValueError):
        colonna_a_numero("0")


def test_colonna_a_numero_booleani_accettati():
    assert colonna_a_numero(True) == 1  # True == 1 (parentesi: comportamento attuale)


def test_crea_voce_rifiuta_cella_unita(tmp_path):
    """A3: `crea_voce` rifiuta con ValueError una cella dentro un merge
    (sia anchor che MergedCell) invece di scrivere a metà dell'unione."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "unita.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "V"
    ws["B3"].font = Font(bold=True)
    ws.merge_cells("A2:A4")
    wb.save(p)
    wb.close()

    with pytest.raises(ValueError):
        crea_voce(p, "X", 2, 1)  # anchor inside merge
    with pytest.raises(ValueError):
        crea_voce(p, "X", 4, 1)  # MergedCell non-attiva


def test_crea_voce_pulisce_testo(scheda):
    crea_voce(scheda, "  FOTOMETRIA  ", 28, 2)
    from openpyxl import load_workbook

    wb = load_workbook(scheda)
    assert wb.active["B28"].value == "FOTOMETRIA"
    wb.close()


def test_scrivi_sotto_accanto_a_zona_unita(tmp_path):
    """A4: una cella unita in una colonna LATERALE non blocca la scrittura
    nella colonna dell'etichetta."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "laterali.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "VELOCITA'"
    ws["B3"].font = Font(bold=True)
    ws.merge_cells("C4:C6")
    ws.merge_cells("C9:C10")
    wb.save(p)
    wb.close()

    from openpyxl import load_workbook

    wb = load_workbook(p)
    ws = wb.active
    assert _spazio_libero(ws, 3, 2) >= 3  # la B è libera anche con merge a fianco
    wb.close()

    scritti = scrivi_sotto(p, 3, 2, [1, 2, 3])
    assert scritti == 3

    wb = load_workbook(p)
    ws = wb.active
    assert [ws["B4"].value, ws["B5"].value, ws["B6"].value] == [1, 2, 3]
    wb.close()


def test_multi_foglio_legge_solo_il_primo(tmp_path):
    """A5: i file con più fogli vengono letti SOLO sul primo foglio (active);
    il secondo foglio è ignorato."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "multi.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "FOGLIO1"
    ws1["A1"] = "VELOCITA'"
    ws1["A1"].font = Font(bold=True)
    ws1["A2"] = 100
    ws2 = wb.create_sheet("FOGLIO2")
    ws2["A1"] = "PRESSIONE"
    ws2["A1"].font = Font(bold=True)
    ws2["A2"] = 9
    wb.save(p)
    wb.close()

    args = trova_argomenti(p)
    assert [(a.etichetta, a.riga) for a in args] == [("VELOCITA'", 1)]
    assert estrai_colonna(p, 1) == [100]
    assert trova_cella_etichetta(p, "PRESSIONE") is None
    assert trova_cella_etichetta(p, "VELOCITA'") == (1, 1)


def test_estrai_colonna_oltre_max_row(tmp_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "alto.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "V"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = 7
    wb.save(p)
    wb.close()

    assert estrai_colonna(p, 999) == []


def test_estrai_colonna_booleano_e_decimali(tmp_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "tipi.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "V"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = True
    ws["A3"] = 1.5
    ws["A4"] = 2
    wb.save(p)
    wb.close()

    assert estrai_colonna(p, 1) == [True, 1.5, 2]


def test_spazio_libero_soglia_zero(scheda_stessa_colonna):
    from openpyxl import load_workbook

    wb = load_workbook(scheda_stessa_colonna)
    ws = wb.active
    assert _spazio_libero(ws, 3, 2, soglia=0) == 0
    wb.close()


def test_spazio_libero_fermo_a_due_vuote(tmp_path):
    """Se sotto l'etichetta ci sono due righe vuote consecutive, la conta."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "duevuote.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "V"
    ws["B3"].font = Font(bold=True)
    ws["B4"], ws["B5"] = 1, 2
    ws["B8"] = "ALTRO"
    ws["B8"].font = Font(bold=True)
    wb.save(p)
    wb.close()

    from openpyxl import load_workbook

    wb = load_workbook(p)
    ws = wb.active
    # spazio illimitato in realtà fino al bold riga 8; con soglia basta
    assert _spazio_libero(ws, 3, 2, soglia=3) == 3
    wb.close()


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


def test_spazio_libero_oltre_due_vuote_davanti_a_bold(tmp_path):
    """BUG 5 (set 26): sul DEST le due righe vuote consecutive NON riducono lo
    spazio scrivibile: il blocco cresce fino al grassetto/ostacolo. Le due
    righe vuote sono il delimitatore dei SORGENTI (regola di lettura), il DEST
    non viene mai ri-letto a blocchi (le voci si trovano per etichetta)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    p = tmp_path / "duevuote_bold.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "VELOCITA'"
    ws["B3"].font = Font(bold=True)
    ws["B4"] = 100
    ws["B5"], ws["B6"] = None, None   # due righe vuote = delimitatore nei SORGENTI
    ws["B7"] = "PRESSIONE"
    ws["B7"].font = Font(bold=True)
    wb.save(p)
    wb.close()

    from openpyxl import load_workbook

    wb = load_workbook(p)
    ws = wb.active
    # B4 (valore), B5, B6 (vuote): tutte scrivibili, il conteggio si ferma a B7
    assert _spazio_libero(ws, 3, 2) == 3
    wb.close()

    scritti = scrivi_sotto(p, 3, 2, [10, 20, 30])
    assert scritti == 3  # riempie anche le due vuote, fermandosi davanti a B7

    wb = load_workbook(p)
    ws = wb.active
    assert [ws["B4"].value, ws["B5"].value, ws["B6"].value] == [10, 20, 30]
    assert ws["B7"].value == "PRESSIONE"   # l'etichetta sotto non viene toccata
    assert ws["B3"].value == "VELOCITA'"
    wb.close()