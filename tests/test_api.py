"""Test API Overlord: cicli end-to-end over rotte, con file isolati in tmp_path."""
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent

sys.path.insert(0, str(ROOT / "app"))


@pytest.fixture()
def app(tmp_path: Path, monkeypatch):
    import app as appmod

    # Reindirizza cartelle/config/state su una dir temporanea isolata.
    files = tmp_path / "files"
    (files / "source").mkdir(parents=True)
    (files / "dest").mkdir()
    (files / "export").mkdir()
    (files / "config").mkdir()
    (files / "state").mkdir()

    monkeypatch.setattr(appmod, "SOURCE_DIR", files / "source")
    monkeypatch.setattr(appmod, "DEST_DIR", files / "dest")
    monkeypatch.setattr(appmod, "EXPORT_DIR", files / "export")
    monkeypatch.setattr(appmod, "CONFIG_PATH", files / "config" / "selezione.json")
    monkeypatch.setattr(appmod, "STATE_PATH", files / "state" / "last_value.json")

    # Mockup minimi (1 sorgente, 1 scheda).
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws["A1"] = "VELOCITA'"
    ws["A1"].font = Font(bold=True)
    ws["A2"], ws["A3"] = 100, 101
    wb.save(files / "source" / "imp.xlsx")

    wb = Workbook()
    ws = wb.active
    ws["B3"] = "VELOCITA'"
    ws["B3"].font = Font(bold=True)
    # valore vecchio sotto l'etichetta: deve essere svuotato alla compilazione
    ws["B4"], ws["B5"], ws["B6"] = 99, 98, 97
    wb.save(files / "dest" / "scheda.xlsx")

    appmod.app.config["TESTING"] = True
    return appmod


@pytest.fixture()
def client(app):
    return app.app.test_client()


def _get(client, url):
    return client.get(url)


def test_sorgenti_destinazioni(client):
    assert _get(client, "/api/sorgenti").get_json() == ["imp.xlsx"]
    assert _get(client, "/api/destinazioni").get_json() == ["scheda.xlsx"]


def test_argomenti_e_404(client):
    r = _get(client, "/api/argomenti?file=imp.xlsx")
    assert r.status_code == 200
    data = r.get_json()
    assert data[0]["etichetta"] == "VELOCITA'"
    assert data[0]["valori"] == [100, 101]
    assert _get(client, "/api/argomenti?file=x.xlsx").status_code == 404


def test_scheda_voci(client):
    data = _get(client, "/api/scheda?file=scheda.xlsx").get_json()
    assert any(v["etichetta"] == "VELOCITA'" and v["posizione"] == "B3" for v in data)


def test_aggiungi_arriva_in_config(client):
    r = client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    assert r.status_code == 200
    assert r.get_json()["argomenti"][0]["riga_etichetta"] == 1


def test_aggiungi_duplicato_ignorato(client):
    payload = {"sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1}
    client.post("/api/aggiungi", json=payload)
    client.post("/api/aggiungi", json=payload)
    cfg = _get(client, "/api/config").get_json()
    assert len(cfg["argomenti"]) == 1


def test_aggiungi_riga_invalida(client):
    r = client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": "abc",
    })
    assert r.status_code == 400


def test_compila_svuota_e_riscrive(app, client):
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    client.post("/api/destinazione", json={"destinazione": "scheda.xlsx"})
    r = client.post("/api/compila")
    assert r.status_code == 200
    esiti = r.get_json()["esiti"]
    assert esiti[0]["trovata"] is True
    assert esiti[0]["scritti"] == 2

    # Verifica fisica sul file isolato: valori nuovi sotto B3, vecchi svuotati.
    import openpyxl

    wb = openpyxl.load_workbook(app.DEST_DIR / "scheda.xlsx")
    ws = wb.active
    assert ws["B3"].value == "VELOCITA'"
    assert ws["B4"].value == 100
    assert ws["B5"].value == 101
    assert ws["B6"].value is None
    wb.close()


def test_export_traversal_bloccato(app, client):
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    r = client.post("/api/export", json={"nome": "../../fuga.xlsx"})
    assert r.status_code == 200
    f = r.get_json()["file"]
    assert "../" not in f
    assert "/" not in f
    assert (app.EXPORT_DIR / f).exists()


def test_nome_foglio_invalido_sanificato(client):
    # etichetta con carattere vietato in un foglio Excel: non deve rompere l'export
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VEL/OCITA'", "riga_etichetta": 1,
    })
    r = client.post("/api/export", json={"nome": "ok.xlsx"})
    assert r.status_code == 200


def test_export_download(client):
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    f = client.post("/api/export", json={"nome": "out.xlsx"}).get_json()["file"]
    r = _get(client, f"/api/download/{f}")
    assert r.status_code == 200


def test_download_sconosciuto(client):
    assert _get(client, "/api/download/nope.xlsx").status_code == 404


def test_elabora_prima_volta_poi_nessuna_variazione(client):
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    client.post("/api/destinazione", json={"destinazione": "scheda.xlsx"})
    r1 = client.post("/api/elabora")
    assert r1.status_code == 200
    assert r1.get_json()["ok"] is True
    r2 = client.post("/api/elabora")
    assert r2.get_json()["ok"] is False
    assert "nessuna variazione" in r2.get_json()["motivo"]


def test_elabora_config_incompleta(client):
    r = client.post("/api/elabora")
    assert r.status_code == 200
    assert r.get_json()["ok"] is False


def test_rimuovi(client):
    client.post("/api/aggiungi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'", "riga_etichetta": 1,
    })
    r = client.post("/api/rimuovi", json={
        "sorgente": "imp.xlsx", "etichetta": "VELOCITA'",
    })
    assert r.get_json()["argomenti"] == []


def test_modella_crea_voce(client):
    r = client.post("/api/modella", json={
        "destinazione": "scheda.xlsx", "testo": "FOTOMETRIA", "riga": 28, "colonna": "B",
    })
    assert r.status_code == 200
    data = _get(client, "/api/scheda?file=scheda.xlsx").get_json()
    assert any(v["etichetta"] == "FOTOMETRIA" and v["posizione"] == "B28" for v in data)


def test_modella_colonna_invalida(client):
    r = client.post("/api/modella", json={
        "destinazione": "scheda.xlsx", "testo": "X", "riga": 1, "colonna": "1",
    })
    assert r.status_code == 400