import json
import os
import re
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from parser import (
    colonna_a_numero,
    crea_voce,
    estrai_colonna,
    trova_argomenti,
    trova_argomenti_ovunque,
    trova_cella_etichetta,
    scrivi_sotto,
)

PROJECT = Path(__file__).resolve().parent.parent
FILES = PROJECT / "files"
SOURCE_DIR = FILES / "source"
DEST_DIR = FILES / "dest"
EXPORT_DIR = FILES / "export"
CONFIG_PATH = FILES / "config" / "selezione.json"
STATE_PATH = FILES / "state" / "last_value.json"

for d in (SOURCE_DIR, DEST_DIR, EXPORT_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


@app.errorhandler(Exception)
def _errore_generico(exc):  # noqa: BLE001
    traceback.print_exc()
    return jsonify({"errore": "Errore interno del server"}), 500


def _xlsx_files(cartella: Path) -> list[str]:
    return sorted(p.name for p in cartella.glob("*.xlsx"))


def _leggi_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            cfg = {}
    else:
        cfg = {}
    cfg.setdefault("destinazione", None)
    cfg.setdefault("argomenti", [])
    return cfg


def _scrivi_config(payload: dict) -> None:
    tmp = CONFIG_PATH.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(tmp, CONFIG_PATH)


def _leggi_state() -> dict:
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def _scrivi_state(ultimi: list) -> None:
    payload = {
        "ultimi_valori": ultimi,
        "data": datetime.now().isoformat(),
    }
    tmp = STATE_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, STATE_PATH)


def _estraai_tutti(cfg: dict) -> list[dict]:
    estratti = []
    for a in cfg.get("argomenti", []):
        path = SOURCE_DIR / a["sorgente"]
        if not path.exists():
            a_ = dict(a)
            a_["valori"] = []
            a_["errore"] = "sorgente mancante"
            estratti.append(a_)
            continue
        a_ = dict(a)
        try:
            a_["valori"] = estrai_colonna(path, int(a_["riga_etichetta"]))
        except (TypeError, ValueError):
            a_["valori"] = []
            a_["errore"] = "riga_etichetta non valida"
        estratti.append(a_)
    return estratti


def _compila_tutti(cfg: dict, estratti: list[dict]) -> list[dict]:
    if not cfg.get("destinazione"):
        return [{"errore": "Destinazione non configurata"}]
    dest = DEST_DIR / cfg["destinazione"]
    if not dest.exists():
        return [{"errore": "File di destinazione non trovato"}]
    esiti = []
    for e in estratti:
        pos = trova_cella_etichetta(dest, e["etichetta"])
        if pos is None:
            esiti.append(
                {
                    "sorgente": e["sorgente"],
                    "etichetta": e["etichetta"],
                    "trovata": False,
                }
            )
            continue
        riga, colonna = pos
        scritti = scrivi_sotto(dest, riga, colonna, e.get("valori", []))
        esiti.append(
            {
                "sorgente": e["sorgente"],
                "etichetta": e["etichetta"],
                "trovata": True,
                "riga": riga,
                "colonna": colonna,
                "valori": e.get("valori", []),
                "scritti": scritti,
            }
        )
    return esiti


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/sorgenti")
def api_sorgenti():
    return jsonify(_xlsx_files(SOURCE_DIR))


@app.get("/api/destinazioni")
def api_destinazioni():
    return jsonify(_xlsx_files(DEST_DIR))


@app.get("/api/argomenti")
def api_argomenti():
    nome = request.args.get("file", "")
    path = SOURCE_DIR / nome
    if not path.exists():
        return jsonify({"errore": "File sorgente non trovato"}), 404
    argomenti = [
        {
            "etichetta": a.etichetta,
            "riga": a.riga,
            "valori": estrai_colonna(path, a.riga),
        }
        for a in trova_argomenti(path)
    ]
    return jsonify(argomenti)


@app.get("/api/scheda")
def api_scheda():
    nome = request.args.get("file", "")
    path = DEST_DIR / nome
    if not path.exists():
        return jsonify({"errore": "Scheda non trovata"}), 404
    voci = [
        {
            "etichetta": a.etichetta,
            "riga": a.riga,
            "colonna": a.colonna,
            "posizione": f"{_colonna_lettera(a.colonna)}{a.riga}",
        }
        for a in trova_argomenti_ovunque(path)
    ]
    return jsonify(voci)


def _colonna_lettera(n: int) -> str:
    s = ""
    while n:
        n, resto = divmod(n - 1, 26)
        s = chr(ord("A") + resto) + s
    return s


def _nome_foglio_valido(etichetta: str) -> str:
    """Nome foglio Excel ammissibile (niente caratteri vietati, max 31 char)."""
    nome = re.sub(r"[\\/*?:\[\]]", "_", str(etichetta)).strip() or "ARGOMENTO"
    return nome[:31]


def _nome_export_sicuro(nome: str) -> str:
    """Ricava un nome file valido dentro EXPORT_DIR (niente path traversal né
    caratteri illegali)."""
    base = Path(nome or "").name.strip() or "esportazione.xlsx"
    base = re.sub(r"[\\/*?:<>\"|]", "_", base)
    if not base.lower().endswith(".xlsx"):
        base += ".xlsx"
    return base


@app.get("/api/config")
def api_config():
    return jsonify(_leggi_config())


@app.post("/api/aggiungi")
def api_aggiungi():
    payload = request.get_json(force=True)
    sorgente, etichetta = payload.get("sorgente"), payload.get("etichetta")
    riga = payload.get("riga_etichetta")
    if not all((sorgente, etichetta, riga)):
        return jsonify({"errore": "Parametri mancanti"}), 400
    try:
        riga_n = int(riga)
    except (TypeError, ValueError):
        return jsonify({"errore": f"Riga non valida: {riga}"}), 400
    if riga_n < 1:
        return jsonify({"errore": "Riga non valida"}), 400
    path = SOURCE_DIR / sorgente
    if not path.exists():
        return jsonify({"errore": "File sorgente non trovato"}), 404
    cfg = _leggi_config()
    duplicato = any(
        a["sorgente"] == sorgente and a["etichetta"] == etichetta
        for a in cfg["argomenti"]
    )
    if not duplicato:
        cfg["argomenti"].append(
            {"sorgente": sorgente, "etichetta": etichetta, "riga_etichetta": riga_n}
        )
        _scrivi_config(cfg)
    return jsonify(cfg)


@app.post("/api/rimuovi")
def api_rimuovi():
    payload = request.get_json(force=True)
    sorgente, etichetta = payload.get("sorgente"), payload.get("etichetta")
    cfg = _leggi_config()
    cfg["argomenti"] = [
        a
        for a in cfg["argomenti"]
        if not (a["sorgente"] == sorgente and a["etichetta"] == etichetta)
    ]
    _scrivi_config(cfg)
    return jsonify(cfg)


@app.post("/api/destinazione")
def api_destinazione():
    payload = request.get_json(force=True)
    cfg = _leggi_config()
    cfg["destinazione"] = payload.get("destinazione")
    _scrivi_config(cfg)
    return jsonify(cfg)


@app.post("/api/modella")
def api_modella():
    payload = request.get_json(force=True)
    scheda, testo, riga = (
        payload.get("destinazione"),
        payload.get("testo"),
        payload.get("riga"),
    )
    colonna = payload.get("colonna")
    if not all((scheda, testo, riga, colonna)):
        return jsonify({"errore": "Parametri mancanti"}), 400
    path = DEST_DIR / scheda
    if not path.exists():
        return jsonify({"errore": "Scheda non trovata"}), 404
    try:
        riga_n = int(riga)
        colonna_n = colonna_a_numero(colonna)
    except ValueError as exc:
        return jsonify({"errore": str(exc)}), 400
    if riga_n < 1 or riga_n > 1_048_576:
        return jsonify({"errore": f"Riga fuori range: {riga_n}"}), 400
    if colonna_n < 1 or colonna_n > 16_384:
        return jsonify({"errore": f"Colonna fuori range: {colonna}"}), 400
    try:
        crea_voce(path, str(testo), riga_n, colonna_n)
    except ValueError as exc:
        return jsonify({"errore": str(exc)}), 400
    return jsonify({"ok": True})


@app.post("/api/compila")
def api_compila():
    cfg = _leggi_config()
    estratti = _estraai_tutti(cfg)
    esiti = _compila_tutti(cfg, estratti)
    return jsonify({"esiti": esiti})


@app.post("/api/export")
def api_export():
    """Salva con nome: esporta tutti gli argomenti configurati in un xlsx."""
    payload = request.get_json(force=True)
    nome = _nome_export_sicuro(payload.get("nome", ""))
    cfg = _leggi_config()
    estratti = _estraai_tutti(cfg)
    if not estratti:
        return jsonify({"errore": "Nessun argomento configurato"}), 400

    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    for e in estratti:
        nome_foglio = _nome_foglio_valido(e["etichetta"])
        ws = wb.create_sheet(title=nome_foglio)
        ws["A1"] = "ARGOMENTO"
        ws["B1"] = "VALORE"
        for i, v in enumerate(e.get("valori", []), start=2):
            ws.cell(row=i, column=1, value=e["etichetta"])
            ws.cell(row=i, column=2, value=v)

    dest = EXPORT_DIR / nome
    wb.save(dest)
    return jsonify({"ok": True, "file": nome, "argomenti": len(estratti)})


@app.get("/api/download/<nome_file>")
def api_download(nome_file: str):
    path = (EXPORT_DIR / nome_file).resolve()
    if not path.exists() or EXPORT_DIR not in path.parents:
        return jsonify({"errore": "File non trovato"}), 404
    return send_file(path, as_attachment=True)


@app.post("/api/elabora")
def api_elabora():
    """Bridge per l'automazione (n8n): se i valori sono cambiati, compila."""
    try:
        cfg = _leggi_config()
        if not cfg.get("argomenti") or not cfg.get("destinazione"):
            return jsonify({"ok": False, "motivo": "configurazione incompleta"}), 200
        estratti = _estraai_tutti(cfg)

        for e in estratti:
            e.pop("riga_etichetta", None)

        stato = _leggi_state()
        if stato.get("ultimi_valori") == estratti:
            return jsonify({"ok": False, "motivo": "nessuna variazione"}), 200

        esiti = _compila_tutti(cfg, estratti)
        if any(not x.get("trovata", True) for x in esiti):
            non_trovate = [
                x["etichetta"] for x in esiti if not x.get("trovata", True)
            ]
            return jsonify(
                {"ok": False, "motivo": "voci non trovate nella scheda", "voci": non_trovate}
            ), 200

        _scrivi_state(estratti)
        return jsonify({"ok": True, "esiti": esiti})
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        return jsonify({"errore": str(exc)}), 500


if __name__ == "__main__":
    porta = int(os.environ.get("PORT", 8010))
    app.run(host="0.0.0.0", port=porta, debug=True)