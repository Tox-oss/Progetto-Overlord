# Contesto - Overlord (ex "Progetto ExEL")

## 🎯 Obiettivo
Strumento che monitora dozzine di file Excel sorgente, estrae "argomenti" (etichette in grassetto con valori sotto) e compila un unico file Excel di destinazione ("scheda") in posizioni specifiche. Ruolo principale: rilevare modifiche ai file, prelevare valori selezionati e immetterli nella scheda da compilare. Futuro: automazione con n8n self-hosted e notifiche email.

## 📍 Stato corrente
App e API funzionanti e testate end-to-end (multi-sorgente → compila in posizioni giuste; dedup ok; modella scheda crea voci; export ok). Server Flask attivo in background su porta **8010**. n8n NON ancora attivato. Email non integrata. Prossimo passo (con l'utente): **validare la parser su file reali** e **attivare n8n**.

## 🗂 Struttura progetto
Cartella: `/home/vitalianorossi/Scrivania/Lavori/Progetto Overlord/` (rinominata da "Progetto ExEL")
- `files/source/` : sorgenti Excel (mockup: `impianto_mock.xlsx` = VELOCITA', PRESSIONE, TEMPERATURA, LIVELLO; `ambiente_mock.xlsx` = UMIDITA', CO2, ILLUMINAZIONE). Struttura: colonna verticale, etichetta in grassetto, valori consecutivi sotto; **1 riga vuota = buco dati da saltare; 2 righe vuote consecutive = blocco terminato**.
- `files/dest/scheda_mock.xlsx` : scheda finale come modulo LIBERO con etichette in grassetto SPARSE (VELOCITA' in B3, PRESSIONE in E7, TEMPERATURA in C12, LIVELLO in F15, UMIDITA' in B19, CO2 in E21); valori da scrivere SOTTO l'etichetta.
- `files/config/selezione.json` : configurazione corrente — attualmente `destinazione: null`, 1 argomento: CO2 (sorgente `ambiente_mock.xlsx`, riga_etichetta 8).
- `files/state/last_value.json` : stato per dedup — attualmente resettato (`ultimi_valori: null`, `data: null`).
- `files/export/` : esportazioni "Salva con nome".
- `app/` : web app Flask. `app.py` (API), `parser.py` (logica Excel), `templates/index.html` (UI multisezione con menu a tendina), `requirements.txt`, `.venv` (da NON committare).
- `tools/make_mockups.py` : rigenera i mockup.
- `n8n/docker-compose.yml` : compose ("selettore" build da Dockerfile.selettore, porta host 8010 → 8000; "n8n" con profilo `auto`, porta 5678, volumi `../files:/files` e `./.n8n:/home/node/.n8n`).
- `Dockerfile.selettore` : python:3.12-slim + flask + openpyxl, EXPOSE 8000.

## 🔑 Punti chiave
- API disponibili: GET `/api/sorgenti`, `/api/destinazioni`, `/api/argomenti?file=`, `/api/scheda?file=`, `/api/config`; POST `/api/aggiungi`, `/api/rimuovi`, `/api/destinazione`, `/api/modella`, `/api/compila`, `/api/export`, `/api/elabora` (bridge automazione: rilegge config, estrae tutti gli argomenti, confronta con state, se cambiati compila e aggiorna state; risponde "nessuna variazione" se invariato); GET `/api/download/<nome>`. (fonte: `app/app.py`)
- Funzioni parser: `_cell_is_bold`, `carica_workbook`, `trova_argomenti` (colonna A per sorgente), `trova_argomenti_ovunque` (per scheda), `estrai_colonna`, `trova_cella_etichetta`, `scrivi_sotto` (svuota blocco e riscrive), `crea_voce`, `colonna_a_numero`. (fonte: `app/parser.py`)
- **Porte**: web app su **8010** (8000 e 8001 occupate da altri progetti dell'utente: InventarioAuto ecc.). Confermata in ascolto su 8010. (verifica ambiente)
- **Regole parser da validare** con l'utente: scartare grassetti che NON sono argomenti (titoli) — euristica: cella sotto in grassetto → titolo; font size > 11.5 → titolo.
- **Compilazione**: testo etichetta uguale tra sorgente e scheda → scrivi valori sotto, svuotando il blocco vecchio (fonte: `app/parser.py` — `scrivi_sotto`).
- Al termine dei test: ripristinare mock/config/state con `tools/make_mockups.py` + reset JSON (config attualmente con destinazione null e un solo argomento CO2; state resettato). (fonte: `files/config/selezione.json`, `files/state/last_value.json`)
- n8n compose: selettore parte con `docker compose up -d`; n8n solo con `docker compose --profile auto up -d`. (fonte: `n8n/docker-compose.yml`)
- Git: push imminente su `https://github.com/Tox-oss/Progetto-Overlord.git` (branch main); `.gitignore` escluderà `.venv` e file temporanei.

## ⚠️ Criticità aperte
- [MEDIA] Euristiche parser (grassetto-titolo, font size) NON ancora validate su file reali dell'utente — da testare insieme (fonte: contesto utente).
- [INFO] n8n non avviato; email non integrata (fonte: contesto utente).
- [INFO] Config/state attualmente vuoti o parziali (residuo di test): destinazione null, un solo argomento, state resettato. Da ripristinare/popolare prima dell'uso reale.
- [BASSA] `scrivi_sotto` (fissata) non scrive più oltre una etichetta grassetto: se i valori da scrivere superano lo spazio disponibile, vengono scartati e segnalati nel campo `scritti` dell'esito (`api/compila`/`api/elabora`); comportamenti da confermare con l'utente.

## ✅ Correzioni applicate (bughunt esteso, set 24)
Applicati i principi delle istruzioni del cliente (`~/Scrivania/CLAUDE.md`, `DIRECTIVES.md`): verificare prima di dichiarare, autorità spec > piano > codice, non toccare gli invarianti del parser, isolamento dei test. Le regole del parser (1 riga vuota = buco; 2 vuote = fine blocco; etichetta grassetto sopra; scrittura sotto etichetta omonima) sono INVIOLATE e ora coperte da test.

1. **Export sicuro** (`app/app.py`): nuovo `_nome_export_sicuro` (solo basename, mai path traversal, caratteri illegali Windows `<>:"/\|?*` sostituiti, estensione `.xlsx` forzata) e `_nome_foglio_valido` (caratteri vietati Excel `\ * ? : [ ]` → `_`, max 31 char, fallback "ARGOMENTO"). Prima un nome del tipo `../../fuga.xlsx` scriveva fuori da `export/`.
2. **`scrivi_sotto` guardato** (`app/parser.py`): ora ritorna il numero di valori scritti e NON sovrascrive mai un'etichetta in grassetto sottostante (si ferma prima); aggiunto `wb.close()`.
3. **Scritture atomiche** config/state (`app/app.py`): `_scrivi_config` e `_scrivi_state` scrivono su `.tmp` e fanno `os.replace` (niente file corrotti in caso di crash a metà).
4. **`api_modella` validato**: risp. 400 su riga fuori range (1..1.048.576), colonna fuori range (1..16.384) o colonna non alfanumerica; `api_aggiungi` accetta solo `riga_etichetta` numerica.
5. **Errori puliti**: `_errore_generico` (handler globale) → JSON `{"errore": "Errore interno del server"}` con status 500 + traceback su console; `_estraai_tutti` tollera `riga_etichetta` non valida (restituisce `errore` per quell'argomento).
6. **Suite pytest** (`tests/`, isolata su tmp_path/venv): 26 test verdi (`app/.venv/bin/python -m pytest tests -q` → `26 passed`) — coprono parser (trova/estrai/scrivi/crea voce) e API (ciclo aggiungi→compila→elabora×2 con "nessuna variazione", export traversal/caratteri vietati, download 404, modella, rimuovi).
7. **E2E live verificato** su server 8010 (reloader attivo): VELOCITA' e CO2 compilati al posto giusto (B4-B7, E22-E24), vecchi valori svuotati, seconda `api/elabora` → "nessuna variazione", export con path traversal bloccato in `export/`. Mockups/config/state ripristinati dopo il test.

## 📌 Prossimi passi
1. Validare la parser su file Excel reali dell'utente (euristica titoli/grassetti: cella sotto in grassetto, font size > 11.5). Aggiungere test a `tests/` per eventuali nuovi regole emerse.
2. Attivare n8n (`docker compose --profile auto up -d`) e agganciare `/api/elabora` come bridge di automazione.
3. (Futuro) Integrare notifiche email.
4. Push del progetto su GitHub `Tox-oss/Progetto-Overlord` (branch main) con `.gitignore` per `.venv` e temporanei.
5. Documentare nel README il comando di test: `app/.venv/bin/python -m pytest tests -q` (venv già esistente con pytest 9.1.1; `tests/conftest.py` genera mockup isolati in tmp_path).