# Contesto - Overlord (ex "Progetto ExEL")

## 🎯 Obiettivo
Strumento che monitora dozzine di file Excel sorgente, estrae "argomenti" (etichette in grassetto con valori sotto) e compila un unico file Excel di destinazione ("scheda") in posizioni specifiche. Ruolo principale: rilevare modifiche ai file, prelevare valori selezionati e immetterli nella scheda da compilare. Futuro: automazione con n8n self-hosted e notifiche email.

## 📍 Stato corrente
App e API funzionanti e testate end-to-end (multi-sorgente → compila in posizioni giuste; dedup ok; modella scheda crea voci; export ok). Bug sweep QA del 24/9 completato: **45 test verdi** + smoke E2E 14/14, server live su **8010** da riavviare per caricare i fix. n8n NON ancora attivato. Email non integrata. Prossimo passo (con l'utente): **validare la parser su file reali**, **riavviare il server** e **attivare n8n**.

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
- API disponibili: GET `/api/sorgenti`, `/api/destinazioni`, `/api/argomenti?file=`, `/api/scheda?file=`, `/api/config`; POST `/api/aggiungi`, `/api/rimuovi`, `/api/destinazione`, `/api/modella`, `/api/compila` (**400** se non completa), `/api/export`, `/api/elabora` (bridge automazione: rilegge config, estrae tutti gli argomenti, confronta con state, se cambiati compila tutto-o-niente e aggiorna state; risponde "nessuna variazione" se invariato); GET `/api/download/<nome>`. Accessi file in `/api/argomenti`/`/api/scheda`/`/api/aggiungi`/`/api/modella` bloccati fuori dalle cartelle di lavoro (404); `/api/destinazione` rifiuta nomi pericolosi (400). Avvio: `app/.venv/bin/python app.py`; `debug` solo se `FLASK_DEBUG=1`. (fonte: `app/app.py`)
- Funzioni parser: `_cell_is_bold`, `_cell_is_merged`, `_cell_is_ostacolo`, `_cella_in_unione`, `carica_workbook`, `trova_argomenti` (colonna A, apre con data_only, scarta titoli font>11.5), `trova_argomenti_ovunque`, `estrai_colonna`, `trova_cella_etichetta`, `scrivi_sotto` (ferma la pulizia e la scrittura davanti a grassetto/celle unite), `_spazio_libero`, `crea_voce`, `colonna_a_numero`. (fonte: `app/parser.py`)
- **Porte**: web app su **8010** (8000 e 8001 occupate da altri progetti dell'utente: InventarioAuto ecc.). Confermata in ascolto su 8010. (verifica ambiente)
- **Regole parser da validare** con l'utente: scartare grassetti che NON sono argomenti (titoli) — euristica: cella sotto in grassetto → titolo; font size > 11.5 → titolo.
- **Compilazione**: testo etichetta uguale tra sorgente e scheda → scrivi valori sotto, svuotando il blocco vecchio (fonte: `app/parser.py` — `scrivi_sotto`).
- Al termine dei test: ripristinare mock/config/state con `tools/make_mockups.py` + reset JSON (config attualmente con destinazione null e un solo argomento CO2; state resettato). (fonte: `files/config/selezione.json`, `files/state/last_value.json`)
- n8n compose: selettore parte con `docker compose up -d`; n8n solo con `docker compose --profile auto up -d`. (fonte: `n8n/docker-compose.yml`)
- Git: push imminente su `https://github.com/Tox-oss/Progetto-Overlord.git` (branch main); `.gitignore` escluderà `.venv` e file temporanei.

## ⚠️ Criticità aperte
- [MEDIA] Euristiche parser (grassetto-titolo, font size > 11.5) testate sui mockup e applicate a sorgenti+scheda, ma NON ancora validate su file reali dell'utente — da verificare insieme.
- [INFO] n8n non avviato; email non integrata (fonte: contesto utente).
- [INFO] Config/state attualmente vuoti o parziali (residuo di test): destinazione null, un solo argomento (CO2), state resettato. Da ripristinare/popolare prima dell'uso reale.
- [INFO] Server live su 8010 da riavviare per caricare i fix del bug sweep (restart non ancora concordato con l'utente).
- [INFO] Nota formule: in `estrai_colonna` (data_only) le formule senza valore in cache risultano vuote. Su file Excel reali restituisce il valore calcolato.

## ✅ Correzioni applicate (bughunt esteso, set 24)
Applicati i principi delle istruzioni del cliente (`~/Scrivania/CLAUDE.md`, `DIRECTIVES.md`): verificare prima di dichiarare, autorità spec > piano > codice, non toccare gli invarianti del parser, isolamento dei test. Le regole del parser (1 riga vuota = buco; 2 vuote = fine blocco; etichetta grassetto sopra; scrittura sotto etichetta omonima) sono INVIOLATE e ora coperte da test.

1. **Export sicuro** (`app/app.py`): nuovo `_nome_export_sicuro` (solo basename, mai path traversal, caratteri illegali Windows `<>:"/\|?*` sostituiti, estensione `.xlsx` forzata) e `_nome_foglio_valido` (caratteri vietati Excel `\ * ? : [ ]` → `_`, max 31 char, fallback "ARGOMENTO"). Prima un nome del tipo `../../fuga.xlsx` scriveva fuori da `export/`.
2. **`scrivi_sotto` guardato** (`app/parser.py`): ora ritorna il numero di valori scritti e NON sovrascrive mai un'etichetta in grassetto sottostante (si ferma prima); aggiunto `wb.close()`.
3. **Scritture atomiche** config/state (`app/app.py`): `_scrivi_config` e `_scrivi_state` scrivono su `.tmp` e fanno `os.replace` (niente file corrotti in caso di crash a metà).
4. **`api_modella` validato**: risp. 400 su riga fuori range (1..1.048.576), colonna fuori range (1..16.384) o colonna non alfanumerica; `api_aggiungi` accetta solo `riga_etichetta` numerica.
5. **Errori puliti**: `_errore_generico` (handler globale) → JSON `{"errore": "Errore interno del server"}` con status 500 + traceback su console; `_estraai_tutti` tollera `riga_etichetta` non valida (restituisce `errore` per quell'argomento).
6. **Suite pytest** (`tests/`, isolata su tmp_path/venv): 26 test verdi (`app/.venv/bin/python -m pytest tests -q` → `26 passed`) — coprono parser (trova/estrai/scrivi/crea voce) e API (ciclo aggiungi→compila→elabora×2 con "nessuna variazione", export traversal/caratteri vietati, download 404, modella, rimuovi).
7. **E2E live verificato** su server 8010 (reloader attivo): VELOCITA' e CO2 compilati al posto giusto (B4-B7, E22-E24), vecchi valori svuotati, seconda `api/elabora` → "nessuna variazione", export con path traversal bloccato in `export/`. Mockups/config/state ripristinati dopo il test.

## ✅ Correzioni applicate (bug sweep QA, set 24)
Sweep in sola lettura (riproduzioni in `/tmp/overlord-qa`): 14 finding (3 Alta, 8 Media, 3 Bassa), tutti corretti e coperti da test. Decisioni utente (vincolanti): **compilazione tutto-o-niente per entrambi** `/api/compila` e `/api/elabora` (se anche un solo argomento è non estraibile/non trovato/spazio insufficiente → NON si scrive nulla); euristica "font size > 11.5 = titolo" applicata **anche ai sorgenti**. Invarianti parser (1 vuota = buco, 2 vuote = fine blocco, scrittura sotto etichetta omonima) preservati.

1. **Non cancella l'etichetta sotto** (P1, `parser.py`): la pulizia di `scrivi_sotto` si ferma PRIMA di una cella grassetto o unita (`_cell_is_ostacolo`).
2. **Spazio insufficiente = blocco totale** (P2/P7): pre-check `_spazio_libero(ws, riga, colonna, soglia=len(valori))`; se i valori non ci stanno nel blocco, tutto-o-niente: scheda intatta, stato NON salvato (niente "nessuna variazione" congelato), `api_compila` → **400** con `errore`+`esiti`.
3. **Sorgente mancante non svuota più la scheda** (P3): `_estraai_tutti` marca `errore` su quel argomento → compilazione annullata, scheda intatta.
4. **Path traversal bloccato** (P4): `_nome_valido` + `_risolvi_in` su `argomenti`/`scheda`/`aggiungi`/`modella` (404) e `destinazione` (400). Prima `?file=../segreto.xlsx` leggeva fuori da `SOURCE_DIR`.
5. **Date nel state non rompono lo stato** (P5): `_valore_seriale`/`_estratti_seriali` serializzano datetime/date/time/Decimal → `json.dumps` non esplode più; confronto "nessuna variazione" regge.
6. **Destinazione/file sparito = esito trasparente** (P6): `_compila_tutti` torna `completa:false` ("Destinazione non configurata" / "File di destinazione non trovato"), `api_elabora` risponde `ok:false` senza salvarlo.
7. **Etichetta assente in scheda = blocco totale** (P7): basta un `trovata:false` per fermare l'intera compilazione (decisione utente: all-or-nothing anche per `/api/compila`).
8. **Font size > 11.5 nei sorgenti** (P8): `trova_argomenti` ora scarta i titoli grandi come `trova_argomenti_ovunque` (simmetrico).
9. **Celle unite non esplodono** (P9): `_cella_in_unione` (via `ws.merged_cells`) tratta come ostacolo ANCHE l'anchor di un intervallo unito → niente AttributeError su `MergedCell` (i 94 cataloghi hanno celle unite).
10. **Formule lette come valore calcolato** (P10): `estrai_colonna` apre con `data_only=True`. Nota: formule senza valore in cache (file generati da openpyxl e mai aperti da Excel) risultano vuote → comportamento ammesso e documentato.
11. **Pulsante Compila coerente col server** (P11): abilitato solo con argomenti E scheda selezionati (`aggiornaBottoni`); il server risponde 400 e la UI mostra il motivo per ogni voce.
12. **XSS sanato** (P12): helper `esc()` su tutte le interpolazioni dinamiche in `index.html`.
13. **Nomi foglio export dedup** (P13): controllo case-insensitive, collisioni risolte con suffissi `CO2_2`, `CO2_3`.
14. **Debug non esposto in rete** (P14): `app.run(..., debug=...)` da env `FLASK_DEBUG` (default off); host `0.0.0.0` mantenuto (Docker).
- **Suite pytest**: ora **45 test verdi** (`app/.venv/bin/python -m pytest tests -q`), isolati su tmp_path/venv: nuovi casi parser (non-cancella-ettichetta, celle unite, titolo font grande, formula, spazio libero) e API (traversal su 5 rotte, sorgente mancante, destinazione mancante, datetime, export dedup, all-or-nothing, compila 400).
- **Smoke E2E live verificato** su copia in `/tmp/overlord-smoke` (porta 8099, mockup rigenerati): 14/14 check verdi — flusso aggiungi→destinazione→compila (scritti 3 in E22-E24), elabora "nessuna variazione", compila senza destinazione → 400, traversal negati, export ok, modella ok. Server a fine test terminato; mockup/config/state del progetto intatti.
- ⚠️ Il server live su **8010** gira ancora col codice vecchio: serve un **restart** per caricare i fix del bug sweep (non ancora richiesto dall'utente).

## 📌 Prossimi passi
1. Validare la parser su file Excel reali dell'utente (euristica titoli/grassetti: cella sotto in grassetto, font size > 11.5 — ora applicata anche ai sorgenti). Aggiungere test a `tests/` per eventuali nuove regole emerse.
2. Riavviare il server live su 8010 per caricare i fix del bug sweep (restart non ancora concordato).
3. Attivare n8n (`docker compose --profile auto up -d`) e agganciare `/api/elabora` come bridge di automazione.
4. (Futuro) Integrare notifiche email.
5. Push del progetto su GitHub `Tox-oss/Progetto-Overlord` (branch main) con `.gitignore` per `.venv` e temporanei.
6. Documentare nel README il comando di test: `app/.venv/bin/python -m pytest tests -q` (venv già esistente con pytest 9.1.1; `tests/conftest.py` genera mockup isolati in tmp_path).