# Contesto - Overlord (ex "Progetto ExEL")

## 🎯 Obiettivo
Strumento che monitora dozzine di file Excel sorgente, estrae "argomenti" (etichette in grassetto con valori sotto) e compila un unico file Excel di destinazione ("scheda") in posizioni specifiche. Ruolo principale: rilevare modifiche ai file, prelevare valori selezionati e immetterli nella scheda da compilare. Futuro: automazione con n8n self-hosted e notifiche email.

## 📍 Stato corrente
App e API funzionanti e testate end-to-end (multi-sorgente → compila in posizioni giuste; dedup ok; modella scheda crea voci; export ok). Bug sweep elemento-per-elemento del 25/9 completato (set 25 + 25-bis: parser, API, UI, infra; set 26: hygiene parser/test): **75 test verdi** + smoke E2E 25/25 + build Docker verificata. Server live su **8010** gira col codice del set 25 (fix 25-bis/26 nel repo, attendono commit). n8n NON ancora attivato. Email non integrata. Prossimo passo (con l'utente): **validare la parser su file reali**, **committare i tre set**, **riavviare il server**, **attivare n8n**.

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

## ✅ Correzioni applicate (bug sweep elemento-per-elemento, set 25)
Sweep in sola lettura su copia isolata `/tmp/overlord-sweep` (metodologia approvata: scope "tutto, inclusa infra"; multi-foglio "mantieni active ma testalo"; riproduzioni delle anomalie PRIMA di toccare il progetto; fix+test; commit solo su richiesta). Invarianti parser (1 vuota = buco, 2 vuote = fine blocco, scrittura sotto etichetta omonima) preservati e ricoperti da test.

**Class A — parser.py (13 funzioni esaminate)**
1. **A1 `colonna_a_numero`**: `""`/`"  "` restituivano 0 silenzioso → ora `ValueError` (posizione non valida). Test.
2. **A2 `trova_cella_etichetta`**: riconfermato *case-sensitive*; etichette duplicate → prima occorrenza nel foglio. Test che fissa il comportamento.
3. **A3 `crea_voce`**: ora valida riga/colonna ≥ 1, ripulisce il testo dagli spazi e **rifiuta con `ValueError` le celle dentro intervalli uniti** (anchor e `MergedCell`) — prima poteva scrivere "a metà" di un'unione. Test (anche via API: modella su cella unita → 400).
4. **A4 `scrivi_sotto`/`_spazio_libero`**: un merge in colonna LATERALE non blocca più la colonna dell'etichetta (test); `_spazio_libero` accetta `soglia` (tetto di conta) e non guarda più `ws.max_row` (Excel crea celle al volo; un'etichetta sotto con `max_row` piccolo azzerava lo spazio).
5. **A5 multi-foglio**: documento "LIMITE DOCUMENTATO" nel docstring — si legge/scrive solo il primo foglio attivo; test che fissa la lettura del solo primo foglio. Vincolo confermato dall'utente.
6. **A6/A7/A8**: `estrai_colonna` oltre max_row → nessun crash ([]); booleani/Decimal/int conservati; `_spazio_libero` con soglia=0 e con due vuote consecutive. Test.

**Class B — app.py (API, 13 rotte verificate)**
7. **B1 handler errori**: `_errore_generico` (solo Exception) inghiottiva 404/405/400 trasformandoli in 500 → aggiunto `_errore_http` (kiwi `HTTPException`) che risponde JSON con status e `errore` corretti. Riprodotto pre-fix (5 casi → 500) e verificato post-fix (rotta inesistente→404, GET su rotta POST-only→405, JSON malformato→400).
8. **B2 guardia `_payload_dict`** su `aggiungi`/`rimuovi`/`destinazione`/`modella`/`export`: payload non-dict (lista/stringa) → **400** invece di 500.
9. **B3 config senza `riga_etichetta`** (modificata a mano): `_estraai_tutti` → errore per-voce "riga_etichetta non valida", compila → 400, niente 500. Test.
10. **B4 export con argomento non estraibile** → **400** "Nessun argomento estraibile" + `esiti` (strategia tutto-o-niente coerente con compila). Test + E2E.
11. **B5 `_nome_foglio_valido`**: apostrofi ai bordi rimossi ("VEL_OCITA'" → "VEL_OCITA"), max 31 char confermato; test dedup aggiornato.
12. **B6 `/api/download`**: test di conferma — path url-encoded `..%2F...` → 404 (nessuna lettura fuori da `export/`).
13. **B7 `api_elabora`**: niente più `str(exc)` (info-leak) → messaggio generico + `traceback.print_exc()` su console.
14. **B8 tutto-o-niente non transazionale**: Documentato in README "Limiti noti" — la pre-validazione è trasversale (se un solo argomento è errato nulla si scrive); dopo la validazione le celle si scrivono una a una, quindi un'interruzione a metà ciclo può lasciare scritture parziali.
15. **B9 `api_modella`**: testo ripulito dagli spazi (strip, coerente con `crea_voce`). Test + E2E.

**Class C — index.html (UI)**
16. **C1 `btnModella`**: disabilitato all'avvio (la `carica()` ora chiama `caricaVoci()`), non più "acceso" di default.
17. **C2 `argomentoScelto`**: niente più regex fragile `/riga (\d+)/` sul testo della tendina; la riga viaggia in `data-riga` dell'`<option>` (un'etichetta con "riga" nel nome non rompe più la selezione).
18. **C3 rete**: tutti i fetch interattivi passano da `richiesta()` (try/catch → "Errore di rete: impossibile contattare il server").
19. **C4 doppio click**: guardia `conGuardia()` (un'azione alla volta su Aggiungi/Rimuovi/Compila/Export/Modella/Destinazione).
20. **C5 `esc()`**: riconfermato sui flussi di errore 400. Validato con `node --check`.

**Class D — infrastruttura**
21. **`.dockerignore`** (nuovo): prima `COPY app/ /app/` copiava `app/.venv` (centinaia di MB) e l'intero albero (files/, tests/, contexto/, .git) nel build context → ora esclusi. Immagine risultante 252 MB con `/app` pulito (verificato montando il container).
22. **`app/requirements.txt` pinnato**: `flask==3.1.3`, `openpyxl==3.1.5` (coincidono col venv).
23. **`README.md`**: nuova sezione "Limiti noti" (solo primo foglio; scrittura non transazionale; dedup per intero; celle unite; nomi foglio 31 char + apostrofi; `riga_etichetta` mancante).
24. **`n8n/docker-compose.yml`**: `healthcheck` aggiunto al servizio `selettore`.

**Riproduzioni e verifiche**
- Riproduzione pre-fix B1/B2 su copia `/tmp/overlord-sweep` (5 casi → 500), post-fix corretti.
- Suite pytest: **63 test verdi** (18 in più del QA del 24/9).
- Smoke E2E su `/tmp/overlord-smoke` (porta 8099): **25/25** (+11 controlli: 404/405/400, payload non-dict, body non-JSON, export con sorgente errata, download `..%2F`, modella su cella unita → 400, modella con strip testo).
- Build Docker di `Dockerfile.selettore` eseguita e container avviato su porta di test 8098 (serve su `/`, `/files` montato, `/app` senza `.venv`); immagine di test rimossa.
- Server live su **8010 riavviato col codice nuovo** (PID rinnovato): root→200, rotta inesistente→404, GET su compila→405, payload non-dict→400, compila senza destinazione→400.

## ✅ Correzioni applicate (bug sweep elemento-per-elemento, set 25-bis)

Sweep di completamento dopo il set 25 (approvato dall'utente: "Sì, tutti A1-A2, B1-B7, C1"). Tutte le anomalie riprodotte su copia isolata PRIMA del fix, poi corrette con test dedicati.

- **A1** (`parser.trova_cella_etichetta`): ora scarta i TITOLI come `trova_argomenti` (font > 11.5 oppure cella sottostante in grassetto) → prima matchava un header grande e compilava sotto al titolo.
- **A2** (`parser.colonna_a_numero`): colonne 0/negative ora danno `ValueError`.
- **B1** (`app._estraai_tutti`): config `argomenti` come dict / lista di stringhe → niente 500, errore per-voce "voce di configurazione non valida".
- **B2** (avvio `app.py`): ora crea anche `config/` e `state/` (prima solo source/dest/export → primo aggiungi = 500 su deployment fresco).
- **B3** (`ap_estraai_tutti`): sorgente xlsx corrotto/illeggibile (es. non-zip) → errore per-voce "sorgente non leggibile" su compila; elabora → `ok:false` con motivo.
- **B4** (`/api/aggiungi`): `riga_etichetta` frazionaria (3.5) → 400; float intero (3.0) resta accettato.
- **B5** (`/api/compila`): zero argomenti configurati → 400 "Nessun argomento configurato".
- **B6** (`/api/export`): suffisso di dedup dei fogli duplicati contenuto in 31 char (base troncata prima di `_N`), niente più nomi >31.
- **B7** (`app._nome_valido`): `..` interno al nome (es. `rel..dati.xlsx`) ora accettato e selezionabile; traversal reale (separatori/`.`, `..`, root) ancora bloccato (test 404 di conferma).
- **C1** (`index.html`): fetch di `carica`/`aggiornaConfig`/`caricaArgomenti`/`caricaVoci`/`mostraAnteprimaCorrente` ora passano da `richiesta()` con gestione errori (niente `voci.length` su null / elaborazioni mute); `rimuoviArgomento` protetto da indice fuori range.

**Verifiche**
- Riproduzioni pre-fix su `/tmp/overlord-sweep2` (FR2/B1, FR4/B2, FR7/B3, FR9/A1, FR10/A2, allineamenti B4/B5/B6/B7).
- Suite pytest: **73 test verdi** (10 in più).
- Smoke E2E su `/tmp/overlord-smoke` (porta 8099, codice aggiornato): **25/25**.
- Token di rete extra: `node --check` sul blocco `<script>` di `index.html` OK.

## ✅ Correzioni applicate (bug sweep, set 26 — hygiene parser/test)

Report esterno analizzato punto-per-punto; NAD (nessuna anomalia) per il "BUG 5" dopo verifica sul codice.

- **BUG 1/2** (`parser.py`): `wb.close()` prima di `wb.save()` in `scrivi_sotto` e `crea_voce` → ordine invertito. Swappati a `save()→close()`. Funzionava per caso (openpyxl non distrugge i dati in memoria): fix formale/robustezza. Unica occorrenza nel repo.
- **BUG 3** (`app.py`): docstring "Flasck" → "Flask".
- **BUG 4** (`test_api.py`): rinomina `test_compila_ersenza_destinazione_400` → `test_compila_senza_destinazione_400`.
- **BUG 6** (`test_api.py`): 3× `write_text(json.dumps(...))` senza encoding → `ensure_ascii=False` + `encoding="utf-8"` (come fa l'app); nuovo test round-trip config UTF-8 accentata (salvataggio + `/api/config`).
- **BUG 5 — analizzato, dichiarato comportamento corretto (NAD)**: il report sosteneva che `_spazio_libero` "sovrastima" lo spazio perché non si ferma alle 2 righe vuote e che `scrivi_sotto` scriverebbe solo 1 valore. Verifica sul codice: (1) `_spazio_libero` e il write-loop di `scrivi_sotto` si fermano agli STESSI ostacoli (grassetto/merge) → nessuna sovrastima rispetto alla capacità reale; (2) la regola "2 vuote = fine blocco" governa i SORGENTI (`estrai_colonna`), il DEST non viene mai ri-letto a blocchi (le voci si trovano per etichetta con `trova_cella_etichetta`) → il rischio di "merging dei blocchi" non è raggiungibile; (3) applicare le 2 vuote al DEST regredirebbe la crescita in colonne vuote. Scelta utente: **Proposta A — documentare e pinnare**. Aggiornata docstring di `_spazio_libero` (DEST fino all'ostacolo; 2 vuote = regola sorgenti) + nuovo test `test_spazio_libero_oltre_due_vuote_davanti_a_bold`.

**Verifiche**
- Suite pytest: **75 test verdi** (73 + 2 nuovi: round-trip UTF-8 e pin BUG 5).
- Smoke E2E su `/tmp/overlord-smoke` (porta 8099): **25/25**.

## 📌 Prossimi passi
1. **Committare bug sweep set 25 + 25-bis** (modifiche di QA + parser/API/UI/infra non ancora committate; su conferma dell'utente) ed eventuale push su `Tox-oss/Progetto-Overlord`.
2. Validare la parser su file Excel reali dell'utente (euristica titoli/grassetti: cella sotto in grassetto, font size > 11.5 — ora applicata anche ai sorgenti). Aggiungere test a `tests/` per eventuali nuove regole emerse.
3. Attivare n8n (`docker compose --profile auto up -d`) e agganciare `/api/elabora` come bridge di automazione.
4. (Futuro) Integrare notifiche email.
5. Push del progetto su GitHub `Tox-oss/Progetto-Overlord` (branch main) con `.gitignore` e `.dockerignore`.
6. Documentare nel README il comando di test: `app/.venv/bin/python -m pytest tests -q` (venv già esistente; `tests/conftest.py` genera mockup isolati in tmp_path).