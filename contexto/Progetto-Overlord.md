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

## ✅ Correzioni applicate
- Nessuna correzione registrata in questo ciclo: stato di sviluppo funzionante, non di correzione.

## 📌 Prossimi passi
1. Validare la parser su file Excel reali dell'utente (euristica titoli/grassetti: cella sotto in grassetto, font size > 11.5).
2. Attivare n8n (`docker compose --profile auto up -d`) e agganciare `/api/elabora` come bridge di automazione.
3. (Futuro) Integrare notifiche email.
4. Push del progetto su GitHub `Tox-oss/Progetto-Overlord` (branch main) con `.gitignore` per `.venv` e temporanei.