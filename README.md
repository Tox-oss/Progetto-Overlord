# Overlord

Strumento che monitora diverse cartelle di file Excel sorgente, estrae gli **argomenti** (etichette in grassetto con valori sotto) e compila un'**unica scheda** Excel di destinazione, scrivendo i valori nelle posizioni giuste in base a dove sta l'etichetta.

## Com'è fatto

- `files/source/` — file Excel sorgente (leggere). Struttura: colonna verticale, etichetta **in grassetto**, valori consecutivi sotto; una riga vuota singola = buco dati (saltata), **due righe vuote consecutive** = fine del blocco.
- `files/dest/` — scheda finale: modulo libero con etichette in grassetto **in posizioni sparse**; i valori si scrivono **sotto** l'etichetta con lo stesso testo.
- `app/` — web app Flask (porta `8010`): tendine per scegliere file, argomento e scheda, configurazione multi-sorgente, "Modella scheda" per creare voci in grassetto vuote, esportazione.
- `n8n/` — docker-compose con `n8n` (profilo `auto`, attivabile in seguito per l'automazione) e il servizio `selettore` (la web app in Docker).
- `tools/make_mockups.py` — rigenera i file di esempio.

## Come eseguire (test)

```bash
cd "~/Scrivania/Lavori/Progetto Overlord"

# 1. Avvia la web app (richiede Python 3.12 e il venv)
app/.venv/bin/python app/app.py
# oppure con Docker:
#   (cd n8n && docker compose up -d selettore)
```

Poi apri **http://localhost:8010**.

## Come testarlo passo passo (mockup inclusi)

I mockup sono già pronti: `impianto_mock.xlsx` e `ambiente_mock.xlsx` in `files/source/`, `scheda_mock.xlsx` in `files/dest/`.

1. **Aggiungi un argomento dal sorgente**
   - Sezione 1: scegli `impianto_mock.xlsx` → tendina argomento → `VELOCITA'` (anteprima valori sotto) → **Aggiungi alla configurazione**.
2. **Aggiungi argomenti da altri sorgenti** (test multi-sorgente / "la dozzina di Excel")
   - Ripeti: `impianto_mock.xlsx` → `TEMPERATURA`, poi `ambiente_mock.xlsx` → `UMIDITA'`.
3. **Imposta la scheda finale**
   - Sezione 2 → tendina → `scheda_mock.xlsx`.
4. **Compila**
   - Sezione 2 → **Compila modulo**. Esito: ogni argomento scrive i valori **sotto** l'etichetta con lo stesso testo nella scheda (VELOCITA'→B4..., TEMPERATURA→C13..., UMIDITA'→B20...). Controlla con un doppio click su `scheda_mock.xlsx`.
5. **Overwrite in-place**
   - Sotto `TEMPERATURA` (C12) nel mockup ci sono valori vecchi (99, 98): la compilazione li **svuota e riscrive** con quelli nuovi (22, 23, 21, 22).
6. **Crea una voce nuova (Modella scheda)**
   - Sezione 3: scheda `scheda_mock.xlsx`, voce es. `FOTOMETRIA`, riga `28`, colonna `B` → **Crea voce**; la vedi nella lista con posizione `B28`. Da ora può essere usata come destinazione per un argomento omonimo.
7. **Dedup (bridge per l'automazione)**
   - Riavvio i mock con `app/.venv/bin/python tools/make_mockups.py`, poi chiama l'API:
     ```bash
     curl -X POST http://localhost:8010/api/elabora
     ```
     Prima volta: compila e salva lo stato. Seconda volta: risponde `"nessuna variazione"`.
8. **Esporta con nome**
   - Sezione 2 → **Esporta con nome…** → scarica un xlsx con un foglio per ogni argomento configurato.

## Ripristino dei mock

```bash
app/.venv/bin/python tools/make_mockups.py
```

Resetta anche `files/config/selezione.json` (`destinazione`/`argomenti`) e `files/state/last_value.json` se vuoi partire da zero.

## API principali

| Metodo | Percorso | Scopo |
|---|---|---|
| GET | `/api/sorgenti`, `/api/destinazioni` | elenco file |
| GET | `/api/argomenti?file=` | argomenti in grassetto del sorgente |
| GET | `/api/scheda?file=` | voci in grassetto della scheda |
| POST | `/api/aggiungi`, `/api/rimuovi` | configurazione argomenti |
| POST | `/api/modella` | crea voce in grassetto nella scheda |
| POST | `/api/compila` | compila tutti gli argomenti nella scheda |
| POST | `/api/export` | esporta con nome |
| POST | `/api/elabora` | bridge n8n (estrai + dedup + compila) |

## Prossimi passi

- Validare la parser su file reali dell'utente (euristiche titolo/grassetto).
- Attivare l'automazione: `(cd n8n && docker compose --profile auto up -d)` e collegare il workflow n8n a `/api/elabora`.
- Notifiche email.