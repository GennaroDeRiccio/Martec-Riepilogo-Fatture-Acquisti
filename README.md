# Martec Riepilogo Fatture Acquisti

Questa versione e pronta per una soluzione gratuita e condivisa:

- frontend statico pubblicabile su GitHub Pages;
- database cloud condiviso su Supabase Free;
- storage cloud dei PDF su Supabase Storage;
- sincronizzazione in tempo reale tra utenti tramite Supabase Realtime.

In pratica: chi apre la web app vede lo stesso archivio, puo caricare fatture e bonifici, e tutti lavorano sullo stesso database online.

## Architettura scelta

Soluzione gratuita consigliata e gia preparata nel progetto:

1. GitHub Pages per ospitare la web app statica dalla cartella `static/`
2. Supabase Free per:
   - tabella `records`
   - tabella `suppliers`
   - bucket `documents`
   - realtime sugli aggiornamenti

Non serve piu usare SQLite per la versione condivisa online.

## Cosa fa la web app

- Upload multiplo di PDF fatture e bonifici
- Estrazione dati direttamente nel browser
- Abbinamento automatico fattura/bonifico
- Salvataggio nel database cloud
- Archivio fornitori sincronizzato
- Import dello storico Excel
- Export in Excel con il template MARTEC
- Modifica manuale delle celle
- Aggiornamento condiviso in tempo reale

## File principali

- `static/index.html`: dashboard
- `static/suppliers.html`: archivio fornitori
- `static/app.js`: logica dashboard cloud
- `static/suppliers.js`: logica fornitori cloud
- `static/parser.js`: parser PDF/XLSX nel browser
- `static/cloud.js`: connessione Supabase
- `static/domain.js`: regole condivise di normalizzazione e matching
- `supabase/schema.sql`: schema e policy del database cloud
- `.github/workflows/deploy-pages.yml`: deploy automatico su GitHub Pages

## Setup Supabase

1. Crea un progetto su Supabase
2. Apri SQL Editor
3. Esegui tutto il contenuto di `supabase/schema.sql`
4. In Storage verifica che esista il bucket `documents`
5. In Database > Replication abilita `records` e `suppliers` per Realtime
6. Copia:
   - `Project URL`
   - `anon public key`

## Setup GitHub Pages

1. Carica il progetto in un repository GitHub
2. Mantieni il branch principale come `main`
3. In GitHub > Settings > Pages imposta GitHub Actions come source
4. Fai push del progetto: il workflow `.github/workflows/deploy-pages.yml` pubblichera automaticamente la cartella `static/`

## Primo avvio online

Quando apri la web app pubblicata:

1. comparira il box `Connessione cloud`
2. inserisci:
   - URL Supabase
   - anon key Supabase
3. premi `Salva connessione`

Da quel momento la web app:

- legge e scrive sul database cloud;
- salva i PDF nel bucket cloud;
- sincronizza i dati tra utenti.

## Note importanti

- GitHub Pages ospita solo il frontend: il database e lo storage stanno su Supabase.
- La chiave usata nel browser e la `anon key`, quindi non stai esponendo la service role key.
- Questa soluzione e gratuita, ma i limiti del piano Free di Supabase vanno tenuti presenti se aumentano molto file e utenti.
- Per massima sicurezza futura conviene aggiungere autenticazione utenti e policy piu restrittive. In questa prima versione ho lasciato policy aperte per permettere l'uso condiviso immediato.

## Preview locale

Se vuoi vedere la versione statica in locale puoi continuare a usare:

```bash
python3 server.py
```

Poi apri:

```text
http://127.0.0.1:8000
```

Il server locale qui serve solo per visualizzare i file statici. La persistenza condivisa online passa da Supabase.
