# Comando serve nella root — 2026-10-06

Richiesta utente: aggiungere al package.json principale il comando npm di avvio SPFx.

Script aggiunto: `npm run build:library && npm run serve --workspace @apvee/spfx-react-toolkit-test --`. Compila la libreria prima di avviare Gulp nell’app; un errore di compilazione impedisce l’avvio. Il separatore finale inoltra gli argomenti, ad esempio `npm run serve -- --config=spFxReactToolkitTest`.

README e guide di sviluppo/tenant aggiornate. Nessuna dipendenza, lockfile, API o implementazione runtime modificata.

Verifiche: `npm run serve -- --help` exit0 (build library + gulp serve --help), `verify:public-docs` exit0 e diff whitespace pulito. Non avviati server/browser o trust certificato; la verifica tenant resta non eseguita. Nessun test nuovo che replichi questa configurazione.

Reviewer indipendente root_serve_review, GPT-6.1 Sol: PASS, dispatch/forwarding senza ricorsione, documentazione e configurazione Gulp coerenti. Log persistenti in evidence/root-serve-*.txt.
