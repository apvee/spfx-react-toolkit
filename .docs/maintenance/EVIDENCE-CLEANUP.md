# Razionalizzazione documentazione ed evidenze — 2026-10-06

Richiesta utente: rimuovere ulteriori file non necessari tra test, documentazione e materiale di manutenzione.

## Decisione e perimetro

Audit indipendente maintenance_cleanup_audit (GPT-6.1 Sol): tutte le 6 suite e i 3 harness sono attivi, così tutti i 6 script e le 20 pagine pubbliche. Nessun test, controllo, API o pagina pubblica eliminato. Conservati gli 8 documenti del registro, i 4 rapporti autori con dettagli unici e le note recenti di integrazione/pulizia/serve.

Due baseline operative spostate da docs/maintenance/evidence a tests/fixtures; verify-api usa i nuovi percorsi. Tutti i byte e i rispettivi hash sono invariati. Non sono log eliminabili e non sono state rigenerate dal codice corrente.

Rimossi 22 file (251210 byte): log intermedi superati dai gate finali/re-review, un duplicato byte-identico store-concurrent-b e un harness che importava il worktree archiviato. Le due condizioni del vecchio harness sono test persistenti in async-hooks.test.cjs (CRUD/query e offset refetch).

## Prove conservate

- Baseline sequenziale originale, pack iniziale e misura build; risultati iniziali e limite dei test nominali registrati in BASELINE.
- RED dei bug/boundary/documentazione e dei sette casi PnPList; GREEN mirato e re-review indipendente, incluse le due condizioni aggiuntive options/laziness.
- Clean install/release/ship, gate library83, tarball consumer e pack; gate checkout dopo pulizia root90/API82.
- Compilazione snippet, albero runtime, source/artifact hash e manifest della pulizia root.
- Lo store fu verificato con due processi simultanei: entrambi produssero esattamente il contenuto conservato in store-concurrent-a.txt. Il duplicato B è stato confrontato dall'audit prima della rimozione.

Riferimenti aggiornati ai gate finali conservati. File eliminati e SHA256 nel manifest maintenance-cleanup-manifest.json. Backup completo verificato: `/private/tmp/spfx-maintenance-cleanup-9rgj22_9/removed-evidence.tgz`. I file precedentemente tracked sono anche recuperabili con `git show e59686c:docs/maintenance/evidence/NOMEFILE`.

Copertura e garanzie restano invariate. Verifiche finali: npm run verify exit0, 90/90 test, typecheck, lint, tre verifier e API82 PASS. Reviewer indipendente maintenance_cleanup_audit PASS: backup completo contro commit storico, fixture byte-identiche, sorgenti e test immutati, link integri. Prova tenant ancora BLOCKED/NOT RUN.
