# Pulizia root dopo integrazione — 2026-10-06

Richiesta utente: eliminare cartelle/file della root non più necessari dopo la migrazione e il merge locale su dev.

## Rimozioni motivate

| Directory root | File | Evidenza |
|---|---:|---|
| lib | 445 | JavaScript, dichiarazioni, mappe e asset del vecchio build accoppiato; entrypoint library ora nel workspace packages |
| dist | 30 | Bundle debug e chunk del vecchio progetto SPFx; output attivi sotto apps/test |
| release | 31 | Asset, manifest e audit generati dal vecchio progetto; output attivi sotto apps/test |
| temp | 1 | Vecchio temp/build/manifests.js; serve usa quello dell’app |
| src | 1 | Solo modulo Sass generato .scss.ts, nessun sorgente di prodotto |

Tutti i 508 file erano ignorati da Git, nessuno tracked/unignored, nessun symlink. Ricerca di riferimenti statici/dinamici in script, config, manifest, test e VSCode: i percorsi attivi sono workspace-relative sotto packages/apps. Il banner assets/banner.png è ancora usato dal README; node_modules, artifacts, docs, scripts, configurazioni npm/VSCode e sorgenti workspace sono conservati.

## Backup e verifiche

Backup completo verificato per dimensioni e SHA256 prima della rimozione: `/private/tmp/spfx-legacy-root-backup-62qxde84/legacy-root.tgz`. Totale 13410937 byte. Manifest persistente: evidence/root-cleanup-manifest.json. È un backup temporaneo dei vecchi output generati, non un artefatto del nuovo monorepo.

Dopo la rimozione: npm run build → npm run verify → npm run package:solution, tutti exit0. Test90/90, API82, tipi/lint e verificatori PASS; soluzione prodotta nell’app. Le cinque directory root non sono state rigenerate.

Revisore indipendente root_cleanup_review (GPT-6.1 Sol) PASS: verificati integralmente tutti i 508 file del backup, log reali, assenza delle directory dopo build/ship, linking workspace e 116 sorgenti apps/packages identici a HEAD. Nessuna mutazione dal reviewer.

Nessun sorgente/API/dipendenza di prodotto modificato da questa pulizia. Verifica tenant resta BLOCKED/NOT RUN.
