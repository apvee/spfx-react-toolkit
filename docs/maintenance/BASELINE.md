# Stato iniziale — 2026-10-06

- HEAD: `7e0653622e8523d782ab9cf5dbbd185b03ace93b`, checkout originale `dev`, stato Git pulito. Nessun AGENTS.md trovato nel repository o negli antenati originali.
- Worktree Codex isolato: `/Users/fabiofranzini/.codex/worktrees/spfx-toolkit-monorepo/spfx-react-toolkit`; originale preservato. Scritture worktree autorizzate via sandbox escalation.
- Vincoli: API/nome/versione pacchetto conservati; nessun publish, deploy, push, merge, aggiornamento major o tool senza necessità. API pubbliche non rimosse per assenza usi interni.
- Ambiente osservato: Node 22.20.0, npm 11.19.1; SPFx build/runtime/eslint 1.21.1, TS 5.3.3, React/DOM 17.0.1, PnP 4.17.0, Fluent8 8.125.0. Engine dichiarato >=22.14 <23.
- Struttura: root pubblicabile con `src/{core,hooks,services,helpers,utils,webparts,extensions}`, Gulp/config/sharepoint/teams/assets, tre script di verifica, documentazione API. Nessuna CI in HEAD.
- Build originale: `npm run build` PASS (bundle ~6.0s, totale 7.6s). `npm test` PASS (~6.3s task, 8.6s totale), ma esegue sass/lint/tsc/webpack, senza assertion test: copertura comportamentale non dimostrata. Eseguiti con dipendenze preesistenti, installazione pulita ancora da verificare.
- `npm run verify:examples`: PASS, 40 hook e 4 provider nel registry (copertura dichiarativa, non esecuzione tenant). `verify:runtime-store`: PASS, assertion su stato/subscription. `verify:public-docs`: PASS, nomi helper/service e riferimenti obsoleti.
- I comandi build/test sono stati lanciati contemporaneamente sul baseline; tempi non sono benchmark affidabili. Misure prestazionali successive saranno sequenziali. Log in `docs/maintenance/evidence/baseline-*.txt`.
- API e dichiarazioni baseline saranno confrontate con compilazione finale, inclusi deep path `lib/{core,hooks,services,helpers,utils}`. Main ESNext `lib/index.js`, types `lib/index.d.ts`, senza exports map.
- SharePoint/Graph/tenant: nessun accesso fornito. Build non prova funzionamento tenant; preparare matrice e procedura riproducibile.

## Stato fasi

1 Brainstorming/design/piano: IN PROGRESS, tre agenti indipendenti; revisione nuova richiesta prima prodotto.
2 Baseline/criteri: IN PROGRESS, revisione indipendente ancora necessaria.
3 Migrazione: PENDING. 4 Bug: PENDING. 5 Ottimizzazioni: PENDING. 6 Pulizia: PENDING. 7 Docs: PENDING. 8 Integrità: PENDING.

- Installazione baseline worktree: `npm ci --ignore-scripts --no-audit --no-fund` PASS (2425 packages, 9s), lifecycle scripts non esercitati in questo passo; finale userà npm ci completo. Pack baseline: 331 file, 176648 bytes compressi, 817065 bytes estratti.

- Prepublish baseline sequenziale PASS: clean → bundle → gulp test (secondo bundle) → tre verifier, prova `evidence/baseline-release.txt`. Dichiarazioni baseline persistite con SHA256 in `evidence/api-baseline.json`; log .txt per non essere esclusi da .gitignore.
