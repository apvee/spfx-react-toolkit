# Monorepo Implementation Plan

> Esecuzione con orchestratore e subagenti GPT-6.1 Sol. Skill: subagent-driven-development per task indipendenti; integrazione/config/lock centralizzati per richiesta utente.

**Goal:** libreria indipendente + app SPFx con API pubblica, bugfix testati, pulizia e docs coerenti.
**Architecture:** npm workspaces; tsc library ESNext, Gulp app1.21.1.
**Tech:** Node22, npm11, TS5.3.3, React17.0.1, SPFx1.21.1.
**Spec:** [DESIGN.md](DESIGN.md).

## Vincoli globali
- Conservare nome/versione/peer/export/deep path e semantica salvo bugfix documentati. Nessuna API rimossa.
- No publish/deploy/push/merge, major o strumenti immotivati; checkout originale preservato.
- Modelli subagenti sempre gpt-6.1-sol esplicito, max3 worker; niente scritture concorrenti stessi file.
- Revisori diversi autori; risultati basati stato corrente e comandi, non solo report. Finding bloccanti corretti/rivalidati prima PASS.

## Review focus
Concorrenza/risposte invertite; unmount/identity provider; side effect PnP e cache expiry; deep import/distribuzione ESNext; hoist che maschera peer mancanti. Test/consumer e review dei task relativi coprono questi rischi.

## Task e dipendenze
### T1 Design e baseline (F1/F2) — orchestratore, reviewer nuovo
- [x] Ispezione pulita, versioni, scripts, tre brainstorming.
- [x] Design/alternative/piano e baseline con log.
- [x] Review indipendente design+baseline/criteri PASS; rilievi corretti e re-review prima prodotto.

### T2 Migrazione (F3) — orchestratore integrazione, reviewer nuovo
Files: root package/lock/scripts/config; packages/library; apps/test.
- [x] Check preliminare confine workspace deve fallire sul layout originale.
- [x] Spostare sorgenti e config pertinenti preservando identità; aggiornare import public e tsconfig.
- [x] Package dev versions già lockate, root orchestration library→app, verifier path e npm test reale.
- [x] Install lock/ci, library build, verifier, app dev/ship/package, pack allowlist.
- [x] Consumer tarball compilato/bundled fuori workspace; confronto d.ts baseline.
- [x] Review struttura/export/types/boundaries, fix+re-review.

### T3 Bug runtime e servizi (F4) — subagenti con file disgiunti, reviewer diverso
Files: packages/library/src/core e hooks per provider/runtime; hooks storage/async/KV/search/tenant/OneDrive/List; service PnP cache; CrossSitePermissions/UserPhoto/Performance aggiunti dopo riproduzione audit.
- [x] Convertire riproduzioni in suite persistenti, eseguire RED contro baseline.
- [x] Fix minimi per B01–B09 confermati (inclusi storage inline default, provider properties/displayMode e identity races) con invarianti firme pubbliche; errori e casi limite.
- [x] GREEN suite reale + tsc/lint; registry scenario suggerimenti se bug confermato.
- [x] Review indipendente ogni gruppo, fix e re-review.

### T4 Ottimizzazione/pulizia (F5/F6) — dopo T2, disgiunto T3
Files: scripts/package configs.
- [x] Misurare baseline sequenziale lib packaging; confrontare tsc/release e contenuto pack.
- [x] Elimina duplicazione build/test nominale; verifier temp unico con due esecuzioni parallele.
- [x] Prova riferimenti prima rimozioni dev ajv/lodash/fabric/decorator/config residua; installazione+build.
- [x] Review beneficio/rimozioni+fix/re-review.

### T5 Documentazione (F7) — dopo stabilizzazione T2/T3/T4, worker docs
Files: README, package README, docs API/introduction/index e docs/maintenance; solo docs worker, ledger orchestratore.
- [x] Struttura/requisiti/dev/build/serve/test/pack/migrazione, comportamento bug corretto, limiti peer/tenant.
- [x] Aggiornare/verificare link, esempi/API/export e guide tenant scenari significativi.
- [x] Review indipendente istruzioni reali+fix/re-review.

### T6 Integrità finale (F8) — orchestratore controlli, due reviewer nuovi
- [x] Clean copy tracked/current source + npm ci (script effettivi), npm ls runtime.
- [x] npm run build:library, npm test, lint/typecheck, tre verifier e API/pack consumer.
- [x] npm run build:app, bundle:ship, package:solution; manifest XML e scenari.
- [x] Revisioni trasversali library/runtime e monorepo/app/docs/CI; fix all findings e revalidate definitive tree.
- [x] Matrice conclusiva con log e limiti, nessun PASS per test tenant assente. Stato goal secondo criteri e regole blocco.

## Ownership / preflight
T2 config+lock centrale → T3 soltanto sorgenti/test assegnati; T4 verifier/package integrati centralmente; T5 riceve interfacce/comandi definitivi. T2/T3 condividono directory ma non file dopo spostamento; T3 parte solo dopo T2 path stabili. T5 non tocca maintenance ledger. T6 read-only. Tutte le task coerenti col design, nessuna API nuova.

- [ ] Verifica reale autenticata tenant: BLOCKED/NOT RUN; nessun accesso o risultato manuale disponibile, goal non completo.
