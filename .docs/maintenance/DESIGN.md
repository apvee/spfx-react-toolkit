# Design: monorepo SPFx React Toolkit

## Obiettivo e vincoli
Separare libreria pubblicabile e applicazione dimostrativa, mantenendo `@apvee/spfx-react-toolkit@2.1.0`, API/export/deep path, React 17 e SPFx peer >=1.18 <2. Nessuna pubblicazione/deploy/push/merge. Utente autorizza decisioni motivate, revisione agentica del design e prosecuzione autonoma. Nessun AGENTS applicabile trovato.

## Alternative considerate
1. **npm workspaces**, root privata, `packages/spfx-react-toolkit` e `apps/spfx-react-toolkit-test`: un lock, linking nativo e ordine esplicito. Rischi hoisting e tsconfig Rush mitigati con dipendenze dichiarate, rootDir/outDir espliciti e consumer tarball esterno.
2. **Pacchetti autonomi con file: e npm --prefix**: isolamento cwd semplice ma due lock/installazioni e runtime duplicabili; linking non riproduce comunque tarball.
3. Libreria root e app annidata: meno movimenti ma gestione asimmetrica e due modalità.
Scelta 1: npm già in uso; due pacchetti non giustificano Nx/Turbo/Rush né package manager nuovo.

## Struttura e responsabilità
- Root privata: workspace elenco library/app, un lock v3, script sequenziali e verificatori/test di integrazione, docs condivise e CI minima.
- `packages/spfx-react-toolkit/src/{core,hooks,services,helpers,utils,index.ts}`; build `tsc`, ESNext/es2020 come baseline, `lib/index.js` e `lib/index.d.ts`. Nessuna exports map restrittiva; deep path e moduli internal necessari mantenuti. README/LICENSE locali includibili nel tarball. Niente sideEffects:false globale perché import PnP registrano estensioni.
- `apps/spfx-react-toolkit-test`: WebPart, Application Customizer, manifest, XML SharePoint, teams/config/Gulp/eslint/yo pertinenti (VSCode/MCP resta al root con launch path app); private, dipendenza toolkit `2.1.0` e peers runtime espliciti. Ogni import library usa `@apvee/spfx-react-toolkit`, nessun alias a source. ID manifest/solution conservati.
- Dev SPFx fissati a 1.21.1 già nel lock; TS5.3.3 React17.0.1. Peer library invariati. Toolchain Gulp conservata, nessun upgrade a Heft.
- Config Rush app deve correggere rootDir/outDir/include/typeRoots ereditati in presenza hoisting; libreria tsconfig esplicito riproduce opzioni baseline.

## Correttezza e test
Bug confermati iniziali: TenantKV get/list non awaitano dentro try/finally; useAsyncInvoke loading termina mentre altre chiamate pendono; Search mantiene vecchi refiners per closure; TenantProperty può applicare risultato di chiave vecchia; PnP cache timeout ignorato da Caching che usa expireFunc. Riproduzioni React reali confermano inoltre provider proprietà/displayMode e render loop storage default oggetto inline; promise controllate confermano OneDrive/PnPList stale identity. Coprire anche ricerca/paging e demo suggerimenti prima fix. Formato serializzazione tenant legacy conservato: envelope versionato richiede progetto dedicato e non viene introdotto senza necessità confermata. Nessuna modifica su sole ipotesi.
Test con Node built-in e TypeScript locale, mock limitati confine SPFx/browser; eventuale react-test-renderer17 dev giustificato per veri lifecycle React se necessario. Casi concurrency con deferred promise, unmount, identity change e error/retry. Non aggiungere API pubbliche.

## Ottimizzazione e pulizia
Build lib tsc elimina Sass/webpack demo dal packaging library e duplicazione gulp test senza Jest. Misure sequenziali prima/dopo su build/release e tarball (nessuna promessa bundle prima misura). Verifier runtime tmp unico/cleanup, compilatore locale. Rimuovere dipendenze dirette ajv/lodash/fabric solo dopo ricerca completa e clean build; no API eliminata. Conservare configurazione deploy incerta documentandola.

## Verifiche e compatibilità
Confronto dichiarazioni baseline salvate in /private/tmp/spfx-api-baseline e barrel source HEAD; deviazioni documentate solo bugfix. Root test esegue assertion reali; tre verifier conservano inventario e store. Nuovi checks confini/export/types/docs link/compilazione consumer. npm pack reale e consumer esterno con SPFx1.21.1/React17/PnP4, import root+deep, TypeScript e bundle reale: no require CommonJS per ESNext. npm ci pulito con script effettivi e npm ls per runtime. App bundle dev+ship e package-solution ship. CI ripete check statici/test/build/package/tarball.

## Limiti tenant
Provider FieldCustomizer/CommandSet nel registry hanno sola verifica export, non montaggio tenant. WebPart e ApplicationCustomizer sono scenari reali da preservare. SharePoint/Graph/Teams/auth/storage/network non provabili senza tenant/accessi: matrice e istruzioni riproducibili, esito NOT RUN/BLOCKED esplicito, nessuna simulazione. Compatibilità peer >=1.18 resta dichiarata; prova automatizzata esatta 1.21.1.

## Criteri di accettazione
Ogni fase ha produzione, revisore diverso, controlli, fix, re-review e PASS/FAIL/BLOCKED registrati. F1 alternative/design/piano revisionati; F2 baseline completa/test nominali esplicitati; F3 layout/confini/package/link/tarball build; F4 ogni bug riproduzione RED/GREEN e review; F5 beneficio misurato/documentato e review; F6 prove riferimenti/rimozioni+install/build+review; F7 comandi/link/API/esempi coerenti+review; F8 clean install/build/test/static/verifier/export/tarball/app/CI e revisioni trasversali senza finding bloccanti. Tenant indisponibile deve lasciare stato distinto e goal non completo se costituisce verifica obbligatoria residua.

## Fonti tecniche verificate
- https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility
- https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-toolchain
- https://docs.npmjs.com/cli/using-npm/workspaces/
