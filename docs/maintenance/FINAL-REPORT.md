# Rapporto finale locale — 2026-10-06

Il monorepo è implementato, verificato e integrato localmente su `dev` con autorizzazione utente. Checkout operativo: `/Users/fabiofranzini/GitHub/apvee/spfx-react-toolkit`. **Tutte le fasi locali e la verifica post-merge hanno ottenuto PASS. La validazione autenticata nel tenant resta BLOCKED / NOT RUN; il goal persistente non è dichiarato completo.** Nessun push, publish o deploy. Le prove iniziali svolte nel worktree sono conservate come storico.

## Architettura

npm workspaces, root privata, un lockfile v3 e build esplicito libreria → app. `packages/spfx-react-toolkit` pubblica `@apvee/spfx-react-toolkit@2.1.0`, compilata da TypeScript in ESNext/ES2020. `apps/spfx-react-toolkit-test` contiene WebPart, Application Customizer e Gulp SPFx 1.21.1; usa il pacchetto pubblico. Sono conservati nome, versione, peer, dipendenze runtime, 82 contratti di dichiarazioni, main/types, allowlist e deep path.

npm era già presente. I pacchetti autonomi con file:/npm --prefix avrebbero richiesto doppi lock/installazioni e linking più fragile. Due workspace non giustificano un nuovo orchestratore. La risoluzione Sass dopo hoisting è stata corretta e provata anche nel consumer esterno.

## Correzioni e verifiche

Corretti loop storage con default oggetto, sincronizzazione displayMode/Property Pane e ServiceScope, loading/errori concorrenti, await TenantKV, risultati obsoleti OneDrive/TenantProperty/AppCatalog, query/refiners/paging PnP, timeout e keyFactory cache, permessi cross-site/foto, misure con stessa label, wrapper Teams, precheck con provider sostituito, autocomplete e lifecycle del placeholder.

La revisione indipendente ha trovato anche due casi PnPList residui: CRUD concluso dopo cambio query e offset perso dopo refetch fallito. Corretti per tutti i sei metodi CRUD/batch, con nuova revisione. Risultato: **90/90 test locali e 82 dichiarazioni compatibili**. Prove RED/GREEN persistenti, React reale e trasporto/SDK controllati. Cause, impatti e correzioni in [FINDINGS](FINDINGS.md) e nei rapporti autori.

## Ottimizzazione e pulizia

Build della sola libreria osservato a **1,24 s**, rispetto ai **7,64 s** della precedente pipeline accoppiata: misura singola, non confronto statistico del gate completo. Eliminati Sass/webpack demo dal build della libreria e secondo bundle nominale di gulp test, ora sostituito da test reali. Verifier store con compiler locale, directory temporanee uniche e cleanup; due esecuzioni concorrenti PASS.

Rimosse le dichiarazioni dirette inutilizzate ajv, sp-lodash-subset e sp-office-ui-fabric-core, il flag decorators e .npmignore obsoleto. Quattro nodi Fabric eliminati, nessuna versione esistente cambiata. API/deep path/moduli internal e configurazioni CDN/MCP conservati quando usi esterni non erano escludibili. Nessuna riduzione bundle rivendicata: tarball finale 331 file, 179922 byte; baseline 176648 byte.

## Documentazione e revisioni

README repository/npm, API e guide [sviluppo](../DEVELOPMENT.md)/[tenant](../SHAREPOINT-VALIDATION.md) allineate ai sorgenti. Corretti link e membri/esempi inesistenti. Tre ulteriori rilievi documentali del revisore sono stati corretti e ricompilati: URL cross-site, servizi SDK e lifecycle Customizer.

Tutti i subagenti effettivi usano GPT-6.1 Sol, con autori e revisori diversi. [REVIEWS](REVIEWS.md) e [matrice definitiva](VERIFICATION-MATRIX.md) separano PASS locali, FAIL corretti e prove esterne non eseguite.

## Comandi esatti

```bash
cd /Users/fabiofranzini/GitHub/apvee/spfx-react-toolkit
npm ci
npm run build
npm run verify
npm run verify:package
npm run package:solution
npm run pack:library
```

`npm run prepublishOnly` esegue il gate completo locale senza pubblicare. La sola libreria usa `npm run prepublishOnly --workspace @apvee/spfx-react-toolkit`: 83 test, nessun build app.

Dopo configurazione URL e certificato secondo guida:

```bash
npm run build:library
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve --workspace @apvee/spfx-react-toolkit-test
```

Il trust del certificato è un passo manuale del developer, non eseguito qui. Dopo modifiche alla libreria, ricompilarla e riavviare serve. Artefatti: `artifacts/apvee-spfx-react-toolkit-2.1.0.tgz` e `apps/spfx-react-toolkit-test/sharepoint/solution/apvee-spfx-react-toolkit.sppkg`; hash in evidence/artifact-sha256.json.

## Limiti e lavoro residuo

Nessun finding locale bloccante. Serve validazione autenticata SharePoint/Graph/Teams/permessi/batch secondo guida, in un ambiente già autorizzato senza deploy. Nessun URL/accesso tenant o risultato manuale è stato fornito. Nessuna simulazione è presentata come prova tenant.

Peer SPFx >=1.18 <2 conservato, ma prova effettiva limitata a SPFx 1.21.1/React 17.0.1/TS 5.3.3/Node 22.20.0. CI remota non lanciata; tutti i suoi comandi locali passano. Serializzazione tenant legacy e mappe senza sorgenti incorporati conservano limiti preesistenti documentati.

## Decisioni e costo se errate

- npm workspaces e SPFx dev 1.21.1 fissato: eventuali esigenze diverse richiederanno adattamento tooling.
- Formato tenant legacy conservato: restano ambiguità di stringhe numeriche/booleane e bigint; nuovo formato richiede disegno e migrazione dati/contratto dedicati.
- Deep/internal API e configurazioni CDN/MCP conservati: possibile superfluità residua, preferita alla rottura di usi esterni non dimostrabilmente assenti.
- Worktree conservato senza integrazione Git: integrazione successiva richiede autorizzazione separata per push/merge.

Stato goal finale: **BLOCKED**, dopo tre audit consecutivi del medesimo impedimento esterno. Ripresa possibile quando saranno disponibili accesso tenant autorizzato o prove manuali reali; il perimetro completo è conservato.

## Integrazione locale autorizzata su dev

Il 2026-10-06 l’utente ha autorizzato il merge locale e la prosecuzione nel checkout corrente. Il commit `c69e8e1` è stato integrato in fast-forward da `7e065362` su `dev`, senza conflitti. Le affermazioni precedenti su checkout preservato e assenza di merge descrivono la consegna precedente a questa autorizzazione. Checkout operativo: `/Users/fabiofranzini/GitHub/apvee/spfx-react-toolkit`. Nessun push, publish o deploy. Verifica e gestione del worktree documentate in [INTEGRATION.md](INTEGRATION.md). Il residuo tenant resta invariato.

## Pulizia root dopo merge

Rimossi lib/dist/release/temp/src della root: 508 residui generati, circa 13,4 MB, dopo backup verificato. Build, 90 test, API82 e packaging sono passati nel checkout principale; review indipendente PASS. Il vecchio worktree è archiviato in modo recuperabile. Dettagli in [ROOT-CLEANUP.md](ROOT-CLEANUP.md).
