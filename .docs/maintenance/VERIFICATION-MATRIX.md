# Matrice finale — 2026-10-06

**Stato locale: PASS. Goal globale non dichiarato completo: prova autenticata SharePoint/Graph/Teams indisponibile.** Nessun push/merge/publish/deploy. Comandi eseguiti sul sorgente definitivo dopo i fix e le re-review; log in [evidence](evidence/).

| Controllo | Esito | Evidenza / limiti |
|---|---|---|
| Stato Git iniziale / preservazione originale | PASS | HEAD7e065362, `dev` originale pulito inizialmente e alla consegna; modifiche nel worktree Codex isolato |
| Design/baseline/alternative/piano | PASS | Review indipendente F1/F2 e baseline82hash/pack331; [REVIEWS](REVIEWS.md) |
| Installazione pulita dal lock | PASS | `npm ci --no-audit --no-fund` in `/private/tmp/spfx-toolkit-clean-y3wx2d65`,2423 packages,[log](evidence/clean-install.txt); nessun ignore-scripts esplicito. npm11 default policy non esegue4 install script legacy/optional non allowlisted: fsevents/es5-ext. Non si dichiara esecuzione di tutti scripts |
| Clean/build library+app | PASS | `npm run prepublishOnly` nella copia pulita esegue clean→buildlib→Gulpdevapp→verify; [log](evidence/final-clean-release.txt) |
| Suite comportamentale finale | PASS |90/90 Node:test, realReact17/JSDOM; provider/storage14,async42,services27,demo5,app1,confini1; [independent](evidence/runtime-rereview-integration.txt). Mock SDK/HTTP/token controllati, non tenant |
| Gate library indipendente app build | PASS | Library `prepublishOnly`: clean/tsc/lint/test83/83,[log](evidence/final-library-prepublish.txt); nessun build/bundle app richiesto |
| Typecheck/lint entrambi workspace | PASS | Parte `verify` clean,0 errori/0 warning rilevanti; nessuna regola/tipo indebolito |
| verify:examples | PASS inventario |40 hook + 4 provider, registry preservato. Non dimostra mount tenant di FieldCustomizer/CommandSet |
| verify:runtime-store | PASS | Store assertions +2 run realmente paralleli senza interferenze/tmpresidui, compiler locale |
| verify:public-docs / link / snippet | PASS nel perimetro | Inventory helper/service e relativefilelinks incl npmREADME; ancore manuali reviewer;16snippet iniziali+4fixsnippet compilati; non ogni frammento illustrativo legacy compilato |
| API/declarations/peer/runtime deps | PASS |82 d.ts AST contract ugualebaseline; name/version/main/types/files/peer/deps invariati,moduloESNext; [verify final](evidence/final-clean-release.txt) |
| Workspace integration / confini | PASS | ASTimports app rootpackageAPI; library nessuna dipendenza app; zero alias source; reviewer package/types/manifest/XML |
| Dependency versions/dedup | PASS |0 cambi di versione esistenti; runtimeReact/DOM17.0.1 SPFx1.21.1 PnP4.17.0 dedup. TS4.2.4 transitiveAPI-extractor distinta dal buildTS5.3.3 |
| App ship+solution clean | PASS |`npm run package:solution`, [clean](evidence/final-clean-ship.txt) e [worktree](evidence/final-worktree-ship.txt), `.sppkg` prodotto senza warning lint |
| Pacchetto npm reale | PASS |331 file, 179922 byte compressi; path esatti baseline, solo library/README/LICENSE/package; nessun appartifact/src/config/test; [pack](evidence/final-pack.txt) |
| Tarball consumer esterno | PASS |Install npm del.tgz senza workspace/symlink, TypeScript public+deep+documentedcontracts, runtime `npm ls`, SPFxship e.sppkg; [log](evidence/final-tarball-consumer.txt). Tempeliminato automaticamente |
| Regressioni aggiuntive reviewer | PASS |B17/B18 originalrepro2/2, options/laziness2/2, authorRED7/7→GREEN7/7, scopedre-reviewpass |
| CI coherence | PASS statico/localcommands | `.github/workflows/verify.yml` stessa sequenza install/build/verify/ship/tarball; checked setup-node/checkout official v4 docs. Esecuzione runner GitHub NOT RUN, nessun push autorizzato |
| Tenant WebPart/ApplicationCustomizer, permessi/Graph/Teams/storage/batch | BLOCKED / NOT RUN |Nessun tenantURL/accesso/autorizzazione test reale disponibile. [Procedura riproducibile](../../docs/SHAREPOINT-VALIDATION.md); no deploy/simulazione. Risultati manuali devono essere riportati separatamente |
| Compatibilità SPFx>=1.18<2 eReact17.x | DICHIARATA |Peer conservati. Prova effettiva SPFx1.21.1/React17.0.1/TS5.3.3/Node22.20.0 soltanto; non matrice universale |
| Formato persistito / sourcemaps | CONSERVATO, limite noto |Tenant legacyJSON può coerce stringnumeric/bool/null eperdere bigint precision; no runtimevalidation promessa. Sourcemaps nonembed source comebaseline. Publicdeeppaths/internal conservati |

Le prime failure finali non sono nascoste: warning prefer-const app e annotation fixture consumer corrette e riprovate; tre docsP2 e due PnPListP2 corretti e re-reviewPASS. Nessun finding locale bloccante residuo. L'assenza di bug assoluta non è affermata.
