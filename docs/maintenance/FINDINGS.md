# Findings — registro iniziale

Tutti i file iniziali sotto src; dopo migrazione prefisso packages/spfx-react-toolkit/src. Autore analisi runtime_analysis, maintenance_analysis e architecture_analysis, GPT-6.1 Sol. Nessuna correzione ancora.

| ID | Priorità | Evidenza / atteso vs osservato | Correzione prevista | Verifica / stato |
|---|---|---|---|---|
| B01 | P1 | useSPFxStorage effect defaultValue oggetto inline; React reale 15 render e loop | default ref e sync key/storage | real React missing/persisted object + key switch; CONFIRMED |
| B02 | P1 | OneDrive file A→B, B risponde poi A; data A | identity/request epoch e createIfMissing vincolato | reverse promises, old404→no PUT current B, read/write ordering; CONFIRMED stale read, overwrite è conseguenza da verificare |
| B03 | P2 | provider stesso instance displayMode1→2 runtime1 | dipendenza valore scalare | real React1→2→1; CONFIRMED |
| B04 | P2 | hook setter → pane muta host in-place → runtime stale | snapshot/reconcile host senza loop | real React setter/pane/delete/isolation; CONFIRMED |
| B05 | P2 | TenantKV get/list return senza await; loading false prima risposta e catch bypassato | await nel try, concurrency coerente | deferred resolve/reject get/list; CONFIRMED |
| B06 | P2 | AsyncInvoke prima di2 promise termina loading | pending count + lifecycle | entrambi completion order/error/unmount; CONFIRMED |
| B07 | P2 | Search nuova usa refiners closure vecchia | map esplicita vuota | refined→new search vs refetch; CONFIRMED |
| B08 | P2 | TenantProperty keyA→B, B poi A lascia A; PnPList listaA→B lascia A | request/identity epoch | query switch/append/loadMore double call; CONFIRMED |
| B09 | P2 | Caching PnP legge expireFunc e ignora timeout | adapter timeout→expireFunc API invariata | frozen clock custom/default/zero; CONFIRMED da implementation PnP |
| B10 | P2 | JSON configKey omette keyFactory funzione | dipendenza identità keyFactory | due factory stesso config; da testare |
| M01 | P2 | gulp test senza config Jest: ricompila/bundla senza assertion; prepublish fa doppia build | test Node espliciti, separa app/lib | suite comportamentale e misura pipeline; CONFIRMED |
| M02 | P2 | verifier store tmp fisso + npx può scaricare compiler | mkdtemp/cleanup/local compiler | due run paralleli; CONFIRMED rischio |
| M03 | P2 | manifest SPFx >=1.18 può risolvere major toolchain successiva | pin versioni effettive1.21.1 | lock/clean install; CONFIRMED |
| C01 | P3 | ajv/lodash/fabric dev diretti nessun uso; transitivi toolchain | rimuovi solo dichiarazioni dirette | rg+ci+build; CANDIDATE |

## Ipotesi / aree da completare
Scope ready replacement, CrossSitePermissions stale/unmount, AppCatalog URL invalidation, TenantKV permission/provision mutex, PnPSearch batched hasMore, suggerimenti demo stale request. Verificare prima modifiche. Serializzazione legacy perde tipizzazione string numeriche/boolean/bigint: non cambiare formato senza progetto compatibilità, conservare documentando. Hook metadata, performance/logger, batch service/error REST/Graph e API precheck ancora da analizzare.

## Fonti riproduzioni
Harness analisi read-only in `/private/tmp/spfx-runtime-analysis.cjs`, `/private/tmp/spfx-provider-react-analysis.cjs`, `/private/tmp/spfx-storage-react-analysis.cjs`. Saranno trasformati in suite persistenti; fixture confinano solo SDK SPFx/browser indisponibile.

## B16 — Application Customizer placeholder lifecycle (P2)
Autore orchestratore. onDispose placeholder non azzera topPlaceholder; evento navigazione riprova render nello stesso elemento disposed anziché ricrearlo. `tests/app-lifecycle.test.cjs` React reale + boundary SDK RED1!=2 dopo disposal. Fix associa callback al PlaceholderContent creato, azzera riferimento corrente dopo disposal e ignora callback vecchie; cleanup React/event subscription verificati. Review indipendente ancora pending.

## Finding aggiuntivi confermati e revisionati

| ID | P | Atteso/osservato, causa | Correzione e verifica | Stato review |
|---|---|---|---|---|
| B11 | P2 | CrossSite URL A→B mantiene maskA quando A termina per ultima; effect senza cleanup | target change reset + disposed guard, realReact reverse/clear/error/unmount | PASS runtime review |
| B12 | P2 | UserPhoto userA→B blobA sostituisce B, mountflag solo | service/user/email/size identity + latest request/revoke URLs, reverse/reload403/unmount | PASS runtime review |
| B13 | P2 | Performance time same label sovrascrive markstart:150ms diventa50ms | marknames unici scalari+cleanup per operazione, deterministicclock150/50/throw/manualmark | PASS runtime review |
| B14 | P2 | Teams SPFx wrapper teamsJs/context ignorato; supportedfalse, initialized cache oltrecontext | unwrap wrapper, preservev1/v2, reset per context, realstore+Reactcontextswitch | PASS runtime review |
| B15 | P2 | Precheck tokenProvider cambiato risultati/key vecchi conservati | request/key/results invalidate, recheck, passiveevents/retry/unmount/groupingtimeout | PASS runtime review |
| B17 | P2 | PnPList queryA creatependente queryB createresolve→refetchA | currentquery/generation debounce, regression authorfixround1 | FAIL2 async review; fix pending |
| B18 | P2 | PnPList refetch fallito resetta currentSkip0 ma items precedenti restano; loadMoreduplica prima pagina | preserveoffsetuntilsuccessfulquery, regression authorfixround1 | FAIL2 async review; fix pending |
| D01 | P2 | HTTP CrossSite doc legge baseUrl closure prima nuovo render | explicit targetUrl, TS snippet compile | re-review pending |
| D02 | P2 | Doc EventAggregator export nonesistente/SPPermission nonserviceKey SDK1.21.1 | PageContext/customservice reali, TSsnippet compile | re-review pending |
| D03 | P2 | Provider app doc ignora delayedplaceholder/disposal e riproduceB16 | changedEvent+capturedplaceholder+Reactcleanup/recreate, TSsnippet compile | re-review pending |
| V01 | P2 | Nuova fixture consumer senza returnannotation rompe --ship fatal warning | returntypeesplicito con tipi pubblici storage/perf; no check attenuato | tarballrerunPASS, re-review pending |

B01–B10 correzioni descritte nei tre report autori; runtime reviewer approva provider/storage e servizi, async gruppo attende B17/B18. Le due regressioni review PnPList sono confermate anche su HEAD iniziale, non introdotte dalla migrazione. B16 app approvato revisore indipendente (SDK callback creation/disposal letto direttamente).

## Pulizia deliberata / conservazioni
Rimosse soltanto dichiarazioni dirette ajv/sp-lodash-subset/sp-office-ui-fabric-core, decoratorflag senza decoratori e npmignore obsoleto. Lock non cambia versioni esistenti; Fabric transitivo esclusivo sparisce (4nodes). Ajv/lodash necessari alla toolchain restano transitivi. APIs/deeppaths/internal modules sono conservati: vecchi TenantKV .internal non hanno usi library attuali ma erano inclusi nel tarball e potrebbero avere consumer deep, quindi nessuna cancellazione. Config deploy-CDN eMCP conservate perché uso esterno/dinamico plausibile, senza esecuzione/deploy. Source/declaration maps senza sourceembedding conservate come baseline, limite documentato.

## Disposizione finale
B01–B18 eAppCatalog/demo/storage-events nei report: CORRECTED+VERIFIED+independentreviewPASS. B17/B18 re-review2/2+90testsPASS. D01–D03 eV01 ADDRESSED re-reviewPASS. M01/M02/M03 eC01 verificate eF3/F5/F6PASS. Nessun finding localebloccante residuo. Limiti legacyserialization/maps e test esterni sono documentati e nonpresentati come correzioni oPASS.
