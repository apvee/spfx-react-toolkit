# F4 — servizi PnP, lifecycle e audit complementare

Autore: `services_fix`, con ownership espansa dall'orchestratore dopo riproduzione dei nuovi bug. Revisione indipendente ancora richiesta. Nessuna dipendenza, configurazione, firma pubblica, peer range, export/deep path o serializzazione legacy modificati.

## Correzioni e causa

| Area | Causa provata | Correzione | Prova comportamentale |
|---|---|---|---|
| B09 cache PnP | `Caching` installato legge `expireFunc`; il vecchio `timeout` è ignorato | adapter millisecondi → `new Date(Date.now() + timeout)` a ogni scrittura | PnP reale: prima/dopo scadenza1250ms in session/local; default e0 restano300000ms |
| B10 factory | JSON.stringify omette funzioni, perciò factory diverse producevano configKey identico | dipendenza esplicita `config.cache.keyFactory` nel memo dell'istanza | React reale: stesso contenuto inline conserva istanza; nuova factory produce nuova istanza e chiave cache |
| CrossSitePermissions | nessuna invalidazione dell'effect vecchio | cleanup `disposed`, azzera grant al cambio target | A→B, B poi A; clear URL; errore vecchio; unmount; stesso target senza nuovo fetch |
| UserPhoto | mount guard consentiva richieste vecchie dello stesso componente | identity di servizio/utente/email/size e requestId; reset dati e revoke URL al cambio identity | reverse identity e reload; loading ultima richiesta;403; unmount e rilascio URL |
| Performance | `time(name)` sovrascriveva sempre `name-start` globale | sequenza scalare per nomi mark distinti, cleanup del proprio mark nel finally |150ms/50ms con stessa label; nome/result preservati; throw preservato; mark/measure manuali intatti |
| Teams | SDK SPFx1.21 espone `{teamsJs, context}`, mentre hook cercava solo API dirette; cache initialized non seguiva context replacement | unwrap `teamsJs`, conserva direct v1/v2; reset runtime Teams al nuovo SPFx context | wrapper/direct v2 + direct v1; theme; replacement; old response e unmount |
| API precheck | autoCheckKey e risultati restavano validi dopo cambio tokenProvider | provider incluso nell'invalidazione di risultati/requestId/autoCheckKey | provider nuovo acquisisce token; risultato vecchio e in-flight scartati; passive popup/redirect guard+cleanup; retry cached/fresh; unmount; grouping e timeout |
| Demo suggestions | mount guard accettava qualunque risposta e input appena cambiato non invalidava richiesta durante debounce | token incrementato all'input/selezione/unmount, propagato al callback300ms | reverse completion, clear input, input durante debounce, unmount, selezione suggestion |

Non è stato aggiunto batching implicito tramite config: il parametro esistente resta invariato. Anche timeout0 mantiene esplicitamente il fallback storico di cinque minuti.

## RED / GREEN osservati

- Primo `node --test tests/services.test.cjs`:5 test,3 RED attesi (timeout session/local e factory cambiata),2 fallback baseline PASS. Dopo fix B09/B10:5/5 GREEN.
- Primo `node --test tests/demo-suggestions.test.cjs`:3/3 RED attesi per reverse/input clear/input durante debounce. Dopo fix:3/3 GREEN.
- Espansione CrossSite/UserPhoto/Performance:12 test,7 RED attesi sui nuovi comportamenti,5 precedenti PASS; dopo fix:12/12 GREEN.
- Espansione Teams/API:30 test complessivi servizi+demo,4 RED attesi (wrapper v2, context replacement inizializzato, provider reset, vecchio in-flight). Dopo fix:30/30 GREEN.
- Suite finale `node --test tests/services.test.cjs tests/demo-suggestions.test.cjs`:32/32 PASS,0 fail. Log locale `/private/tmp/spfx-services-final-green.log`.
- `npm test` finale sul worktree corrente:79/79 PASS,0 fail. Log `/private/tmp/spfx-services-npm-test.log`.
- `npm run typecheck`: library+app PASS; `npm run lint --workspace @apvee/spfx-react-toolkit`: PASS; lint del file demo assegnato PASS. Log locali `spfx-services-typecheck-final.log`, `spfx-services-lint-final.log` in `/private/tmp`.
- `git diff --check`: PASS sul worktree corrente.

Test delle guardie sono React17.0.1 e ReactDOM reali in jsdom15.2.1. Il loader transpila sorgenti TypeScript e dipendenze PnP ES modules installate senza modificarle: queryable, SPFx behavior, caching/storage e parsing sono reali; solo HTTP send è sostituito con Response controllate. Teams usa runtime store/selector/actions reali e SDK boundary controllato; API precheck usa servizio e JWT helpers reali, token acquisition/eventi controllati. Performance ha adapter clock/mark deterministico della piattaforma. Debounce usa mock timers Node, con timer ID window/global allineati come nel browser.

Warning atteso: Node22 segnala MockTimers come experimental. Full suite stampa anche logging di errori intenzionalmente iniettati nei test async degli altri owner; nessun test fallisce. I log sono locali temporanei, il documento riporta gli esiti osservati senza dipendere dalla loro persistenza.

## Audit read-only delle aree restanti

- Metadata (helper page-context e facades user/site/list/locale/correlation/environment/page type/permissions): ispezionate normalizzazione, optional fallback e dipendenze memo; nessun nuovo difetto dimostrato nella mappatura scalare. La rilevazione host dipende dalle forme reali del contesto e non ha validazione tenant in questa task.
- Logger: letti estrazione WebPart tag, handler custom, livelli e dipendenze callback; nessun bug confermato nel contratto provider valido.
- Performance, Teams, UserPhoto, CrossSitePermissions e API precheck: difetti sopra riprodotti prima di edit e risolti con regressioni persistenti.
- Batch PnP generic/list: letti avvio richieste prima di execute, allSettled, estrazione ID/errori e summary; nessun ulteriore bug dimostrato durante audit circoscritto. Batch HTTP multipart reali e dipendenze sequenziali nel callback non sono coperti da questa suite.
- REST/Graph error mapping: letti controlli response.ok nei servizi catalog/property/KV e fallback read/write; testato rejection network CrossSite,403 GraphPhoto e stale rejection. Il Graph client installato seleziona blob per Content-Type image, quindi non è stato inventato un problema responseType. La semantica dei singoli endpoint REST, throttling/permission payload e statusText richiede tenant e resta non verificata end-to-end.
- Serializzazione legacy mantenuta; tipizzazione delle stringhe numeriche/boolean/bigint non cambiata.

L'audit è circoscritto e non equivale a una validazione completa delle risposte SharePoint/Graph o dell'host Teams reale. Il codice gestisce risultati obsoleti senza cancellare richieste remote già avviate. Nessun publish/push/merge/deploy eseguito.
