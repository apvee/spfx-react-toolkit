# F4 — Provider e storage

Autore: provider_storage_fix. Worktree: `/Users/fabiofranzini/.codex/worktrees/spfx-toolkit-monorepo/spfx-react-toolkit`.
Baseline sorgenti: commit `7e0653622e8523d782ab9cf5dbbd185b03ace93b` (path originali `src/`). Data: 2026-10-06.

## Esito del gruppo

**PASS locale:** 14/14 assertion comportamentali con React 17.0.1, ReactDOM 17.0.1, JSDOM 15.2.1, compilatore TypeScript 5.3.3 locale e Node test runner. Typecheck libreria ed eslint dei due file modificati: exit 0. Revisione indipendente ancora da eseguire dall'orchestratore.

**Suite globale corrente: FAIL**, 31 test, 26 PASS e 5 FAIL, tutti nei test PnPList/PnPSearch assegnati all'altro gruppo ancora in lavorazione. Questo report non certifica l'integrazione finale.

## Cause e correzioni

| Difetto | RED osservato | Causa e correzione minima |
|---|---|---|
| B01 default oggetto inline | Assertion `storage must settle without repeated effect updates`, sia key assente sia JSON persistito | `defaultValue` ricreava l'effect e la lettura restituisce un nuovo oggetto, provocando un nuovo render. Ref aggiornato al default attuale; lettura iniziale/sync solo quando cambiano kind/key/instance ID; remove ed evento di eliminazione usano il ref corrente. |
| B03 displayMode stesso host | `1 !== 2` | Effect dipendeva dall'identità dell'istanza; aggiunta dipendenza scalare `instanceAny.displayMode`, verificati 1→2→1. |
| B04 Property Pane dopo setter | `'hook' !== 'pane'`; proprietà eliminata ancora presente | Host e runtime divergevano dopo il setter; riferimento host invariato impediva la sync. Ogni commit provider confronta chiavi proprie e valori con snapshot superficiale. Solo un cambiamento effettivo pubblica un nuovo snapshot; runtime→host aggiorna lo snapshot prima del refresh. |
| Scope replacement, confermato dai test | `ServiceScope consumed before finish`; callback dello scope vecchio poteva liberare quello corrente | Readiness associata allo scope invece che boolean globale; scope estratto ogni render, anche quando context viene mutato in-place. Guard immediata nel render; cleanup ignora completamento dello scope sostituito o dopo unmount. |

Nessun deep clone: proprietà annidate restano riferimenti condivisi. I normali aggiornamenti top-level, cambi di riferimento di valori annidati ed eliminazioni top-level sono osservati; la mutazione interna di un oggetto annidato con identità invariata non introduce una nuova capacità di osservazione. Il provider deve essere renderizzato dall'host perché la mutazione Property Pane venga riconciliata.

Semantica storage documentata nei JSDoc: il default inizializza **la key corrente** se manca o non è leggibile; cambiarlo da solo preserva il valore corrente e il riferimento dell'oggetto persistito. `remove()` ed eventi browser di eliminazione usano il **default attuale**. Cambiare key, instance ID o local/session rilegge il valore relativo oppure usa il default corrente. API/export/peer/versione invariati.

## Test e isolamento

`tests/provider-storage.test.cjs` e `tests/provider-test-harness.cjs` eseguono provider, store, selector, hook, guard strutturali e theme subscription **reali**. Stub solo sui confini SDK indisponibili (ThemeProvider service key, DisplayMode enum, ServiceScope e theme event fixture); browser/storage reali JSDOM. Nessun mock di lifecycle React o hook interni.

Casi: displayMode; setter→pane; delete prima/dopo setter e updater sostitutivo; parent rerender stabile senza refresh aggiuntivi; sostituzione bag e host; isolamento di due provider; theme subscription rimossa; vecchio setter dopo unmount non muta host; scope pronto→pending con context nuovo e context mutabile; completamento vecchio scope e dopo unmount; storage missing/persisted con oggetto inline; functional setter e parent rerender; cambio default/key/kind/instance; evento key/area diverso; eliminazione esterna e fallback attuale; listener storage rimossi all'unmount.

Il harness mantiene un unico DOM per il target eventi che ReactDOM cattura al caricamento, ripulisce DOM/storage/moduli source per test e chiude il DOM a fine file. Durante il caricamento di ReactDOM espone il MessageChannel del browser JSDOM (assente): usare il MessageChannel Node attiverebbe un port worker_threads che impedisce l'uscita del test runner. È adattamento browser di test, non modifica a React. Nessun warning soppresso nel percorso GREEN.

## Evidenza RED/GREEN e comandi

1. **Prima delle correzioni:** `node --test tests/provider-storage.test.cjs`, primi 11 test: 7 FAIL / 4 PASS. Log `/private/tmp/provider-storage-red.log`.
2. **Suite finale contro sorgenti baseline immutate:** archivio `git archive HEAD src`, ricollocato sotto packages in `/private/tmp/provider-storage-red.LZG2oQ`, test e harness finali copiati, node_modules locale collegato. Comando `node --test /private/tmp/provider-storage-red.LZG2oQ/tests/provider-storage.test.cjs`: exit **1**, 14 test, **10 FAIL / 4 PASS**. Log `/private/tmp/provider-storage-final-red.log`. Nessuna modifica temporanea ai sorgenti di lavoro.
3. **GREEN locale:** `node --test tests/provider-storage.test.cjs`: exit **0**, **14 PASS / 0 FAIL**. Log `/private/tmp/provider-storage-green.log` (esecuzione finale dopo la rimozione della soppressione console nel test).
4. `npm run typecheck --workspace @apvee/spfx-react-toolkit`: exit **0**.
5. `./node_modules/.bin/eslint packages/spfx-react-toolkit/src/core/provider-base.internal.tsx packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts`: exit **0**, nessun output/warning.
6. **Suite globale ripetuta sullo stato corrente**, `npm test`: exit **1**, 31 test, 26 PASS / 5 FAIL. Log `/private/tmp/provider-storage-npm-test.log`. Failure esterne allo scope:
   - `PnPList identity and latest query prevent stale state/error`
   - `PnPList rejects duplicate page dispatch and ignores page from superseded query`
   - `PnPSearch clears refiners on new search but keeps them on refetch`
   - `PnPSearch latest query ignores reversed responses and stale errors`
   - `PnPSearch same-tick loadMore dispatches one page and superseded page cannot append`

I casi TenantKV della suite globale stampano errori network attesi pur passando: non sono warning del gruppo provider/storage. Gli altri file possono cambiare durante il lavoro parallelo; l'orchestratore deve ripetere npm test e controlli d'integrazione dopo stabilizzazione.

## File modificati e limiti

Solo `packages/spfx-react-toolkit/src/core/provider-base.internal.tsx`, `packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts`, i due file test assegnati e questo report. Nessun package/lock/config/API pubblica, commit, push, publish, deploy o merge. Tenant SPFx reale non eseguito; fixture SDK verifica i contratti lifecycle dichiarati ma non sostituisce la verifica SharePoint.
