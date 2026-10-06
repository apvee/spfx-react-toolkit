# Registro di ripresa — piano PLAN.md

## Stato corrente

- Checkout operativo: `/Users/fabiofranzini/GitHub/apvee/spfx-react-toolkit`, branch `dev`. Baseline `7e065362`; commit monorepo `c69e8e1` integrato in fast-forward con autorizzazione utente del 2026-10-06. Nessun push/publish/deploy. Il precedente worktree è archiviato in modo recuperabile dopo verifica e conservazione degli artefatti ignorati. Vedere [INTEGRATION.md](INTEGRATION.md).
- F1–F7 PASS dopo revisioni e fix. F8 locale PASS: 90 test, API 82 moduli, clean install, build dev/ship, soluzione SPFx e tarball consumer. Validazione tenant richiesta: BLOCKED / NOT RUN.
- Stato autorevole: [REVIEWS](REVIEWS.md), [MATRICE](VERIFICATION-MATRIX.md), [RAPPORTO](FINAL-REPORT.md). Rapporti e log iniziali mantengono lo storico FAIL, superato dalle re-review finali.
- Ultime modifiche di prodotto: PnPList e tre esempi documentali, revisionati indipendentemente. Poi eseguiti clean prepublishOnly/ship, library prepublish 83 test, consumer tarball e pack definitivi. Nessun finding locale necessario residuo.
- Ultima integrazione documentale: guida manuale batch multipart/cache, sottoposta a review mirata. Modifiche maintenance successive riguardano soltanto registrazione e rapporto.

## Goal e impedimento esterno

Goal persistente BLOCKED dopo il terzo audit consecutivo; non completo e non paused. Mancano URL/accesso tenant autenticato e ambiente test autorizzato già disponibile oppure prove manuali reali secondo guida. Non eseguire deploy. Domanda async presentata, nessuna risposta tenant registrata.

Audit consecutivi del medesimo blocco esterno: **3** (turn originale e due continuazioni automatiche). La continuazione precedente è classificata NO PROGRESS: nessun nuovo accesso o risultato tenant, nessun processo live da attendere. Il terzo audit riconferma sorgenti invariati, placeholder tenant e assenza di nuovi input; nessuna ulteriore azione locale può completare la prova richiesta. update_goal consente BLOCKED solo dopo tre goal turn consecutive con lo stesso blocco e senza progressi significativi possibili. Non contare tool call come goal turn; alla ripresa verificare nuovi input e stato prima di incrementare. Non ripetere verifiche già concluse senza modifiche o nuovi finding.

## Decisioni conservate

- Ruling: npm workspaces e SPFx dev 1.21.1 fissato; semplicità e nessun major implicito. Costo se errato: adattamento tooling consumer.
- Ruling: serializzazione legacy conservata; nuovo formato può rompere contratti/dati persistiti. Costo: ambiguità documentata residua.
- Ruling: API deep/internal e configurazioni CDN/MCP conservate; usi esterni non escludibili. Costo: possibile superfluità residua.
- Ruling aggiornato: merge locale su dev autorizzato dall’utente; lavoro nel checkout corrente e worktree archiviato. Nessun push/publish/deploy.

## Non ripetere

Brainstorming delle tre prospettive, design/baseline, tutti i fix e le revisioni (incluse due PnPList e tre docs), clean install/gate finali e consumo tarball sono conclusi. Leggere la matrice e le prove finali. Restano i controlli esterni tenant; artefatti tgz/sppkg e guida sono pronti. Nessuna API eliminata e tutti i contratti peer/runtime/dichiarazioni baseline conservati.

## Prima continuazione automatica

Ricontrollati 183 hash dei sorgenti e hash tgz/sppkg: tutti invariati. Checkout originale ancora pulito. Configurazione serve conserva `{tenantDomain}`; nessun URL o accesso tenant né risultato manuale è arrivato. Nessun processo verificato ancora in esecuzione da attendere e nessun finding locale aperto. Registrato secondo audit; goal ancora attivo, non completo.

## Seconda continuazione automatica — terzo audit

Blocco esterno confermato per la terza goal turn consecutiva: manca un tenant autenticato già autorizzato, oppure risultati manuali reali. Goal impostato BLOCKED senza ridurre obiettivo o dichiararlo completo. Alla ripresa esplicita dopo BLOCKED, avviare un audit nuovo (non riutilizzare il conteggio3), verificare nuovi elementi e proseguire dal residuo tenant; non rifare lavoro locale concluso.

## Cambio di workflow autorizzato

La nuova richiesta utente autorizza il merge locale su dev e il lavoro nel checkout corrente, sostituendo il precedente vincolo no-merge per questa integrazione. La prova tenant rimane BLOCKED e il goal non è dichiarato completo. In futuro usare il checkout corrente, salvo diversa richiesta.

## Pulizia root conclusa

Rootlegacy lib/dist/release/temp/src rimosse (508 file, 13.410.937 byte), backup completo verificato in /private/tmp e manifest persistente. Build→verify90/API82→ship PASS; directory non rigenerate; root_cleanup_review indipendente PASS. Sorgenti e API immutati. Il prossimo comando di sviluppo si esegue dal checkout principale dev.

## Shortcut di avvio richiesto

Aggiunto npm run serve nella root: build library → app gulp serve, forwarding argomenti dopo --. Aggiornate guide/README. Build+help, verifier documenti e review indipendente PASS; dettagli ROOT-SERVE.md. Delta intenzionale rispetto allo snapshot storico: package.json principale e tre documenti pubblici, nessun runtime o dipendenza.

## Razionalizzazione evidenze richiesta

Rimossi22 log/harness intermedi o duplicati, nessun test o pagina pubblica eliminato. Due baseline operative trasferite byte-identiche in tests/fixtures, verify-api aggiorna solo2percorsi. Conservati ledger e rapporti unici, RED/GREEN e gatefinali; backup completo verificato e manifest persistente. npmverify90/API82 e review indipendente PASS. Il comando serve aggiunto resta invariato. Vedere EVIDENCE-CLEANUP.md; snapshot183 storico non rigenerato per questi delta intenzionali.

## Policy Git del monorepo richiesta

Creato AGENTS.md alla root: checkout corrente, niente worktree automatici; all’inizio di ogni incarico di modifica chiedere branch corrente oppure nuovo branch nello stesso checkout, salvo scelta già esplicita per quell’incarico. La scelta vale per continuazioni e subagenti; preservare modifiche preesistenti. Regola applicata ai workspace e revisione indipendente git_policy_review (GPT-6.1 Sol) PASS. Nessun prodotto, dipendenza, branch o worktree modificato per questa aggiunta; modifica eseguita su dev già concordato.

## Lingua AGENTS e proposta documentazione interna

AGENTS.md riscritto integralmente in inglese senza cambiare le regole; review indipendente git_policy_review PASS. Valutata proposta .docs/maintenance (versionata) e .docs/superpowers (attualmente locale/ignorata), mantenendo docs per API e guide pubbliche e AGENTS alla root. Nessun file spostato: l’utente ha richiesto prima una valutazione. La migrazione eventuale richiede aggiornamento link, .gitignore, verifier pubblico e convenzione in AGENTS; i riferimenti storici restano riconoscibili.

## Approvazione .docs e peer Fluent

Documentazione interna trasferita a .docs/maintenance e .docs/superpowers (67 file originali preservati); docs contiene le guide ufficiali. AGENTS root in inglese conserva policyGit e indica i nuovi percorsi; Superpowers resta ignorato. User ha richiesto esplicitamente entrambi Fluent come peer condivise: library peer+dev e app dependencies con stessi intervalli; tslib diretto superfluo rimosso, resta nei pacchetti SPFx che lo usano. Lock senza aggiornamenti versioni. VerifierAPI conserva baseline storiche e 82 dichiarazioni, con eccezioni precise autorizzate e controllo import dichiarati. 90test/build/typecheck/lint/packconsumer+fresh-no-lockconsumer/shared realpaths PASS, review indipendente PASS. Vedere DOCS-AND-DEPENDENCIES.md. Goaltenant ancoraBLOCKED, nessun esito esterno simulato.
