# Fluent style utilities — analisi ottimizzata

Data: 7 ottobre 2026, Europe/Rome.

Stato: quarta revisione. Consolida descriptor diretti, nomi descrittivi, `container.inlineSize` e la sola ricetta `scrollbar.fluent`. Il piano operativo riallineato è [UseSx Implementation Plan](../superpowers/plans/2026-10-07-use-sx-implementation-plan.md). I checkpoint di integrazione/browser e la validazione dei preset precedono il completamento; nessuna implementazione di `useSx` è attestata da questo documento.

Branch: `apvee/fluent-style-utilities`, creato da `dev` nel checkout corrente. Al momento della registrazione entrambi puntano a `3135cd4735af969ee234b42c720d34ee6fc2ce20`. Il checkout contiene modifiche locali precedenti, incluse le istruzioni aggiornate e `useStableCallback`; queste modifiche sono preservate e non sono state committate nell'ambito di questa analisi.

## Obiettivo e contesto

Introdurre una facciata tipizzata sopra Griffel per accelerare la costruzione di interfacce Fluent UI 9 nei progetti SPFx. Il toolkit serve sia i progetti interni sia sviluppatori esterni. Le utility devono restare indipendenti dal runtime SPFx e avere confini che permettano una futura estrazione.

Il confronto nasce anche dalla chat [Definire layout Fluent UI](chatgpt-conversation://6ac40fb8-cbe8-83ed-8276-621aaa11b2fc). Il materiale recuperato propone un DSL compatto con descriptor tipizzati, utility responsive e una distinzione fra overflow e comportamento dello scrolling. Il suo contenuto è stato utilizzato come materiale di analisi; le decisioni correnti derivano dalle successive conferme dell'utente in questa chat.

Il perimetro comprende layout e stili Fluent, ma la sua organizzazione è stata semplificata: poche utility ricorrenti, ruoli semantici foreground/background, preset cromatici documentati e riutilizzo della tipografia ufficiale. Non si costruisce un vocabolario per ogni proprietà CSS, palette, variante o combinazione di stato.

Le parti non coperte rimangono componibili attraverso classi esistenti prodotte con Griffel. Questa possibilità è il punto di estensione principale: non richiede di duplicare l'intera API CSS nel wrapper.

La revisione è stata confrontata da tre analisi indipendenti: ergonomia e preset, integrazione del catalogo con variabili/stati, compatibilità dei token e verifiche del consumer. I risultati sono integrati di seguito; non costituiscono una revisione di codice implementato.

## Decisioni confermate

- Un unico pacchetto: `@apvee/spfx-react-toolkit`. Nessun nuovo pacchetto o worktree per questa fase.
- `useSx()` restituisce una normale funzione `sx(...)`, utilizzabile inline in `className` e su più elementi.
- Il risultato pubblico è una stringa di classi. Non viene richiesto un attributo `style` né un oggetto da distribuire come props JSX.
- Griffel rimane il motore: gestisce generazione delle classi, hashing, inserimento delle regole e composizione atomica.
- Le regole delle proprietà sono predefinite e consumano CSS variables; anche le classi che assegnano i valori possono essere prodotte da Griffel core.
- Il wrapper usa il renderer attivo tramite un adapter interno su `useRenderer_unstable`. Non introduce un nuovo provider obbligatorio.
- Il responsive usa il contenitore come riferimento predefinito; il viewport si seleziona esplicitamente.
- `container.inlineSize` è un descriptor diretto: stabilisce il query container per i discendenti, senza una chiamata di funzione vuota.
- Le proprietà omesse in un override responsive conservano il comportamento base. I conflitti sulla stessa proprietà e nello stesso contesto devono avere una precedenza ordinata e documentata.
- Le utility devono comporsi con classi esistenti e con i risultati di altre chiamate a `sx`.
- I descriptor completi si passano direttamente a `sx`: niente wrapper di famiglia ridondanti per dimensioni, spaziatura, foreground o preset tipografici. Le funzioni rimangono per costruire valori, applicare scope o stabilire un confine di query.
- I nomi pubblici descrivono proprietà o ruoli: `padding`, `foreground`, `width`, `minWidth`, `typography.body1`. Non mantenere una seconda grammatica abbreviata equivalente.
- Foreground e background usano un insieme ridotto di ruoli Fluent; i nomi semantici del toolkit devono avere mapping espliciti, evitando sinonimi e cataloghi completi di token.
- I preset cromatici sono composizioni del toolkit. La tipografia riutilizza invece i preset ufficiali di `typographyStyles`, senza una scala alternativa.
- Le varianti CSS iniziali sono soltanto `hover`, `active` e `focusVisible`. Gli stati persistenti dell'applicazione scelgono descriptor tramite input condizionali di `sx`.
- La personalizzazione delle scrollbar è una sola ricetta `scrollbar.fluent`. Non si esportano gutter, overscroll, smooth scrolling o ulteriori varianti di scrollbar nella v1.
- Ogni nuova capacità pubblica deve avere JSDoc, documentazione Markdown e un caso di prova funzionante nella web part, oltre ai test automatici appropriati.

## API illustrativa

Questo esempio rappresenta la direzione concordata; le factory e i tipi non sono ancora implementati.

```tsx
const sx = useSx();

return (
  <section className={sx(container.inlineSize)}>
    <div
      className={sx(
        presets.canvas,
        flex.column,
        gap.medium,
        padding.medium,
        width.full,
        minWidth.zero,
        typography.body1,
        foreground.subtle,
        responsive.medium(flex.row, gap.large),
        isSelected && presets.brandTint
      )}
    >
      {/* contenuto */}
    </div>
  </section>
);
```

`container.inlineSize` stabilisce il riferimento per i discendenti: il responsive CSS non misura automaticamente la larghezza dell'elemento al quale è applicata la regola.

Ogni descriptor identifica già una proprietà CSS o una ricetta completa. `sx` li compone senza ulteriori wrapper; `alignItems.center` e `justifyContent.center` restano distinti perché operano su proprietà differenti. Token e valori custom devono essere riconoscibili: `gap.medium` indica un token documentato, `gap.px(12)` un valore con unità CSS esplicita.

## Nomenclatura descrittiva e descriptor diretti

| Nome precedente | Nome pubblico proposto |
| --- | --- |
| `p` | `padding` |
| `m` | `margin` |
| `fg` | `foreground` |
| `bg` | `background` |
| `w`, `h` | `width`, `height` |
| `minW`, `minH` | `minWidth`, `minHeight` |
| `maxW`, `maxH` | `maxWidth`, `maxHeight` |
| `items`, `justify`, `self` | `alignItems`, `justifyContent`, `alignSelf` |
| `item` | `flexItem`, per grow/shrink di un elemento flex |
| `radius`, `shadow` | `borderRadius`, `boxShadow` |
| `grid.cols(...)` | `grid.columns(...)` |
| Suffissi `xs`, `s`, `m`, `l`, `xl` | `extraSmall`, `small`, `medium`, `large`, `extraLarge` |
| Breakpoint `sm`, `md`, `lg` | `responsive.small/medium/large`, oppure `viewport.small/medium/large` |

`paddingInline`, `paddingBlock`, `marginInline` e `marginBlock` sono le proposte per gli assi logici. Non chiamarli `paddingX/Y`: inline/block seguono il writing mode e non sono sinonimi universali di sinistra/destra e alto/basso. Token e proprietà da applicare devono essere documentati con precisione. `.px(...)` conserva invece il nome dell'unità CSS standard; non è un'abbreviazione inventata per una proprietà.

La grammatica pubblica distingue tre casi:

- Selezioni dirette: `width.full`, `minWidth.zero`, `padding.medium`, `foreground.subtle`, `typography.body1`.
- Costruttori con un argomento: `width.px(240)`, `gap.px(12)`, `grid.columns(3)`.
- Operazioni di scope: `responsive.medium(...)`, `viewport.medium(...)`, `hover(...)`, `active(...)`, `focusVisible(...)`.

`flex.column` e `flex.row` sono ricette complete che stabiliscono display flex e direzione. `grid.columns(...)` stabilisce display grid e una ricetta documentata di colonne. `gap`, allineamenti e `flexItem.grow` restano proprietà modificatrici: non introducono implicitamente il display o garantiscono un parent DOM adatto.

Famiglia e origine rimangono metadati interni dei descriptor per diagnostica e validazione. Non è necessario un wrapper pubblico `size(...)`, `spacing(...)`, `surface(...)` o `typography(type...)` per conservare quei metadati. I preset composti si espandono in dichiarazioni prima della risoluzione degli override.

Le stringhe di classi già compilate sono ammesse al livello `sx`; gli scope operano sui descriptor, non su `hover(sx(...))`. Non introdurre automaticamente un `group(...)` o una seconda forma equivalente delle selezioni: `sx` è già il compositore e una lista riutilizzabile può essere distribuita con lo spread.

Questa revisione cambia la forma pubblica, non l'architettura del renderer, delle variabili o della cache. Non dimostra un miglioramento di tree shaking o bundle; non occorrono alias di compatibilità perché la capacità non è ancora implementata o pubblicata.

`isSelected` è uno stato dell'applicazione, non un valore inferito dal wrapper. Gli esempi non assegnano automaticamente ruoli HTML, attributi ARIA o comportamenti ai nodi.

## Ruoli foreground e background

Il livello pubblico propone pochi descriptor semantici. Sono alias del toolkit per token ufficiali: i nomi `primary`, `subtle` e `muted` non sono nuove famiglie esportate da Fluent.

| Foreground proposto | Token Fluent |
| --- | --- |
| `foreground.primary` | `colorNeutralForeground1` |
| `foreground.subtle` | `colorNeutralForeground2` |
| `foreground.muted` | `colorNeutralForeground3` |
| `foreground.disabled` | `colorNeutralForegroundDisabled` |
| `foreground.brand` | `colorBrandForeground1` |
| `foreground.onBrand` | `colorNeutralForegroundOnBrand` |
| `foreground.inverted` | `colorNeutralForegroundInverted` |
| `foreground.link` | `colorBrandForegroundLink` |

Non aggiungere alias come `secondary`, `tertiary` o `neutral2` per duplicare gli stessi ruoli. `muted` indica minore enfasi, non uno stato disabled. `link` applica colore, non comportamento o semantica di collegamento.

| Background proposto | Token Fluent |
| --- | --- |
| `background.canvas` | `colorNeutralBackground1` |
| `background.alternative` | `colorNeutralBackground2` |
| `background.subtle` | `colorSubtleBackground` |
| `background.transparent` | `colorTransparentBackground` |
| `background.brand` | `colorBrandBackground` |
| `background.brandTint` | `colorBrandBackground2` |
| `background.inverted` | `colorNeutralBackgroundInverted` |
| `background.disabled` | `colorNeutralBackgroundDisabled` |

`subtle` e `alternative` sono distinti. Nei temi standard installati, `colorSubtleBackground` è trasparente a riposo e ha token dedicati per hover/pressed; non è un alias di NeutralBackground2. `subtle` e `transparent` restano token distinti anche quando il loro valore a riposo coincide. Un tema custom può cambiare i valori.

Non si esporta un alias per ogni token di colore. Per necessità non coperte, si può comporre una classe Griffel esistente. Un eventuale accesso preciso a token aggiuntivi deve essere un'estensione tipizzata deliberata, non un secondo catalogo di nomi.

I riferimenti ai token restano CSS variables del tema Fluent. L'applicazione deve fornirle nel proprio scope, normalmente attraverso FluentProvider. La decisione di non richiedere un provider aggiuntivo per `useSx` non elimina la necessità di un tema quando si utilizzano quei riferimenti.

## Preset cromatici del toolkit

I preset sono descriptor composti, elaborati dallo stesso normalizzatore delle utility. Non diventano un motore separato e non generano automaticamente varianti interattive.

| Preset candidato | Background | Foreground | Proprietà applicate |
| --- | --- | --- | --- |
| `presets.canvas` | NeutralBackground1 | NeutralForeground1 | Background e colore |
| `presets.alternative` | NeutralBackground2 | NeutralForeground1 | Background e colore |
| `presets.subtle` | SubtleBackground | NeutralForeground2 | Background e colore |
| `presets.transparent` | TransparentBackground | Non impostato | Solo background |
| `presets.brand` | BrandBackground | NeutralForegroundOnBrand | Background e colore |
| `presets.brandTint` | BrandBackground2 | BrandForeground2 | Background e colore |
| `presets.inverted` | NeutralBackgroundInverted | NeutralForegroundInverted | Background e colore |
| `presets.success` | StatusSuccessBackground1 | StatusSuccessForeground1 | Background e colore |
| `presets.warning` | StatusWarningBackground1 | StatusWarningForeground1 | Background e colore |
| `presets.danger` | StatusDangerBackground1 | StatusDangerForeground1 | Background e colore |
| `presets.disabled` | NeutralBackgroundDisabled | NeutralForegroundDisabled | Background e colore |

I nomi nella tabella dei preset omettono il prefisso `color` dei token per leggibilità. Le coppie sono candidate e devono essere verificate sui temi e sfondi reali prima di essere pubblicate. La presenza dei token non dimostra contrasto appropriato per ogni combinazione o equivalenza con un controllo Fluent.

Questi preset non aggiungono padding, dimensioni, display, radius, shadow o geometria dei bordi. Le proprietà omesse rimangono intatte: applicare un preset non significa resettare tutto lo stile precedente. Radius, shadow e bordi rimangono composizioni esplicite.

La ricerca nei componenti Fluent ha individuato ulteriori riferimenti utili: Button ha appearance primary/secondary/outline/subtle/transparent; Card ha filled/filled-alternative/outline/subtle; Badge ha filled/ghost/outline/tint e più ruoli di colore. Queste sono API e ricette specifiche dei componenti, non preset generici ufficiali da importare o replicare integralmente.

`outline` non entra automaticamente nel catalogo cromatico: il solo colore del bordo non lo rende visibile. Un futuro preset outline deve definire esplicitamente width/style e il relativo impatto sul box, oppure lasciare la geometria a `border(...)`. Anche informative, important e severe restano riferimenti da valutare su un caso concreto, non nuove esportazioni automatiche.

L'esempio `presets.subtle` si ispira a una combinazione con Foreground2. Non pretende un significato identico di subtle in tutti i componenti: Button, Card e Badge adottano combinazioni differenti.

## Tipografia ufficiale

`@fluentui/react-theme` esporta già `typographyStyles`. Si propone una facciata di descriptor `typography.*` che conserva esattamente i nomi e i valori ufficiali, utilizzabile direttamente come `typography.body1` in `sx`.

Il catalogo installato contiene 17 preset:

```text
body1, body1Strong, body1Stronger, body2,
caption1, caption1Strong, caption1Stronger,
caption2, caption2Strong,
subtitle1, subtitle2, subtitle2Stronger,
title1, title2, title3, largeTitle, display
```

Ogni preset riutilizza fontFamily, fontSize, fontWeight e lineHeight ufficiali. Non introduce foreground, tag HTML o semantica di heading. Se non viene impostato un foreground, deve restare valido il normale comportamento di ereditarietà CSS del colore.

Non inventare nomi come body2Strong, title1Strong o caption2Stronger attribuendoli a Fluent. Non introdurre una seconda scala di dimensioni/pesi o una facciata raw di oggetti di stile per replicare questi preset.

## Varianti CSS essenziali e stati dell'applicazione

| Variante iniziale | Selettore | Responsabilità |
| --- | --- | --- |
| `hover(...)` | `:hover` | Feedback del puntatore |
| `active(...)` | `:active` | Feedback transitorio durante la pressione |
| `focusVisible(...)` | `:focus-visible` | Trattamento del focus quando indicato dal browser |

`focusWithin` rimane una possibile estensione per un'esigenza concreta di contenitore. Non si introduce una funzione per ogni pseudo-classe, combinazione di selector o stato ARIA.

Selected, expanded, loading, checked e disabled rimangono stati posseduti dall'applicazione o dal controllo. Il wrapper può selezionare descriptor tramite `isSelected && presets.brandTint`, ma non inferisce quegli stati e non assegna automaticamente attributi. Native disabled/checked e le props dei componenti Fluent mantengono i loro comportamenti e le loro semantiche.

`active` non significa selected, toggled o pressed persistente. Un wrapper di stato sposta le dichiarazioni nello scope CSS richiesto; non sceglie implicitamente un token Hover/Pressed in base al nome. Eventuali ricette interattive devono elencare i token e gli stati che applicano.

Un preset cromatico rimane a riposo. La v1 non richiede un catalogo di preset interattivi per tutti i ruoli: le combinazioni si esprimono tramite le tre varianti essenziali, quando servono a regioni personalizzate. I controlli Fluent continuano a utilizzare le proprie props appearance, focus e disabled.

Una variante hover/active presente non viene neutralizzata soltanto aggiungendo un preset disabled nello scope base. Un esempio di composizione corretta deve escludere le varianti abilitate quando disabilitato:

```tsx
<button
  disabled={isDisabled}
  className={sx(
    presets.canvas,
    !isDisabled && hover(presets.alternative),
    !isDisabled && active(presets.brandTint),
    isDisabled && presets.disabled
  )}
>
  Azione
</button>
```

Focus-visible può coincidere con hover e active. Per le proprietà sovrapposte occorre una precedenza documentata; l'indicazione visibile del focus deve rimanere presente durante le altre interazioni. Non rimuovere il focus nativo per imitare l'aspetto di una superficie. Grammatica iniziale: al massimo un wrapper di stato in uno scope responsive ammesso; niente pseudo-stati annidati.

## Modello di integrazione con Griffel

### Unica ricetta scrollbar

```tsx
sx(overflow.vertical.auto, scrollbar.fluent)
```

`overflow.vertical.auto` stabilisce quando il contenuto può scorrere. `scrollbar.fluent` applica solo l'aspetto: `scrollbar-width: thin` e un `scrollbar-color` composto da `colorNeutralStrokeAccessible` per il thumb e un track trasparente. Il token è una scelta iniziale concreta da verificare sulle superfici previste, non una garanzia di contrasto universale.

In forced-colors, la ricetta deve lasciare il controllo al browser tramite `scrollbar-color: auto`. Non forza `forced-color-adjust: none`, non nasconde le scrollbar e non imposta overflow, gutter o comportamento dello scrolling. I browser privi delle proprietà standard conservano il proprio aspetto nativo; non introdurre pseudo-elementi WebKit nella v1.

### Regole predefinite e assegnazioni dei valori

Una regola statica può consumare una variabile:

```css
width: var(--apvee-sx-width-base);
```

Il wrapper traduce un valore in una dichiarazione di custom property e la passa alla factory senza hook di `@griffel/core`:

```ts
const getVariableClasses = makeCoreStyles({
  root: { '--apvee-sx-width-base': '240px' },
});

const variableClass = getVariableClasses({ renderer, dir }).root;
```

Le classi delle proprietà e delle assegnazioni si compongono con `mergeClasses`. In questo modo anche le variabili partecipano al sistema atomico di Griffel. Non serve un generatore separato di nomi o un secondo renderer.

La cache del wrapper deve essere associata all'identità del renderer. Le factory core risolvono i nomi usando il salt del renderer: una factory inizializzata per un renderer non deve essere condivisa indiscriminatamente con renderer che hanno salt differenti. Questa esigenza riguarda anche la materializzazione del catalogo statico.

### Renderer e direzione

`@griffel/react.makeStyles` non supporta la creazione di nuove factory nell'esecuzione di un componente. Le definizioni statiche rimangono a livello modulo; l'integrazione dinamica usa la factory core, che non chiama hook React. Anche `sx()` deve rimanere una normale funzione, senza hook condizionali.

`useRenderer_unstable` recupera il renderer esistente ed è utilizzato anche da FluentProvider. È però esplicitamente marcato come API privata esportata: il collegamento va isolato e verificato sulle versioni supportate.

Il renderer non espone la direzione. Il lettore della direzione Griffel non è esportato dal suo entry point React e le sequenze atomiche contengono metadati LTR/RTL. L'uso di LTR fisso per le classi delle variabili non è una soluzione completa: `mergeClasses` segnala la composizione di sequenze con direzioni diverse.

Proposta da approvare: leggere renderer e direzione in un unico adapter, ricavando la direzione dal contesto Fluent; offrire un'opzione `dir` per l'uso con il solo TextDirectionProvider Griffel. Il contratto preciso, i peer necessari e i casi di direzioni discordanti devono essere fissati prima dell'implementazione.

## Prove eseguite e limiti dell'evidenza

Sono state eseguite prove tramite script in memoria, senza aggiungere fixture o modificare codice di produzione. Versioni ispezionate: React `17.0.1`, Griffel React `1.5.30`, Griffel core `1.19.2`, Fluent shared contexts `9.25.2`.

| Prova | Osservazione | Limite |
| --- | --- | --- |
| Custom properties passate a core `makeStyles` | Generate regole di classe per `240px` e `360px` | Verifica della generazione, non delle dimensioni visuali |
| Assegnazione ripetuta dello stesso valore | Riutilizzati nomi di classe; nessuna regola CSS duplicata | Non dimostra memoria limitata con valori sempre nuovi |
| `mergeClasses` su assegnazioni della stessa variabile | L'ultimo input ha mantenuto la propria classe atomica | Non copre tutti i conflitti fra responsive, reset e shorthand |
| React 17 con RendererProvider esistente | Regole statiche e variabili hanno usato lo stesso renderer | Prova DOM locale, non SharePoint autenticato |
| Aggiornamento `240 → 360 → 240` | Cambiata e poi riutilizzata la classe iniziale | Nessuna misura di bundle o prestazioni |
| Attributi del nodo della prova | Presente `className`, assente `style` | Non è una verifica del futuro DSL completo |
| Numero di regole nella prova React | Tre: una proprietà e due assegnazioni | Catalogo minimo, non rappresentativo dell'intero perimetro |
| Secondo documento/renderer | La regola è stata inserita anche nel documento distinto | Salt differenti ancora da verificare |
| Assegnazione isolata in RTL | Classe atomica invariata, metadati della sequenza differenti | Composizione completa RTL ancora da verificare |

Le prove non persistono come suite del repository e dovranno essere trasformate in test ripetibili durante l'implementazione. Non è stato eseguito un prototipo browser del responsive, né misurato il bundle, né validata la nuova capacità in un tenant SharePoint.

## Perimetro iniziale e dettagli proposti

| Famiglia | Direzione del perimetro |
| --- | --- |
| `flex.*`, `grid.columns(...)`, `flexItem.*` | Direzione, allineamento, gap, wrap, crescita e contrazione |
| Spaziatura e dimensioni | `padding.*`, `margin.*`, `gap.*`, `width.*`, `height.*`, min/max e valori con unità esplicite |
| `typography.*` | Preset ufficiali e comportamenti del testo ricorrenti, senza una scala parallela |
| Colori e superfici | `foreground.*`, `background.*`, preset cromatici, `borderRadius.*`, `boxShadow.*` e geometria dei bordi esplicita |
| `overflow.*`, `scrollbar.fluent` | Disponibilità dello scrolling per asse e unica ricetta visiva Fluent, senza altre utility di comportamento scroll |
| Responsive e stati | Contenitore/viewport espliciti; soltanto hover, active e focus-visible nella v1 |

Dettagli del piano ancora proposti, non decisioni definitivamente approvate:

- Input di `sx`: descriptor readonly tipizzati, classi esistenti, `false`, `null` e `undefined`; argomenti variadici piatti.
- Soglie mobile-first proposte: `responsive.small ≥ 480`, `responsive.medium ≥ 640`, `responsive.large ≥ 1024`; le omonime selezioni `viewport.*` usano la larghezza della finestra. Non sono sinonimi delle categorie SPFx esistenti: `medium`, per esempio, oggi indica 480–639 px.
- Nomi e mapping pubblici dei ruoli/preset proposti nelle tabelle, con accoppiamenti cromatici validati sui temi. Non espandere le factory in un catalogo generale CSS.
- Grammatica di combinazione fra responsive e i tre stati essenziali, compilata senza container query o pseudo-stati annidati.
- Opzione esplicita per la direzione nei contesti Griffel non Fluent.
- Validazione dei numeri in base alla proprietà: unità, valori finiti e dominio consentito.

## Rischi e checkpoint prima di ampliare il catalogo

1. **Precedenza responsive.** Le query sovrapposte non sono risolte automaticamente da `mergeClasses`. Nel core installato il comparatore delle media query è lessicografico e le container query condividono un bucket. L'ordine normalizzato dei descriptor, da solo, non dimostra il risultato della cascata. Un candidato è l'uso di variabili di attivazione per scope con fallback a priorità esplicita; va provato nel browser senza modificare il comparatore dell'host.
2. **Ereditarietà e reset.** Le variabili di input devono essere distinte per proprietà e scope. Le assegnazioni o i reset locali devono impedire che un elemento erediti accidentalmente i valori del padre, conservando però gli override quando si compongono più risultati di `sx`.
3. **Composizione completa.** Shorthand, assi, lati, classi statiche, token e valori custom devono avere risultati prevedibili in entrambi gli ordini. Non concatenare manualmente le sequenze Griffel.
4. **Una sola copia effettiva di Griffel.** `mergeClasses` usa registri di definizioni locali al modulo. Copie duplicate del core possono compromettere la composizione anche se il CSS viene inserito. Servono peer condivisi e una verifica del consumer reale.
5. **Lifecycle di inserimento.** Core inserisce le regole sincronicamente, come il fallback Griffel React 17. Il wrapper deve delegare quel comportamento e verificarlo in mount, aggiornamento, remount e sostituzione del renderer; non introdurre uno spostamento silenzioso negli effetti.
6. **Crescita delle regole.** Una WeakMap non trattiene renderer dismessi, ma l'evizione delle factory non rimuove automaticamente CSS e registri già conservati da Griffel. Documentare e misurare la crescita con valori distinti, senza promettere memoria CSS limitata.
7. **Temi, token e direzione.** I token devono restare riferimenti alle variabili del tema Fluent, con combinazioni semantiche documentate. Verificare i nomi contro il floor dichiarato, non soltanto contro upstream: la versione installata react-theme 9.2.0 usa tokens 1.0.0-alpha.22 e non contiene ogni token delle versioni successive. Variabili e classi statiche devono concordare sulla direzione. Non ereditare variabili di stato interne dal padre al posto del normale colore CSS. Temi high contrast e modalità forced-colors del browser sono prove distinte.
8. **Scrollbar.** L'aspetto nativo dipende da browser, sistema e preferenze utente. La sola ricetta `scrollbar.fluent` applica spessore sottile, thumb a token Fluent e track trasparente. Overflow resta esplicito; gutter, overscroll e smooth scrolling non fanno parte dell'API v1 e non vengono introdotti dalla ricetta.
9. **Overlap degli stati.** Provare hover/active/focus-visible simultanei, anche dentro le query, e passaggi enabled/disabled mentre puntatore o focus restano sul nodo. Una dichiarazione base successiva non prevale automaticamente su una dichiarazione pseudo-state. In particolare, rimuovere i descriptor interattivi quando cambia la condizione applicativa e verificare i relativi reset senza cancellare altri risultati di `sx`.

## Impatto sul bundle

L'incremento è composto da wrapper, catalogo e dipendenze Griffel effettivamente mancanti nel bundle del consumer. La presenza di Griffel in `node_modules` non dimostra che sia già incluso nel download dell'applicazione.

Le variabili permettono di riutilizzare regole delle proprietà per molti valori. Il catalogo cresce soprattutto con proprietà, contesti responsive e stati. Un catalogo unico interamente referenziato da `useSx` può ostacolare il tree shaking anche quando il consumer usa poche utility.

La revisione limita l'asse degli stati a base/hover/active/focus-visible per le proprietà pertinenti. I preset espandono soltanto proprietà già gestite, la tipografia riutilizza il catalogo ufficiale e non si esporta l'intera matrice di token/varianti. Questo riduce la superficie prevista, ma non dimostra ancora un risparmio in byte. Evitare accessi dinamici a grandi registri che obblighino il bundler a conservare tutto il catalogo.

Il piano preferisce export individuali raccolti in namespace ESM, mantenendo la sintassi `width.full` e `typography.body1` senza costruire un grande oggetto pubblico a runtime. L'effetto sul tree shaking deve essere verificato nel bundler SPFx: namespace, nomi estesi e descriptor diretti non garantiscono da soli che ogni membro inutilizzato scompaia. Non rendere un registro completo una dipendenza obbligatoria di `useSx`.

Il build attuale della libreria è `tsc`; non è configurata l'estrazione AOT Griffel. Non promettere precompilazione o CSS estratto solo perché le definizioni sono statiche. Non applicare indiscriminatamente `sideEffects: false`: i servizi PnP esistenti importano moduli di registrazione necessari.

Misure previste: consumer senza la funzione, uso minimo e uso completo, con e senza un costo Fluent già presente. Confrontare JavaScript raw/gzip/brotli, chunk effettivamente utilizzati, CSS inserito, numero di regole e crescita con valori ripetuti/distinti. Il peso del tarball è una misura separata. Nessuna misura o previsione numerica del nuovo bundle è disponibile oggi.

## Sequenza proposta dal piano in chat

| Fase | Risultato atteso prima di proseguire |
| --- | --- |
| 1. Contratto e baseline | Ruoli ridotti, preset documentati, tipografia ufficiale, tre stati, soglie e direzione definiti; baseline del checkout corrente |
| 2. Renderer e variabili | Adapter e factory isolate, coerenza RTL, documenti/salt distinti e riutilizzo verificati |
| 3. Normalizzazione e composizione | Conflitti per proprietà/scope, shorthand e composizione fra chiamate corretti |
| 4. Responsive e stati | Prove browser di soglie, overlap dei tre stati, rimozione condizionale, ereditarietà e isolamento |
| 5. Utility e preset | Famiglie ricorrenti e composizioni cromatiche/typography ufficiale, complete di JSDoc, Markdown, test e scenari |
| 6. Web part e documentazione | Pannello Styles con tutte le capacità osservabili e indici aggiornati |
| 7. Pacchetto e costo | Consumer isolato, identità delle dipendenze e misure comparabili disponibili |
| 8. Gate e revisione | Build, verifiche di repository e limiti del real-host registrati |

L'organizzazione proposta usa `src/helpers/styles/` per conservare i confini attuali del pacchetto e `src/hooks/useSx.ts` per l'hook pubblico. Non è una modifica già applicata. I baseline storici API/package devono restare immutati; i verificatori dovranno consentire soltanto le aggiunte esplicitamente previste.

## Condizioni di completamento

- Sintassi inline in `className`, senza attributi `style` o provider aggiuntivi obbligatori.
- Descriptor diretti tipizzati, nomi pubblici descrittivi e input invalidi rifiutati in modo documentato.
- Renderer, direzione, documenti e istanze isolati correttamente.
- Composizione e responsive dimostrati con prove appropriate, incluse prove browser per gli stili computati.
- Mapping dei ruoli e delle coppie cromatiche verificati nel catalogo supportato; subtle distinto da alternative; normale foreground ereditato quando omesso.
- Tutti i 17 preset tipografici ufficiali riutilizzati senza modificare valori o aggiungere foreground/semantica HTML.
- Solo hover/active/focus-visible nella v1, con comportamento combinato verificato e stati applicativi gestiti tramite condizioni e props ordinarie.
- `container.inlineSize` e `scrollbar.fluent` coerenti con il catalogo del piano; nessuna utility di gutter, overscroll o smooth scrolling pubblicata.
- JSDoc nei sorgenti e nelle dichiarazioni, sezioni Markdown e scenari osservabili per ogni nuova capacità.
- Peer dichiarati e condivisi; public/deep imports del tarball verificati con React 17 e TypeScript 5.3.3.
- Misure di costo e limiti del registro CSS documentati.
- Esecuzione dei gate previsti da [Development](../../docs/DEVELOPMENT.md): build della libreria prima dell'app, `npm run verify`, `npm run verify:package` e verifica del diff.
- Validazione autenticata SharePoint distinta dalle prove locali, secondo [SharePoint validation](../../docs/SHAREPOINT-VALIDATION.md).

La richiesta corrente autorizza la riformulazione e il salvataggio del piano e l'allineamento dell'analisi. Il piano collegato è pronto per la revisione finale prima dell'esecuzione. Non sono autorizzati da questa registrazione commit, push, pubblicazione o modifica dei dati del tenant.

## Riferimenti tecnici

- [Griffel: limitations e valori dinamici](https://griffel.js.org/react/guides/limitations/)
- [Griffel: makeStyles](https://griffel.js.org/react/api/make-styles/)
- [Griffel: mergeClasses](https://griffel.js.org/react/api/merge-classes/)
- [Griffel: createDOMRenderer](https://griffel.js.org/react/api/create-dom-renderer/)
- [Griffel: build optimization](https://griffel.js.org/react/build-optimization/introduction/)
- [Fluent: design tokens](https://fluent2.microsoft.design/design-tokens)
- [Fluent: colori e ruoli](https://fluent2.microsoft.design/color)
- [Fluent: tipografia](https://fluent2.microsoft.design/typography)
- [Preset tipografici ufficiali](https://github.com/microsoft/fluentui/blob/master/packages/tokens/src/global/typographyStyles.ts)
- [Button: appearance e comportamento](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-button/library/src/components/Button/Button.types.ts)
- [Card: appearance e selezione](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-card/library/src/components/Card/Card.types.ts)
- [Badge: appearance e ruoli di colore](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-badge/library/src/components/Badge/Badge.types.ts)
- [FluentProvider: uso del renderer](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-provider/library/src/components/FluentProvider/useFluentProvider.ts)
- [CSS custom properties](https://www.w3.org/TR/css-variables-1/)
- [Container queries](https://developer.mozilla.org/en-US/docs/Web/CSS/Guides/Containment/Container_queries)
- [WebKit: scrollbar-color e differenze di presentazione](https://webkit.org/blog/17640/webkit-features-for-safari-26-2/)

L'ispezione delle versioni installate ha incluso `makeStyles`, renderer/context, registri di composizione, ordinamento delle query e dichiarazioni TypeScript di Griffel, oltre ai contesti e al rendering di FluentProvider. Queste osservazioni devono essere ricontrollate se cambiano le versioni supportate.
