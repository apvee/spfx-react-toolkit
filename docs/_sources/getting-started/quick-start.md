# Build a first SPFx web part

Mount a complete web part with WebPartShell, wire title metadata and Property Pane settings, then unmount its React tree when SPFx disposes the host.

[Documentation home](../INDEX.md) · [Installation](installation.md) · [Provider selection](choosing-a-provider.md)

Use `WebPartShell` in a web part `.tsx` file. It supplies the SPFx and Fluent providers, so its children can use the toolkit hooks:

<!-- toolkit:include symbol="WebPartShell" section="example" title="First web part with metadata and native Property Pane" -->

Initialize these defaults in the web part manifest's `preconfiguredEntries[].properties`. Leave `titleHTML` omitted until an author edits the title:

```json
{
  "title": "My web part",
  "showTitle": true,
  "description": "Built with SPFx React Toolkit."
}
```

### Shell presentation

Choose Card appearance, inner spacing and title placement separately from outer `surface` / `padding`. These are React props; exposing them in the consumer’s Property Pane is optional. {@link WebPartShell} owns the verified defaults, title persistence, host detection, actions and state retention.

For direct provider composition, choose `SPFxWebPartProvider`, `SPFxApplicationCustomizerProvider`, `SPFxFieldCustomizerProvider` or `SPFxListViewCommandSetProvider` for the matching SPFx host. See [Providers](../../api/core/providers.md) and [Introduction](quick-start.md). Generic React controls and utility hooks document their individual provider requirements.

## Direct provider composition

Use this alternative when you own the surrounding layout and only need SPFx context. It does not supply the Shell title or Fluent theme wrapper.

<!-- toolkit:include symbol="SPFxWebPartProvider" section="example" title="Mount and dispose a web part" -->
