# Pagination

[Validation index](../INDEX.md) · [Setup](../setup.md)

Status: **NOT EXECUTED** in an authenticated SharePoint/Teams host. Use the lazy **Pagination** panel after building the library and app; local tests and browser fixtures cannot establish these host results.

- In known-total mode, navigate first/previous/next/last and verify displayed rows and consumer live summary. Check an empty filter, a final partial page and total reduction after a later page. Filter/page-size/total changes reset the application to initial results.
- In unknown-total and cursor modes, verify previous/next capabilities and absence of unsupported extreme actions. In load-more mode, verify appended results and disabled exhaustion.
- Enable manual completion. Request once, verify existing rows remain, loading disables navigation, spinner is labeled and the navigation is busy; complete explicitly and verify exactly the requested rows appear. Compare navigation positions before/during loading.
- Check keyboard Tab/Enter/Space, accessible navigation/group/action labels, Italian overrides and English blank fallback with the host screen reader. Result announcements belong to the consumer summary. No unexpected focus movement should occur.
- Compare light/dark, host RTL icon mirroring, medium/compact sizes and 320px navigation. Confirm wrapping without clipping. Changing one instance must not affect the other.
- Inspect native nav ref and hidden behavior. Navigation clicks must leave form submission count unchanged; only Submit enclosing form increments it. Check retained behavior in SharePoint page layouts and Teams/iframe embedding.

Record host/browser/screen-reader versions and observations separately from local verification.
