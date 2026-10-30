# Prerequisites and debug setup

[Validation index](INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


Use a test tenant/site you can edit, an authenticated account and a modern page with a Top placeholder. Record tenant/site URLs, account role, Node version, library package version and SPFx 1.21.1. Prepare a disposable `Tasks` list with a Title column and at least two pages of items. Use separate test keys/files; list writes, OneDrive writes, site collection writes and tenant key-value writes change live test data.

From the repository root:

```bash
npm ci
npm run build:library
npm run build:app
npm run verify
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
```

In [app serve config](../../../apps/spfx-react-toolkit-test/config/serve.json), replace `{tenantDomain}` with the real SharePoint host and set `initialPage` to the authenticated `/_layouts/workbench.aspx`; set each configured `pageUrl` to an existing modern page. Keep extension ID `4c2e58cc-1253-43e5-bfa3-b22a07ace279` in its Application Customizer custom action. These config edits are local test setup.

```bash
npm run serve
```

For the named Application Customizer configuration:

```bash
npm run serve -- --config=spFxReactToolkitTest
```

Accept the local debug-script prompt only for your test session. For WebPart checks open the authenticated workbench while serve runs and add **SpFxReactToolkitTest** twice. For Application Customizer checks open the configured modern page via the serve configuration and inspect the Top banner. Record URLs with sensitive query parameters redacted. Rebuild library and restart serve after any library source edit.
