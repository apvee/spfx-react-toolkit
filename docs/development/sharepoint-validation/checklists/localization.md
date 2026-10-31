# Localization and UI cultures

[Validation index](../INDEX.md) · [Setup](../setup.md)

Follow the [localization guide](../../../guides/localization.md) before running these procedures. Local tests and packed manifests do not establish authenticated loader behavior.

For host validation, register all eight resources with the consumer postinstall, build the library before the sample and restart serve. With UI Italian and regional English, verify toolbar/UI labels are Italian while DataGrid/RelativeTime formatting follows their configured locales. Add a second web part and change its dictionary; confirm isolation, retained form/editor state and no configuration-triggered content publication. Inspect network resource URLs for only used families and lazy editor behavior, then repeat with an unsupported UI culture and record actual fallback. Open FileFolderPicker and distinguish toolkit outer labels from Microsoft's hosted iframe language. Record deployed asset paths, manifest IDs, chosen cultures and request status. No deployment or tenant configuration change is authorized merely by these build checks.
