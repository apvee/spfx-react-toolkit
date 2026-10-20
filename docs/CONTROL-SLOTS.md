# Picker presentation slots

Picker presentation regions use the private Fluent slot types and resolver. They accept React shorthand or fixed-tag props objects, including objects whose children are render callbacks. These helpers have no public entrypoint. See the [picker API](./api/controls/pickers.md) for each region's tag, default and formatting contract.

Strings, numbers (including zero), JSX, fragments, arrays, iterables and portals are React shorthand. Booleans and undefined mean absent content. React/Fluent null remains a legacy input; new absence uses undefined. Props objects request their fixed native container even for empty static children. Direct function values are unsupported; use the object's children callback.

The control owns semantic shells, IDs, associations, input behavior, query state, retry and selection callbacks. Consumer props customize only the documented inner container. Forward every supplied prop to the callback Component, including composed class names and refs. Refs target that inner container. Fixed tags cannot be replaced with arbitrary elements. Container raw HTML and feature-owned props are filtered; arbitrary consumer JSX remains opaque and its accessibility belongs to the consumer.

Callbacks can run repeatedly through ordinary React rendering. Returning undefined does not remove a control-owned semantic shell. The resolver never invokes callbacks to infer content presence. Theme, direction and Griffel renderer come from the host; no package renderer is created. Consumer styles are forwarded, while toolkit-authored styling uses classes.

Native Fluent Field props and structural/data callbacks retain their specific contracts. See [Property Pane pickers](./api/property-pane/pickers.md), [development](./DEVELOPMENT.md) and [package imports](./PACKAGE-IMPORTS.md).
