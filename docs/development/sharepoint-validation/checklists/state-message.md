# StateMessage

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks are **NOT EXECUTED** until a tenant run records them. Build the library before the app and restart serve, then open the lazy StateMessage panel. See the [control guide](../../../api/controls/state-message.md).

| Check | Expected behavior | Result |
| --- | --- | --- |
| Presets and theme | All five subjects render in light/dark; nearest provider wins; custom/hidden visual overrides preset. | NOT EXECUTED |
| Container sizing | Wide viewport with narrow WebPart stays compact. Medium uses 640 content px, excluding root padding. RTL and long content stay within the host. | NOT EXECUTED |
| Actions and text | Literal heading, Markdown description, native button/anchor, disabled behavior and inspect-ref observations work. Artwork/actions remain outside status/alert text region. | NOT EXECUTED |
| Multiple instances | Each artwork gradient ID is unique; theme/preset changes preserve text/action elements and focus. | NOT EXECUTED |
| Keyboard and forced colors | Actions retain visible focus/legibility and SVGs remain decorative. Actual screen-reader announcements require separate assistive-technology observation. | NOT EXECUTED |
