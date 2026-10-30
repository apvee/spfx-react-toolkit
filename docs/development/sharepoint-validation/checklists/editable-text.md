# EditableText standalone region

[Validation index](../INDEX.md) · [Setup](../setup.md)

The [EditableText guide](../../../api/controls/editable-text.md) documents the caller-owned contract. Local controlled-demo DOM tests and compiled library/app builds are separate from authenticated host evidence.

| Scenario | Expected observation | Host status |
| --- | --- | --- |
| Four policies and Read/Edit | H2 becomes p in paragraph policy without property writes; raw source stays H2 until user input. Root is stable and Read ends its editor. | NOT EXECUTED |
| H2, fallback and empty | Undefined HTML uses literal fallback; defined empty remains empty after save/reload. | NOT EXECUTED |
| Persistence and final flush | Caller saves canonical HTML plus search text only after effective input; latest policy/callback governs Edit exit and pending IME flush. | NOT EXECUTED |
| Isolation | Separate fixture has two policies and a Shell title; editing, undo and policy changes affect only their owner. Supply the actual web part instance. | NOT EXECUTED |
| Host boundaries | Keyboard toolbar/shortcuts, screen-reader naming, IME, live theme/RTL and popup/renderer ownerDocument remain correct in SharePoint/Teams/iframe hosts. | NOT EXECUTED |
