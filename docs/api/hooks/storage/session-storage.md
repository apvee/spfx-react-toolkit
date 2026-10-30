# useSPFxSessionStorage

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Temporary storage that persists only for the current browser session/tab. It has the same default, removal and best-effort behavior as [local storage](local-storage.md#default-and-persistence-behavior).

## Signature

```typescript
function useSPFxSessionStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T>
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Fallback when the current key is missing or unreadable |

## Returns

An object containing:

- `value`: Current value of type `T`
- `setValue`: Setter accepting a value or updater function
- `remove`: Removes the stored key and resets `value` to the current default

## Example: Form Draft

```tsx
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';

interface FormData {
  title: string;
  description: string;
  category: string;
}

function FormWithDraft() {
  const { value: draft, setValue: setDraft } = useSPFxSessionStorage<FormData>('formDraft', {
    title: '',
    description: '',
    category: ''
  });

  const handleChange = (field: keyof FormData, value: string) => {
    setDraft(prev => ({ ...prev, [field]: value }));
  };

  const handleSubmit = async () => {
    // Submit form...
    setDraft({ title: '', description: '', category: '' }); // Clear draft
  };

  return (
    <form onSubmit={e => { e.preventDefault(); handleSubmit(); }}>
      <input
        value={draft.title}
        onChange={e => handleChange('title', e.target.value)}
        placeholder="Title"
      />
      <textarea
        value={draft.description}
        onChange={e => handleChange('description', e.target.value)}
        placeholder="Description"
      />
      <button type="submit">Submit</button>
      <p className="hint">Draft auto-saved for this session</p>
    </form>
  );
}
```

## Example: Wizard State

```tsx
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';

interface WizardState {
  currentStep: number;
  completedSteps: number[];
  data: Record<string, unknown>;
}

function MultiStepWizard() {
  const { value: wizard, setValue: setWizard } = useSPFxSessionStorage<WizardState>('wizard', {
    currentStep: 0,
    completedSteps: [],
    data: {}
  });

  const goToStep = (step: number) => {
    setWizard(prev => ({ ...prev, currentStep: step }));
  };

  const completeStep = (stepData: Record<string, unknown>) => {
    setWizard(prev => ({
      currentStep: prev.currentStep + 1,
      completedSteps: [...prev.completedSteps, prev.currentStep],
      data: { ...prev.data, ...stepData }
    }));
  };

  return (
    <div className="wizard">
      <nav>
        {[0, 1, 2, 3].map(step => (
          <button
            key={step}
            onClick={() => goToStep(step)}
            disabled={!wizard.completedSteps.includes(step - 1) && step !== 0}
            className={wizard.currentStep === step ? 'active' : ''}
          >
            Step {step + 1}
          </button>
        ))}
      </nav>
      <main>
        {/* Render current step */}
      </main>
    </div>
  );
}
```

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts)

---
