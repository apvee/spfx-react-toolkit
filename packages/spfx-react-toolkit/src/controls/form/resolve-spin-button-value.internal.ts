import type { SpinButtonOnChangeData } from '@fluentui/react-components';
/** Resolves Fluent numeric commits without accepting partial/nonfinite display text. */
// Fluent SpinButton's empty committed numeric domain explicitly uses null.
// eslint-disable-next-line @rushstack/no-new-null
export function resolveSpinButtonValue(data: SpinButtonOnChangeData): number | null {
    if (data.value !== undefined) {
        // eslint-disable-next-line @rushstack/no-new-null
        return typeof data.value === 'number' && Number.isFinite(data.value) ? data.value : null;
    }
    const display = data.displayValue?.trim();
    if (!display) {
        // eslint-disable-next-line @rushstack/no-new-null
        return null;
    }
    const value = Number(display);
    // eslint-disable-next-line @rushstack/no-new-null
    return Number.isFinite(value) ? value : null;
}
