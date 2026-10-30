const focusableSelector = 'input:not([type="hidden"]),textarea,select,button,a[href],[tabindex],[contenteditable="true"]';

/** Checks availability through ancestors; geometry is deliberately not used (SSR/test and display:contents). */
function eligible(element: HTMLElement): boolean {
    if (!element.isConnected || element.tabIndex < 0 || element.matches(':disabled')) return false;
    let current: HTMLElement | undefined = element;
    while (current) {
        if (current.hidden || current.hasAttribute('inert') || current.getAttribute('aria-hidden') === 'true' || current.getAttribute('aria-disabled') === 'true') return false;
        const style = current.ownerDocument.defaultView?.getComputedStyle(current);
        if (style?.display === 'none' || style?.visibility === 'hidden' || style?.visibility === 'collapse') return false;
        current = current.parentElement ?? undefined;
    }
    return true;
}
/** Attempts actual native focus and detects targets that reject it rather than stopping discovery early. */
export function focusArrayTarget(target: HTMLElement | undefined): boolean {
    if (!target || !eligible(target)) return false;
    try { target.focus(); }
    catch { return false; }
    return target.ownerDocument.activeElement === target;
}
/** Focuses the first available descendant in DOM order; hidden/inert/disabled and rejecting targets are skipped. */
export function focusArrayContent(root: HTMLElement | undefined): boolean {
    if (!root) return false;
    for (const target of Array.from(root.querySelectorAll<HTMLElement>(focusableSelector))) {
        if (focusArrayTarget(target)) return true;
    }
    return false;
}
