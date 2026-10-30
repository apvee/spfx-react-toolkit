import * as React from 'react';
import { useMergedRefs } from '@fluentui/react-utilities';

function eligible(element: HTMLElement): boolean {
  if (element.matches(':disabled,[aria-disabled="true"]')) return false;
  let ancestor: HTMLElement | undefined = element;
  while (ancestor) {
    if (ancestor.hidden || ancestor.hasAttribute('inert') || ancestor.getAttribute('aria-hidden') === 'true') return false;
    const style = ancestor.ownerDocument.defaultView?.getComputedStyle(ancestor);
    if (style?.display === 'none' || style?.visibility === 'hidden' || style?.visibility === 'collapse') return false;
    ancestor = ancestor.parentElement ?? undefined;
  }
  return true;
}
function clearNativeValidity(element: HTMLElement | undefined): void {
  if (element && 'setCustomValidity' in element && typeof element.setCustomValidity === 'function') element.setCustomValidity('');
}
/** Keeps the public ref on the real group; RHF receives an eligible real native
 * managed primitive target, supporting focus and available native validation with its actual receiver.
 * The owning adapter supplies its primitive selector and field/identity predicate;
 * native inputs remain the default target type.
 */
export function useCompositeFieldFocus<TElement extends HTMLElement, TTarget extends HTMLElement = HTMLInputElement>(publicRef: React.ForwardedRef<TElement>, fieldRef: React.RefCallback<HTMLElement>, nativeSelector: string, acceptTarget: (element: TTarget) => boolean): React.RefCallback<TElement> {
  const rootRef = React.useRef<TElement>();
  const registeredRef = React.useRef<HTMLElement>();
  const capture = React.useCallback<React.RefCallback<TElement>>(element => {
    rootRef.current = element ?? undefined;
  }, []);
  const mergedRef = useMergedRefs(capture, publicRef);
  React.useLayoutEffect(() => () => {
    clearNativeValidity(registeredRef.current);
    registeredRef.current = undefined;
    // React's external ref protocol explicitly uses null for unmount.
    // eslint-disable-next-line @rushstack/no-new-null
    fieldRef(null);
  }, [fieldRef]);
  // Synchronize after every commit: options and native eligibility may change
  // independently of field value. The ref records a DOM target, never validation state.
  React.useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const target = Array.from(root.querySelectorAll<TTarget>(nativeSelector)).find(element => acceptTarget(element) && eligible(element)) ?? root;
    if (registeredRef.current !== target) clearNativeValidity(registeredRef.current);
    registeredRef.current = target;
    fieldRef(target);
  });
  return mergedRef;
}
