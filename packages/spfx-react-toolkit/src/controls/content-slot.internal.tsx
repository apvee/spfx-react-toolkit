import * as React from 'react';
import { getIntrinsicElementProps, slot } from '@fluentui/react-utilities';
/** Private string-only composition contract; accepts useSx and Griffel mergeClasses. */
export type ContentSlotClassComposer = (...classNames: readonly (string | undefined)[]) => string;

/** Structural props after Fluent resolution; the concrete tag controls allowed DOM attributes. */
type ContentProps = React.HTMLAttributes<HTMLElement> & {
  as?: keyof JSX.IntrinsicElements;
  ref?: React.Ref<HTMLElement>;
  children?: React.ReactNode | ((Component: React.ElementType, props: React.HTMLAttributes<HTMLElement>) => React.ReactNode);
};

type ContentContainerProps = Omit<ContentProps, 'as' | 'children' | 'dangerouslySetInnerHTML'>;

/** Static inspection never invokes opaque consumer render callbacks or React components. */
export type ContentSlotInspection =
  | { readonly kind: 'absent' }
  | { readonly kind: 'shorthand'; readonly children: React.ReactNode }
  | { readonly kind: 'props'; readonly children: React.ReactNode; readonly hasAccessibleName: boolean }
  | { readonly kind: 'render'; readonly hasAccessibleName: boolean };

/** Control-owned container composition, with an explicit shorthand DOM policy. */
export interface RenderContentSlotOptions {
  /** Preserve caller shorthand as-is, or keep an existing fixed-tag wrapper. */
  readonly shorthand: 'preserve' | 'wrap';
  /** Toolkit class before the consumer class. */
  readonly className?: string;
  /** Container defaults, followed by consumer attributes. */
  readonly defaultProps?: ContentContainerProps;
  /** Invariant container attributes applied last; refs are composed. */
  readonly ownedProps?: ContentContainerProps;
  /** Consumer keys owned elsewhere by the control, removed even for JavaScript inputs. */
  readonly omitProps?: readonly string[];
  /** Explicit bounded alternatives; omitted for ordinary fixed-tag content slots. */
  readonly allowedTags?: readonly (keyof JSX.IntrinsicElements)[];
  /** Feature-specific formatting of static children; callbacks remain opaque. */
  readonly mapChildren?: (children: React.ReactNode) => React.ReactNode;
}

function isPortal(value: object): value is React.ReactPortal {
  return '$$typeof' in value && value.$$typeof === Symbol.for('react.portal');
}

function isShorthand(value: unknown): value is React.ReactNode {
  return typeof value === 'string' || typeof value === 'number' || React.isValidElement(value) ||
    (typeof value === 'object' && value !== null && (Symbol.iterator in value || isPortal(value)));
}

function isAbsent(value: unknown): boolean {
  // React/Fluent retain null as a legacy input; new absence is always undefined.
  return value === undefined || value === null || typeof value === 'boolean' || typeof value === 'function';
}

function hasStaticContent(value: React.ReactNode): boolean {
  if (isAbsent(value)) return false;
  if (typeof value === 'string') return value.trim().length > 0;
  if (typeof value === 'number') return true;
  if (React.isValidElement<{ children?: React.ReactNode }>(value)) {
    return value.type === React.Fragment ? hasStaticContent(value.props.children) : true;
  }
  if (typeof value === 'object' && value !== null) {
    if (isPortal(value)) return true;
    if (Symbol.iterator in value) {
      for (const child of value as Iterable<React.ReactNode>) if (hasStaticContent(child)) return true;
    }
  }
  return false;
}

/** Inspect static children/naming before passing content to feature-specific label validators. */
export function inspectContentSlot(value: unknown): ContentSlotInspection {
  if (isAbsent(value)) return { kind: 'absent' };
  if (isShorthand(value)) return { kind: 'shorthand', children: value };
  if (typeof value !== 'object') return { kind: 'absent' };
  // The Fluent resolver returns object props unchanged. This structural view is guarded at render time.
  const props = slot.resolveShorthand(value as ContentProps);
  const hasAccessibleName = [props['aria-label'], props['aria-labelledby']].some(name => typeof name === 'string' && name.trim().length > 0);
  return typeof props.children === 'function'
    ? { kind: 'render', hasAccessibleName }
    : { kind: 'props', children: props.children, hasAccessibleName };
}

/** Request presence: explicit props (including empty/callback props) always request a container. */
export function hasContentSlot(value: unknown): boolean {
  const inspection = inspectContentSlot(value);
  return inspection.kind === 'props' || inspection.kind === 'render' ||
    (inspection.kind === 'shorthand' && hasStaticContent(inspection.children));
}

/** Semantic presence for associations: static content, explicit naming or an opaque callback. */
export function hasContentSlotContent(value: unknown): boolean {
  const inspection = inspectContentSlot(value);
  return inspection.kind === 'render' ||
    (inspection.kind === 'props' && (inspection.hasAccessibleName || hasStaticContent(inspection.children))) ||
    (inspection.kind === 'shorthand' && hasStaticContent(inspection.children));
}

function composeRefs(...refs: readonly (React.Ref<HTMLElement> | undefined)[]): React.Ref<HTMLElement> | undefined {
  // Ref null is React's legacy absence contract, including unmount callback values.
  const present = [...new Set(refs.filter((ref): ref is NonNullable<React.Ref<HTMLElement>> => ref !== undefined && ref !== null))];
  if (present.length === 0) return undefined;
  if (present.length === 1) return present[0];
  return node => {
    for (const ref of present) {
      if (typeof ref === 'function') ref(node);
      // React's RefObject surface is readonly; React itself writes current on mount/unmount.
      else (ref as { current: React.RefObject<HTMLElement>['current'] }).current = node;
    }
  };
}

/** Render through installed Fluent normalization, preserving shorthand and owned outer semantics.
 * No hooks, Markdown, consumer element cloning or module-evaluation runtime registrations. */
export function renderContentSlot(
  value: unknown,
  tag: keyof JSX.IntrinsicElements,
  sx: ContentSlotClassComposer,
  options: RenderContentSlotOptions,
): React.ReactNode {
  if (isAbsent(value)) return undefined;
  const shorthand = isShorthand(value);
  if (shorthand && options.shorthand === 'preserve') return options.mapChildren ? options.mapChildren(value) : value;
  if (!shorthand && typeof value !== 'object') return undefined;
  // Fluent 9 does not classify React portals as shorthand; adapt only that legacy React shape.
  const resolved = slot.resolveShorthand(isPortalCandidate(value) ? { children: value } : value as ContentProps);
  const requestedTag = resolved.as;
  const Component = requestedTag !== undefined && (requestedTag === tag || options.allowedTags?.includes(requestedTag)) ? requestedTag : tag;
  const { children, className: consumerClassName, ref: consumerRef, ...consumerProps } = resolved;
  const { className: defaultClassName, ref: defaultRef, ...defaultProps } = options.defaultProps ?? {};
  const { className: ownedClassName, ref: ownedRef, ...ownedProps } = options.ownedProps ?? {};
  const allowedConsumerProps: Record<string, unknown> = { ...consumerProps };
  for (const key of options.omitProps ?? []) delete allowedConsumerProps[key];
  const candidate: ContentProps = { ...defaultProps, ...allowedConsumerProps, as: Component };
  const nativeProps = getIntrinsicElementProps(Component, candidate, ['as', 'children', 'dangerouslySetInnerHTML']);
  const props = {
    ...nativeProps, ...ownedProps,
    ref: composeRefs(defaultRef, consumerRef, ownedRef),
    className: sx(defaultClassName, options.className, consumerClassName, ownedClassName),
  };
  return typeof children === 'function'
    ? children(Component, props)
    : React.createElement(Component, props, options.mapChildren ? options.mapChildren(children) : children);
}

function isPortalCandidate(value: unknown): value is React.ReactPortal {
  return typeof value === 'object' && value !== null && isPortal(value);
}
