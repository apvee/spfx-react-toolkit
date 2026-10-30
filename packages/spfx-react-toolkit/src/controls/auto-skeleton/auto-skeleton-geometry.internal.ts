/** One normalized, overlay-local shape derived from existing rendered HTML. */
export interface AutoSkeletonBone {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly radius: number;
  readonly shape: 'circle' | 'rectangle';
}

/** Configuration already validated by the owning control when it changes. */
export interface AutoSkeletonScanOptions {
  readonly minSize: number;
  readonly ignoreSelectors: readonly string[];
}

/**
 * Validate caller configuration without scanning or modifying the content.
 * Selector parsing uses the content's document, including when no nodes match.
 * @throws {RangeError} When minSize is not positive and finite.
 * @throws {SyntaxError} When an ignored selector is malformed.
 * @internal
 */
export function validateAutoSkeletonOptions(
  content: HTMLElement,
  minSize: number,
  ignoreSelectors: readonly string[]
): void {
  if (!Number.isFinite(minSize) || minSize <= 0) {
    throw new RangeError('AutoSkeleton minSize must be a positive finite number.');
  }
  for (const selector of ignoreSelectors) {
    try {
      content.ownerDocument.querySelector(selector);
    } catch {
      throw new SyntaxError(`AutoSkeleton ignoreSelectors contains an invalid CSS selector: ${selector}`);
    }
  }
}

function isHidden(style: CSSStyleDeclaration, ignoreOpacity = false): boolean {
  return style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse' ||
    (!ignoreOpacity && style.opacity !== '' && Number(style.opacity) === 0);
}

/** A single radius approximates irregular corners conservatively. */
function readRadius(style: CSSStyleDeclaration, width: number, height: number): number {
  const corners = [style.borderTopLeftRadius, style.borderTopRightRadius,
    style.borderBottomRightRadius, style.borderBottomLeftRadius];
  const limit = Math.min(width, height) / 2;
  let radius = limit;
  for (const corner of corners) {
    const parts = (corner || style.borderRadius || '0').trim().split(/\s+/);
    const horizontal = Number.parseFloat(parts[0]);
    const vertical = Number.parseFloat(parts[1] ?? parts[0]);
    const x = parts[0].endsWith('%') ? horizontal * width / 100 : horizontal;
    const y = (parts[1] ?? parts[0]).endsWith('%') ? vertical * height / 100 : vertical;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return 0;
    radius = Math.min(radius, Math.max(0, x), Math.max(0, y));
  }
  return radius;
}

/**
 * Read the content's HTML geometry without DOM writes, caches or subscriptions.
 * The owned wrapper never generates a bone; its loading opacity is intentionally
 * ignored. Authored hiding on that wrapper, the root, ancestors and descendants
 * remains effective. Forced shapes terminate traversal of their subtree.
 * Text uses real Range fragments instead of estimated line counts.
 *
 * Coordinates use the root's padding-box origin and scrolling content space.
 * Scaled/rotated roots or ancestors, nested clipping/scrolling, shadow DOM,
 * portals, pseudo-elements and detailed SVG geometry are outside this scanner.
 * @internal
 */
export function scanAutoSkeleton(
  content: HTMLElement,
  root: HTMLElement,
  options: AutoSkeletonScanOptions
): readonly AutoSkeletonBone[] {
  const document = content.ownerDocument;
  const view = document.defaultView ?? undefined;
  if (!view) return [];
  const getComputedStyle = view.getComputedStyle.bind(view);

  // Loading opacity belongs only to this wrapper, not to authored descendants.
  if (isHidden(getComputedStyle(content), true)) return [];
  let ancestor: Element | undefined = content.parentElement ?? undefined;
  while (ancestor) {
    if (isHidden(getComputedStyle(ancestor))) return [];
    ancestor = ancestor.parentElement ?? undefined;
  }

  const rootRect = root.getBoundingClientRect();
  const originX = rootRect.left + root.clientLeft - root.scrollLeft;
  const originY = rootRect.top + root.clientTop - root.scrollTop;
  if (![originX, originY, rootRect.width, rootRect.height].every(Number.isFinite)) return [];

  const bones: AutoSkeletonBone[] = [];
  const seen = new Set<string>();
  const range = document.createRange();

  function addBone(rect: DOMRect, shape: AutoSkeletonBone['shape'], radius: number, inset = 0): void {
    if (![rect.left, rect.top, rect.width, rect.height, radius].every(Number.isFinite)) return;
    let width = Math.round(rect.width);
    let height = Math.round(rect.height);
    if (width <= 0 || height <= 0 || width < options.minSize || height < options.minSize) return;
    // Filter measured size first; text spacing must not discard eligible fragments.
    const insetX = Math.min(inset, Math.floor((width - 1) / 2));
    const insetY = Math.min(inset, Math.floor((height - 1) / 2));
    const x = Math.round(rect.left - originX) + insetX;
    const y = Math.round(rect.top - originY) + insetY;
    width -= insetX * 2;
    height -= insetY * 2;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    const normalizedRadius = shape === 'circle'
      ? Math.min(width, height) / 2
      : Math.min(Math.round(Math.max(0, radius)), Math.min(width, height) / 2);
    const key = `${x}:${y}:${width}:${height}:${normalizedRadius}:${shape}`;
    if (seen.has(key)) return;
    seen.add(key);
    bones.push({ x, y, width, height, radius: normalizedRadius, shape });
  }

  function visit(node: Node): void {
    if (node.nodeType === 3) {
      if (!node.textContent?.trim()) return;
      range.selectNodeContents(node);
      const rectangles = range.getClientRects();
      for (let index = 0; index < rectangles.length; index++) {
        addBone(rectangles[index], 'rectangle', 0, 2);
      }
      return;
    }
    if (node.nodeType !== 1) return;
    // nodeType establishes Element without depending on another realm's globals.
    const element = node as Element;
    if (element.namespaceURI !== 'http://www.w3.org/1999/xhtml' ||
      element.hasAttribute('data-skeleton-ignore') ||
      options.ignoreSelectors.some(selector => element.matches(selector))) return;
    const style = getComputedStyle(element);
    if (isHidden(style)) return;
    const marker = element.getAttribute('data-skeleton-shape');
    const forced = marker === 'circle' || marker === 'rectangle' ? marker : undefined;
    const hasContent = element.children.length > 0 ||
      Array.from(element.childNodes).some(child => child.nodeType === 3 && !!child.textContent?.trim());
    if (forced || (!hasContent && !element.hasAttribute('data-skeleton-container'))) {
      const rect = element.getBoundingClientRect();
      const radius = readRadius(style, rect.width, rect.height);
      const shape = forced ?? (Math.abs(rect.width - rect.height) <= 2 &&
        radius >= Math.min(rect.width, rect.height) * 0.45 ? 'circle' : 'rectangle');
      addBone(rect, shape, radius);
      return;
    }
    for (const child of Array.from(element.childNodes)) visit(child);
  }

  for (const child of Array.from(content.childNodes)) visit(child);
  return bones;
}
