/**
 * Exclusive container-width categories from caller-supplied measurements.
 *
 * @packageDocumentation
 * @docPage helpers/container
 */
const CONTAINER_SIZE_BREAKPOINTS = {
  small: 480,
  medium: 640,
  large: 1024,
  xLarge: 1366,
  xxLarge: 1920,
  xxxLarge: Infinity,
} as const;

/**
 * Classifies container width with exclusive upper boundaries: small `below 480`, medium `below 640`, large `below 1024`, xLarge `below 1366`, xxLarge `below 1920`, otherwise xxxLarge. No DOM measurement or input validation is performed. Negative widths classify as small; NaN and positive infinity fall through to xxxLarge. CSS style scopes use their own inclusive thresholds.
 *
 * @param width - Container width in CSS pixels; use a finite, nonnegative measurement.
 * @returns Width category for the supplied value.
 *
 * @example
 * ```ts
 * import { getSPFxContainerSize } from '@apvee/spfx-react-toolkit';
 *
 * const category = getSPFxContainerSize(720); // 'large'
 * ```
 * @useWhen You already have a CSS-pixel width and need the toolkit's exclusive container-size categories.
 */
export function getSPFxContainerSize(
  width: number
): 'small' | 'medium' | 'large' | 'xLarge' | 'xxLarge' | 'xxxLarge' {
  if (width < CONTAINER_SIZE_BREAKPOINTS.small) {
    return 'small';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.medium) {
    return 'medium';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.large) {
    return 'large';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.xLarge) {
    return 'xLarge';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.xxLarge) {
    return 'xxLarge';
  }
  return 'xxxLarge';
}
