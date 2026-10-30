const CONTAINER_SIZE_BREAKPOINTS = {
  small: 480,
  medium: 640,
  large: 1024,
  xLarge: 1366,
  xxLarge: 1920,
  xxxLarge: Infinity,
} as const;

/**
 * Classifies container width with exclusive upper boundaries: small <480, medium <640, large <1024, xLarge <1366, xxLarge <1920, otherwise xxxLarge. No DOM measurement or input validation is performed.
 *
 * @param width - Container width in CSS pixels; use a finite, nonnegative measurement.
 * @returns Width category for the supplied value.
 *
 * @example
 * ```ts
 * const category = getSPFxContainerSize(720); // 'large'
 * ```
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
