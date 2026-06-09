const CONTAINER_SIZE_BREAKPOINTS = {
  small: 480,
  medium: 640,
  large: 1024,
  xLarge: 1366,
  xxLarge: 1920,
  xxxLarge: Infinity,
} as const;

/**
 * Gets the Fluent UI 9 aligned SPFx container size category from width.
 *
 * @param width - Container width in pixels
 * @returns Container size category
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
