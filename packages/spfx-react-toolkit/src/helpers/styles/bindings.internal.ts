import type { GriffelStyle } from '@griffel/core';
import type { SxBinding, SxState } from './descriptor.internal';
import { sxScopeKey, SxNormalizedDeclaration, SxQuery } from './normalize.internal';

const queryPriority: readonly SxQuery[] = [
  { target: 'base' },
  { target: 'viewport', breakpoint: 480 }, { target: 'container', breakpoint: 480 },
  { target: 'viewport', breakpoint: 640 }, { target: 'container', breakpoint: 640 },
  { target: 'viewport', breakpoint: 1024 }, { target: 'container', breakpoint: 1024 }
];
// Higher states override lower states, independent of mount and CSS insertion order.
const statePriority: readonly (SxState | undefined)[] = [undefined, 'hover', 'active', 'focus-visible'];
function variableName(property: string, query: SxQuery, state?: SxState): `--${string}` {
  return `--apvee-sx-${property}-${sxScopeKey(query, state)}`;
}
/** Native property/scope identity merges with foreign Griffel classes; every scope reads the same priority chain. */
export function createBindingStyle(binding: SxBinding, query: SxQuery, state?: SxState): GriffelStyle {
  let value = binding.fallback;
  const reset: Record<`--${string}`, string> = {};
  for (const state of statePriority) {
    for (const query of queryPriority) {
      const variable = variableName(binding.property, query, state);
      value = `var(${variable}, ${value})`;
      // CSS-wide initial is the guaranteed-invalid value of an unregistered custom
      // property. Griffel emits .class:where(.class) with specificity 1.
      // Live assignments emit .class.class (specificity 2), so assignments from another
      // sx result still win while inherited private variables are blocked locally.
      reset[variable] = 'initial';
    }
  }
  const hasVendorScrollbar = binding.scrollbarSize !== undefined || binding.scrollbarThumb !== undefined;
  const vendorVariable = variableName(`${binding.property}-vendor`, query, state);
  if (hasVendorScrollbar) {
    // Keep a single native binding in the ordinary Griffel property/scope so a
    // later foreign atomic declaration can replace it. Block inherited vendor
    // choices even while this element's query or interaction scope is inactive.
    value = `var(${vendorVariable}, ${value})`;
    reset[vendorVariable] = 'initial';
  }
  let propertyStyle: GriffelStyle = {
    [binding.property]: value,
    ...(binding.forcedColors === undefined ? {} : {
      '@media (forced-colors: active)': { [binding.property]: binding.forcedColors }
    })
  };
  // Non-auto standard scrollbar properties suppress vendor pseudo-elements in
  // modern Chromium. Reset them only where the vendor recipe can apply, leaving
  // native thin/system colors in other browsers and forced-colors mode.
  if (hasVendorScrollbar) {
    propertyStyle['@supports selector(::-webkit-scrollbar)'] = {
      '@media (forced-colors: none)': {
        '&&&': { [vendorVariable]: 'auto' },
        ...(binding.scrollbarSize === undefined ? {} : {
          '::-webkit-scrollbar': { width: binding.scrollbarSize, height: binding.scrollbarSize }
        }),
        ...(binding.scrollbarThumb === undefined ? {} : {
          '::-webkit-scrollbar-thumb': { backgroundColor: binding.scrollbarThumb, borderRadius: '3px' },
          '::-webkit-scrollbar-track': { backgroundColor: 'transparent' },
          '::-webkit-scrollbar-corner': { backgroundColor: 'transparent' }
        })
      }
    };
  }
  if (state) propertyStyle = { [`:${state}`]: propertyStyle };
  if (query.target === 'container') {
    propertyStyle = { [`@container apvee-sx (min-inline-size: ${query.breakpoint}px)`]: propertyStyle };
  } else if (query.target === 'viewport') {
    propertyStyle = { [`@media (min-width: ${query.breakpoint}px)`]: propertyStyle };
  }
  // Reset all private variables locally even when this native property scope is
  // inactive. Keep the reset below live assignments from independently merged sx.
  return { ...propertyStyle, ':where(&)': reset };
}
/** Scoped assignments are ordinary Griffel declarations, never inline style. */
export function createAssignmentStyle(declaration: SxNormalizedDeclaration): GriffelStyle {
  const { property, value, query, state } = declaration;
  let style: GriffelStyle = { [variableName(property, query, state)]: value };
  // Core 1.19.2 normalizes a leading & twice; &&& emits .class.class.
  style = { [state ? `&&&:${state}` : '&&&']: style };
  if (query.target === 'container') {
    style = { [`@container apvee-sx (min-inline-size: ${query.breakpoint}px)`]: style };
  } else if (query.target === 'viewport') {
    style = { [`@media (min-width: ${query.breakpoint}px)`]: style };
  }
  return style;
}
