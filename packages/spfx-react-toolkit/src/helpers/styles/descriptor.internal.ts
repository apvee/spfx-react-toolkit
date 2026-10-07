import { sxDescriptorBrand, SxBaseDescriptor, SxDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';

/** Finite normalized longhand properties owned by the v1 catalog. */
export type SxProperty =
  | 'display' | 'flexDirection' | 'flexWrap' | 'flexGrow' | 'flexShrink'
  | 'gridTemplateColumns' | 'alignItems' | 'justifyContent' | 'alignSelf'
  | 'rowGap' | 'columnGap'
  | 'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  | 'paddingInlineStart' | 'paddingInlineEnd' | 'paddingBlockStart' | 'paddingBlockEnd'
  | 'marginTop' | 'marginRight' | 'marginBottom' | 'marginLeft'
  | 'marginInlineStart' | 'marginInlineEnd' | 'marginBlockStart' | 'marginBlockEnd'
  | 'width' | 'height' | 'minWidth' | 'minHeight' | 'maxWidth' | 'maxHeight'
  | 'color' | 'backgroundColor' | 'fontFamily' | 'fontSize' | 'fontWeight' | 'lineHeight'
  | 'textAlign' | 'whiteSpace' | 'textOverflow'
  | 'borderTopWidth' | 'borderRightWidth' | 'borderBottomWidth' | 'borderLeftWidth'
  | 'borderTopStyle' | 'borderRightStyle' | 'borderBottomStyle' | 'borderLeftStyle'
  | 'borderTopColor' | 'borderRightColor' | 'borderBottomColor' | 'borderLeftColor'
  | 'borderTopLeftRadius' | 'borderTopRightRadius' | 'borderBottomLeftRadius' | 'borderBottomRightRadius'
  | 'boxShadow' | 'overflowX' | 'overflowY' | 'scrollbarWidth' | 'scrollbarColor'
  | 'containerType' | 'containerName';

/** Property binding's native/base fallback before any scoped value is active. */
export interface SxBinding {
  readonly property: SxProperty;
  readonly fallback: string;
  /** Optional native value used when the forced-colors media feature is active. */
  readonly forcedColors?: string;
  /** Finite vendor sizing used by the Fluent scrollbar recipe. */
  readonly scrollbarSize?: '6px';
  /** Single Fluent thumb color, separate from the standard thumb/track pair. */
  readonly scrollbarThumb?: string;
}
export interface SxDeclarationData extends SxBaseDescriptor {
  readonly kind: 'declaration';
  readonly property: SxProperty;
  readonly value: string;
  readonly binding: SxBinding;
}
export interface SxRecipeData extends SxBaseDescriptor {
  readonly kind: 'recipe';
  readonly name: string;
  readonly declarations: readonly SxBaseDescriptor[];
}
export type SxBaseData = SxDeclarationData | SxRecipeData;
export type SxState = 'hover' | 'active' | 'focus-visible';
export type SxBreakpoint = 480 | 640 | 1024;
export type SxQueryTarget = 'container' | 'viewport';
export interface SxStateData extends SxStateDescriptor {
  readonly kind: 'state';
  readonly state: SxState;
  readonly inputs: readonly SxBaseDescriptor[];
}
export interface SxResponsiveData extends SxResponsiveDescriptor {
  readonly kind: 'responsive';
  readonly target: SxQueryTarget;
  readonly breakpoint: SxBreakpoint;
  readonly inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[];
}
export type SxDescriptorData = SxBaseData | SxStateData | SxResponsiveData;

/** Creates data only: no renderer, CSS insertion, or DOM access. */
export function createDeclaration(property: SxProperty, value: string, binding: SxBinding): SxBaseDescriptor {
  const descriptor: SxDeclarationData = {
    [sxDescriptorBrand]: 'base', kind: 'declaration', property, value,
    binding: Object.freeze({ ...binding })
  };
  return Object.freeze(descriptor);
}
/** Copies and freezes recipe inputs, preserving declaration order. */
export function createRecipe(name: string, declarations: readonly SxBaseDescriptor[]): SxBaseDescriptor {
  const descriptor: SxRecipeData = {
    [sxDescriptorBrand]: 'base', kind: 'recipe', name, declarations: Object.freeze([...declarations])
  };
  return Object.freeze(descriptor);
}
/** Reads factory-owned data behind the opaque public contract. */
export function getDescriptorData(descriptor: SxDescriptor): SxDescriptorData {
  // Every public descriptor is branded by a toolkit factory with the corresponding data shape.
  return descriptor as SxDescriptorData;
}
/** Serializes finite, nonnegative CSS pixel lengths, including explicit zero. */
export function toPixels(value: number): string {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError('Pixel lengths must be finite nonnegative numbers.');
  }
  return `${value}px`;
}
