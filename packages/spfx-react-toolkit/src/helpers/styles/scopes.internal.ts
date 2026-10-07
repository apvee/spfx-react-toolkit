import { sxDescriptorBrand, SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { SxBreakpoint, SxQueryTarget, SxState, SxStateData, SxResponsiveData } from './descriptor.internal';

export function createState(state: SxState, inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  const descriptor: SxStateData = {
    [sxDescriptorBrand]: 'state', kind: 'state', state, inputs: Object.freeze([...inputs])
  };
  return Object.freeze(descriptor);
}
export function createResponsive(target: SxQueryTarget, breakpoint: SxBreakpoint,
  inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  const descriptor: SxResponsiveData = {
    [sxDescriptorBrand]: 'responsive', kind: 'responsive', target, breakpoint, inputs: Object.freeze([...inputs])
  };
  return Object.freeze(descriptor);
}
