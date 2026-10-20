import { createContext } from 'react';
import type { ComponentType, ReactNode } from 'react';
import type { PickerRequestCache } from './picker-request-cache.internal';
import type { PickerValue } from './picker-values.types';

/** A mounted native field host supplies its owner/document lease; ordinary React controls own their cache. */
export const PickerRuntimeContext = createContext<PickerRequestCache | undefined>(undefined);
/** Native fields retain reference values even when no visual model exists; never persisted by this context. */
export const PickerSelectedValueContext = createContext<ReadonlyMap<string, PickerValue> | undefined>(undefined);
/** Authoritative native reset epoch; invalidates actions without restarting discovery or the host. */
export const PickerAuthoritativeEpochContext = createContext<string | undefined>(undefined);

/** Native hosts supply the optional mounted presentation boundary without coupling React controls to it. */
export interface PickerControlBoundaryProps { readonly pending: boolean; readonly children: ReactNode; }
export const PickerControlBoundaryContext = createContext<ComponentType<PickerControlBoundaryProps> | undefined>(undefined);
