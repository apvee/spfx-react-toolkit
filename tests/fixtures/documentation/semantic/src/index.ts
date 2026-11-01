import { peerCallback } from '@fixture/peer';
export type { CallbackOwner } from './member-docs';
export { original, original as renamed } from './local';
/** Local namespace contract. */
export * as metrics from './metrics';
/** Locally documented alias of an external callable. */
export const publicCallback = peerCallback;
interface HiddenResult { value: number; }
/** Callable value returning a non-exported support type. */
export const callable = (value: number): HiddenResult => ({ value });
/** Convert numeric or textual values. */
export function convert(value: string): number;
export function convert(value: number): string;
export function convert(value: string | number): number | string { return typeof value === 'string' ? value.length : String(value); }
interface Base { kept?: string; removed: number; }
/** Branch-sensitive surface options. */
export type SurfaceProps = Omit<Base, 'removed'> & (
  { mode: 'dialog'; trigger?: string; forbidden?: never } |
  { mode: 'popover'; trigger: string; forbidden?: never }
);
const Root = (props: SurfaceProps): string => props.mode;
const Header = (props: { title: string }): string => props.title;
/** Compound callable value. */
export const Compound = Object.assign(Root, {
  /** Mounted header contract. */
  Header,
});
/** Generic result suite. */
export interface Suite<T extends Record<string, unknown>> {
  Root: (props: { values: T }) => string;
  Field: <K extends keyof T>(name: K) => T[K];
}
/** Build a generic suite. */
export declare function makeSuite<T extends Record<string, unknown>>(): Suite<T>;
import type { PeerBase } from '@fixture/peer';
/** Local extension of an opaque peer base. */
export interface ExtendedPeer<T extends string = string> extends PeerBase { local: T; }
const brand: unique symbol = Symbol('brand');
/** Branded opaque token. */
export interface Token { readonly [brand]: true; }
export declare const token: Token;
import * as unitNamespace from './metrics';
/** Return a namespace to exercise checker-generated import-type spelling. */
export const returnNamespace = () => unitNamespace;
