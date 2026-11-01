/** Peer contract must not replace the toolkit alias comment. */
export declare function peerCallback<T extends (...args: never[]) => unknown>(callback: T): T;
/** Peer base contract stays opaque. */
export interface PeerBase { peerOnly: string; }
