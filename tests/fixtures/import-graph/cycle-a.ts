import { b } from './cycle-b';
export function a(): string { return b(); }
