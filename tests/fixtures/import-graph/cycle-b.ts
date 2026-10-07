import { a } from './cycle-a';
export function b(): string { return a(); }
