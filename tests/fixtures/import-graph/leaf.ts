import type { Shape } from './types';
import { onlyType } from './types';
export const value = (shape: Shape): onlyType => shape;
