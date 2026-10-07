import { makeStyles, mergeClasses } from '@griffel/core';
import { createAssignmentStyle, createBindingStyle } from './bindings.internal';
import { getSxRendererCache } from './cache.internal';
import { normalizeSxInputs, sxScopeKey } from './normalize.internal';
import type { SxStyleEnvironment } from './renderer.internal';
import type { SxInput } from './types';

/** Materialize only requested properties/assignments, merging in original segment order. */
export function resolveSxInputs(environment: SxStyleEnvironment, inputs: readonly SxInput[]): string {
  const cache = getSxRendererCache(environment.renderer);
  const classes: string[] = [];
  for (const segment of normalizeSxInputs(inputs)) {
    if (segment.kind === 'class') {
      classes.push(segment.className);
      continue;
    }
    for (const declaration of segment.declarations) {
      const scopeKey = sxScopeKey(declaration.query, declaration.state);
      const bindingKey = JSON.stringify([
        declaration.binding.property, declaration.binding.fallback, declaration.binding.forcedColors, scopeKey
      ]);
      let binding = cache.bindings.get(bindingKey);
      if (!binding) {
        binding = makeStyles({ root: createBindingStyle(declaration.binding, declaration.query, declaration.state) });
        cache.bindings.set(bindingKey, binding);
      }
      const assignmentKey = JSON.stringify([
        declaration.property, scopeKey, declaration.value
      ]);
      let assignment = cache.assignments.get(assignmentKey);
      if (!assignment) {
        assignment = makeStyles({ root: createAssignmentStyle(declaration) });
        cache.assignments.set(assignmentKey, assignment);
      }
      classes.push(binding(environment).root, assignment(environment).root);
    }
  }
  return mergeClasses(...classes);
}
