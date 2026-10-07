/**
 * Public typed style descriptors and the canonical class composer.
 * Importing descriptors does not insert CSS. Token styles need theme variables
 * in scope; useSx uses the existing Griffel renderer without an SPFx provider.
 * Re-exports all public descriptor namespaces/scopes and Sx types, plus useSx,
 * from their canonical modules. Root, this facade, /styles/useSx and historical
 * /lib imports share implementations and the active renderer/direction context.
 * Requires shared Griffel/Fluent peers and an ESM-aware package-exports bundler;
 * does not add native Node/CJS runtime support. Existing CSS/cache lifetime and
 * hook lifecycle limitations apply unchanged.
 * @example
 * ```tsx
 * import { useSx, width, padding } from '@apvee/spfx-react-toolkit/styles';
 * function StyledPreview() {
 *   const sx = useSx();
 *   return <div className={sx(width.full, padding.medium)}>Preview</div>;
 * }
 * ```
 */
export * from '../helpers/styles';
export { useSx } from '../hooks/useSx';
