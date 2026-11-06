/**
 * Equal-width grid columns with minmax(0, 1fr) tracks and explicit positive integer counts.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A collection needs a known positive number of equally sized columns that may shrink below intrinsic content width.
 * @example Basic usage
 * ```ts
 * import { grid } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = grid.columns(3);
 * ```
 *
 * @scenario Select positive grid column counts — actual four-item grid and invalid drafts
 * ```tsx
 * import * as React from 'react';
 * import { Field, Input, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, gap, padding, presets, grid, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * type Selection = SxBaseDescriptor | ((value: number) => SxBaseDescriptor);
 * const choices: Readonly<Record<string, Selection>> = grid;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const [value, setValue] = React.useState(2);
 *   const [draft, setDraft] = React.useState('2');
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]); setValue(2); setDraft('2');
 *   }, [resetRevision]);
 *   const selection = choices[member];
 *   const descriptor = typeof selection === 'function' ? selection(value) : selection;
 *   return render({
 *     controls: <><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *       {typeof selection === 'function' && <Field label="Columns"><Input type="number" min={1} step={1} value={draft} onChange={(_, data) => {
 *         setDraft(data.value);
 *         const number = Number(data.value);
 *         if (data.value.trim() && Number.isFinite(number) && number >= 1 && Number.isInteger(number)) setValue(number);
 *       }} onBlur={() => setDraft(String(value))} /></Field>}
 *     </>,
 *     status: <p aria-live="polite">Applied: grid.{member}{typeof selection === 'function' ? `(${value})` : ''}. Invalid drafts preserve the last valid descriptor.</p>,
 *     children: <section role="region" aria-label="grid descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="grid" className={sx(width.px(320), gap.small, descriptor)}>{[1, 2, 3, 4].map(item => <div key={item} className={sx(padding.small, presets.alternative)}>Item {item}</div>)}</div></section>,
 *     diagnostics: <p>Select positive grid column counts; actual four-item grid and invalid drafts. Token descriptors inherit the host theme; Reset restores the first choice and numeric baseline without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select positive grid column counts — actual four-item grid and invalid drafts. Select actual namespace members and finite values; constants omit the numeric field. Reset restores the first choice, numeric draft/value while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Creates a grid with equally sized columns that may shrink below content width.
 * @param count - Positive integer column count.
 * @returns Immutable display and column-template recipe.
 * @throws RangeError - When count is not a finite positive integer.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, grid, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(grid.columns(3), gap.medium)}>Content</div>;
 * }
 * ```
 */
export function columns(count: number): SxBaseDescriptor {
  if (!Number.isInteger(count) || count <= 0) {
    throw new RangeError('Grid column count must be a positive integer.');
  }
  // CSS repeat() requires decimal integer syntax, while Number stringification may use an exponent.
  const integer = count.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 0 });
  return createRecipe('grid.columns', [
    createDeclaration('display', 'grid', { property: 'display', fallback: 'inline' }),
    createDeclaration('gridTemplateColumns', `repeat(${integer}, minmax(0, 1fr))`, {
      property: 'gridTemplateColumns', fallback: 'none'
    })
  ]);
}
