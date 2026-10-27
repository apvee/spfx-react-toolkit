import * as React from 'react';
import { Button, Card, Field, Tooltip, useId } from '@fluentui/react-components';
import { Add20Regular, ArrowUp20Regular, ArrowDown20Regular, Delete20Regular } from '@fluentui/react-icons';
import type { FieldValues } from 'react-hook-form';
import type { FormArrayContext, FormArrayItemValue, FormArrayPath, FormArraySectionProps } from '../form-composition.types';
import { FieldPathScope, useFieldScopeDisabled, useResolvedFieldPath } from '../field-path.internal';
import { useFormRuntimeContext } from '../form-context.internal';
import { FormArray } from './form-array';
import { formResetTargets, requireFormMethodsMetadata, subscribeToFormResets } from '../form-methods.internal';
import { focusArrayContent, focusArrayTarget } from './form-array-focus.internal';
import { assertSectionName, sectionDescriptionIds, useFormSectionsStyles } from './form-sections.styles.internal';
import { GroupText, hasGroupContent, hasGroupDescriptionContent } from '../../group-header.internal';
import { renderContentSlot } from '../../content-slot.internal';
import { useSx } from '../../../hooks/useSx';
import * as flex from '../../../helpers/styles/flex';
import * as gap from '../../../helpers/styles/gap';
import * as minWidth from '../../../helpers/styles/min-width';
import * as marginBlockStart from '../../../helpers/styles/margin-block-start';
import * as container from '../../../helpers/styles/container';
import * as responsive from '../../../helpers/styles/responsive';
import * as alignItems from '../../../helpers/styles/align-items';

type PendingFocus = { before: string; resetVersion: number } & (
    { type: 'add'; index: number }
    | { type: 'remove'; keys: readonly string[]; announcement: string }
    | { type: 'move'; key: string; action: 'up' | 'down'; announcement: string }
);

function ArraySectionContent<TInput extends FieldValues, TPath extends FormArrayPath<TInput>>({ context, section, disabled, resolvedName }: {
    context: FormArrayContext<FormArrayItemValue<TInput, TPath>>;
    section: Omit<FormArraySectionProps<TInput, TPath>, 'name' | 'disabled'>;
    disabled: boolean;
    resolvedName: string;
}): React.ReactElement {
    const { createItem, renderItem, minItems = 0, maxItems, removable = true, reorderable = false, itemLabel, strings: overrides, title, description, appearance = 'filled', className, 'aria-describedby': describedBy, ...native } = section;
    const familyStrings = useFormRuntimeContext().strings.arraySection;
    const strings = {
        add: overrides?.add ?? familyStrings.add,
        empty: overrides?.empty ?? familyStrings.empty,
        remove: overrides?.remove ?? familyStrings.remove,
        moveUp: overrides?.moveUp ?? familyStrings.moveUp,
        moveDown: overrides?.moveDown ?? familyStrings.moveDown,
        added: overrides?.added ?? familyStrings.added,
        removed: overrides?.removed ?? familyStrings.removed,
        moved: overrides?.moved ?? familyStrings.moved,
    };
    const sx = useSx(), styles = useFormSectionsStyles();
    const descriptionId = useId('form-array-description'), errorId = useId('form-array-error');
    const hasDescription = hasGroupDescriptionContent(description);
    const { methods } = useFormRuntimeContext();
    const root = React.useRef<HTMLFieldSetElement>();
    const addButton = React.useRef<HTMLButtonElement | HTMLAnchorElement>();
    const pending = React.useRef<PendingFocus>();
    // A reset can supersede an action during the same bubbling event, before React commits.
    // Observe each existing lifecycle event so a later unrelated reset cannot hide that cancellation.
    React.useLayoutEffect(() => subscribeToFormResets(methods, event => {
        const request = pending.current;
        if (request && event.version > request.resetVersion && formResetTargets(event, resolvedName)) pending.current = undefined;
    }), [methods, resolvedName]);
    const [announcement, setAnnouncement] = React.useState({ text: '', version: 0 });
    const announce = (text: string): void => setAnnouncement(previous => ({ text, version: previous.version + 1 }));
    const itemName = (item: typeof context.items[number]): string => itemLabel?.(item) ?? familyStrings.itemLabel(item.index + 1);
    const keySignature = context.items.map(item => item.key).join('|');
    const itemRoot = (key: string): HTMLElement | undefined => Array.from(root.current?.querySelectorAll<HTMLElement>('[data-form-array-item]') ?? []).find(element => element.getAttribute('data-form-array-item') === key);
    const focusItem = (key: string): boolean => {
        const element = itemRoot(key);
        return focusArrayContent(element?.querySelector<HTMLElement>('[data-form-array-content]') ?? undefined) || focusArrayContent(element);
    };
    React.useLayoutEffect(() => {
        const request = pending.current;
        if (!request || request.before === keySignature) return;
        pending.current = undefined;
        if (request.type === 'add') {
            const item = context.items[request.index];
            if (item) {
                if (!focusItem(item.key)) focusArrayTarget(addButton.current);
                announce(strings.added(itemName(item)));
            }
        }
        else if (request.type === 'remove') {
            if (!request.keys.some(focusItem)) focusArrayTarget(addButton.current);
            announce(request.announcement);
        }
        else {
            const element = itemRoot(request.key);
            if (!focusArrayTarget(element?.querySelector<HTMLButtonElement>(`[data-form-array-action="${request.action}"]`) ?? undefined) && !focusArrayContent(element?.querySelector<HTMLElement>('[data-form-array-actions]') ?? undefined) && !focusItem(request.key)) focusArrayTarget(addButton.current);
            announce(request.announcement);
        }
    });
    const canAdd = !disabled && (maxItems === undefined || context.items.length < maxItems);
    const canRemove = !disabled && context.items.length > minItems;
    const add = (): void => {
        if (!canAdd) return;
        const value = createItem();
        pending.current = { type: 'add', index: context.items.length, before: keySignature, resetVersion: requireFormMethodsMetadata(methods).resetEvent.version };
        context.append(value);
    };
    const remove = (item: typeof context.items[number]): void => {
        if (!canRemove || !removable) return;
        pending.current = { type: 'remove', before: keySignature, resetVersion: requireFormMethodsMetadata(methods).resetEvent.version, keys: [...context.items.slice(item.index + 1), ...context.items.slice(0, item.index).reverse()].map(candidate => candidate.key), announcement: strings.removed(itemName(item)) };
        context.remove(item.index);
    };
    const move = (item: typeof context.items[number], action: 'up' | 'down'): void => {
        const to = item.index + (action === 'up' ? -1 : 1);
        if (disabled || !reorderable || to < 0 || to >= context.items.length) return;
        pending.current = { type: 'move', key: item.key, action, before: keySignature, resetVersion: requireFormMethodsMetadata(methods).resetEvent.version, announcement: strings.moved(itemName(item), to + 1, item.total) };
        context.move(item.index, to);
    };
    return <fieldset {...native} ref={element => { root.current = element ?? undefined; }} disabled={disabled} aria-invalid={context.errors.length > 0 || native['aria-invalid']} aria-describedby={sectionDescriptionIds(describedBy, hasDescription ? descriptionId : undefined, context.errors.length ? errorId : undefined)} className={sx(styles.group, container.inlineSize, minWidth.zero, className)}>
        {hasGroupContent(title) && <legend className={styles.legend}><GroupText>{title}</GroupText></legend>}
        <Card appearance={appearance} focusMode="off" size="medium">
            <Field className={sx(gap.small)} hint={hasGroupContent(description) ? { id: descriptionId, className: sx(marginBlockStart.none), children: renderContentSlot(description, 'div', sx, { shorthand: 'preserve' }) } : undefined} validationState={context.errors.length ? 'error' : 'none'} validationMessage={context.errors.length ? { id: errorId, className: sx(marginBlockStart.none), children: context.errors.join('; ') } : undefined}>
                <div className={sx(flex.column, gap.medium, minWidth.zero)}>
                    {context.items.length === 0 ? <GroupText description>{strings.empty}</GroupText> : <ol className={sx(styles.list, flex.column, gap.medium)}>{context.items.map(item => {
                        const label = itemName(item);
                        return <li key={item.key} data-form-array-item={item.key} aria-label={label} className={sx(styles.item, flex.column, gap.small)}>
                            <div data-form-array-content="" className={sx(flex.column, gap.medium, minWidth.zero)}>{renderItem(item)}</div>
                            {(removable || reorderable) && <div data-form-array-actions="" className={sx(flex.column, gap.small, responsive.small(flex.row, flex.wrap, alignItems.center))}>
                                {reorderable && <><Tooltip content={strings.moveUp(label)} relationship="label"><Button type="button" appearance="subtle" size="small" aria-label={strings.moveUp(label)} icon={<ArrowUp20Regular aria-hidden="true" />} data-form-array-action="up" disabled={disabled || item.isFirst} onClick={() => move(item, 'up')} /></Tooltip><Tooltip content={strings.moveDown(label)} relationship="label"><Button type="button" appearance="subtle" size="small" aria-label={strings.moveDown(label)} icon={<ArrowDown20Regular aria-hidden="true" />} data-form-array-action="down" disabled={disabled || item.isLast} onClick={() => move(item, 'down')} /></Tooltip></>}
                                {removable && <Tooltip content={strings.remove(label)} relationship="label"><Button type="button" appearance="subtle" size="small" aria-label={strings.remove(label)} icon={<Delete20Regular aria-hidden="true" />} data-form-array-action="remove" disabled={!canRemove} onClick={() => remove(item)} /></Tooltip>}
                            </div>}
                        </li>;
                    })}</ol>}
                    <div><Button type="button" icon={<Add20Regular aria-hidden="true" />} ref={element => { addButton.current = element ?? undefined; }} disabled={!canAdd} onClick={add}>{strings.add}</Button></div>
                </div>
            </Field>
        </Card>
        <div aria-live="polite" aria-atomic="true" className={styles.live}>{announcement.text && <span key={announcement.version}>{announcement.text}</span>}</div>
    </fieldset>;
}

/** Named ordered record-array section built on the existing headless scope.
 * Item fields use medium spacing; descriptions/errors use small spacing.
 * Only explicit visual actions initiate focus. Bounds control actions rather than validation/reset data.
 * @param props Relative array path, accessible group name, complete item factory and item renderer.
 * @returns Native fieldset with Fluent surface/actions, ordered items and polite localized status.
 * @throws {RangeError} Invalid integer bounds or maxItems below minItems.
 * @example `<Order.ArraySection name="rows" title="Rows" createItem={() => ({ sku: '' })} renderItem={({ Fields }) => <Fields.Input name="sku" />} />`
 */
export function FormArraySection<TInput extends FieldValues, TPath extends FormArrayPath<TInput>>(props: FormArraySectionProps<TInput, TPath>): React.ReactElement {
    assertSectionName(props.title, props['aria-label'], props['aria-labelledby']);
    const min = props.minItems ?? 0, max = props.maxItems;
    if (!Number.isInteger(min) || min < 0 || (max !== undefined && (!Number.isInteger(max) || max < min))) throw new RangeError('Form.ArraySection bounds must be nonnegative integers with maxItems at least minItems.');
    const runtime = useFormRuntimeContext();
    const scopeDisabled = useFieldScopeDisabled();
    const { name, disabled: ownDisabled, ...section } = props;
    const disabled = runtime.disabled || scopeDisabled || ownDisabled === true;
    const prefix = useResolvedFieldPath('');
    const resolvedName = useResolvedFieldPath(String(name));
    return <FieldPathScope prefix={prefix} disabled={disabled}><FormArray<TInput, TPath> name={name}>{context => <ArraySectionContent<TInput, TPath> context={context} section={section} disabled={disabled} resolvedName={resolvedName} />}</FormArray></FieldPathScope>;
}
