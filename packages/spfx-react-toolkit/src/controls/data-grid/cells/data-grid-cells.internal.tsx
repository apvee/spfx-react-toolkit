import * as React from 'react';
import { mergeClasses } from '@griffel/react';
import { AvatarGroup, AvatarGroupItem, AvatarGroupPopover, Badge, Button, Link, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Persona, ProgressBar, RatingDisplay, Spinner } from '@fluentui/react-components';
import { CheckmarkCircle20Regular, DismissCircle20Regular, MoreHorizontal20Regular, Star20Filled } from '@fluentui/react-icons';
import type { DataGridCellRenderContext } from '../data-grid-cell.types';
import type { ResolvedDataGridColumn, ResolvedDataGridPerson } from '../column-normalization.internal';
import type { DataGridAction } from '../data-grid-column.types';
import { dataGridIntlLocale, formatDataGridValue } from '../data-grid-values.internal';
import { useDataGridCellStyles } from '../data-grid.styles.internal';
import { DataGridTagsCell } from './data-grid-tags.internal';
import { DataGridSparklineCell } from './data-grid-sparkline.internal';
function booleanLabel<T extends object>(context: DataGridCellRenderContext<T>, value: boolean): string {
    return value ? (context.column.kind === 'boolean' ? context.column.options.trueLabel : undefined) ?? context.strings.yes
        : (context.column.kind === 'boolean' ? context.column.options.falseLabel : undefined) ?? context.strings.no;
}
export function getDataGridCellTooltip<T extends object>(context: DataGridCellRenderContext<T>): string | undefined {
    if (context.value.status === 'invalid')
        return context.column.staticTooltip ?? context.strings.invalidValue;
    const configured = context.column.tooltip(context.value.status === 'missing' ? undefined : context.value.raw, context.item);
    if (configured !== undefined)
        return configured;
    if (context.value.status === 'missing')
        return context.column.kind === 'sparkline' ? context.strings.noObservations : undefined;
    if (context.value.kind === 'boolean')
        return booleanLabel(context, context.value.value);
    const label = formatDataGridValue(context.column, context.value, context.locale, context.item);
    return label === '' ? undefined : label;
}
export function isInteractiveDataGridColumn<T extends object>(column: ResolvedDataGridColumn<T>): boolean {
    return column.kind === 'link' || column.kind === 'people' || column.kind === 'tags' || column.kind === 'actions' || column.kind === 'custom' || column.kind === 'sparkline';
}
function NativePerson<T extends object>({ person, column }: {
    person: ResolvedDataGridPerson;
    column: Extract<ResolvedDataGridColumn<T>, {
        kind: 'person' | 'people';
    }>;
}): React.ReactElement {
    return <Persona {...column.options.personaProps} textAlignment={column.options.personaProps?.textAlignment ?? 'center'} name={person.name} secondaryText={person.secondaryText} avatar={{ ...person.avatarProps, name: person.name }} presence={person.presenceProps}/>;
}
function DataGridPeopleCell<T extends object>({ people, column }: {
    people: readonly ResolvedDataGridPerson[];
    column: Extract<ResolvedDataGridColumn<T>, {
        kind: 'people';
    }>;
}): React.ReactElement {
    const maximum = column.options.maxVisible ?? 3;
    const visible = people.slice(0, maximum), hidden = people.slice(maximum);
    const avatar = (person: ResolvedDataGridPerson): React.ReactElement => <AvatarGroupItem key={`${typeof person.id}:${person.id}`} {...person.avatarProps} name={person.name}/>;
    return <AvatarGroup size={28} layout="stack">{visible.map(avatar)}{hidden.length > 0 &&
            <AvatarGroupPopover count={hidden.length} triggerButton={{ 'aria-label': hidden.map(person => person.name).join(', ') }}>
      {/* AvatarGroupItem owns the avatar; its overflow label must contain text only. */}
      {hidden.map(person => <AvatarGroupItem key={`${typeof person.id}:${person.id}`} {...person.avatarProps}
        name={person.name} badge={person.avatarProps?.badge ?? person.presenceProps}
        overflowLabel={<Persona {...column.options.personaProps} textAlignment={column.options.personaProps?.textAlignment ?? 'center'} presenceOnly
          name={person.name} secondaryText={person.secondaryText}/>}/>)}
    </AvatarGroupPopover>}
  </AvatarGroup>;
}
function actionFlag<T>(flag: boolean | ((item: T) => boolean) | undefined, item: T): boolean { return typeof flag === 'function' ? flag(item) : flag ?? false; }
function DataGridActionsCell<T extends object>({ context, actions }: {
    context: DataGridCellRenderContext<T>;
    actions: readonly DataGridAction<T>[];
}): React.ReactElement {
    const styles = useDataGridCellStyles();
    const visible = actions.filter(action => action.visible === undefined || actionFlag(action.visible, context.item));
    const inline = visible.filter(action => action.placement === 'inline'), menu = visible.filter(action => action.placement === 'menu');
    const click = (action: DataGridAction<T>): React.MouseEventHandler<HTMLElement> => (event: React.MouseEvent<HTMLElement>): void => {
        event.stopPropagation();
        if (!actionFlag(action.disabled, context.item) && !actionFlag(action.loading, context.item))
            action.onClick(event, { item: context.item, rowId: context.rowId, actionId: action.id });
    };
    return <div className={styles.actions} data-grid-actions="">
    {inline.map(action => {
            const loading = actionFlag(action.loading, context.item);
            return <Button key={action.id} size={context.size === 'small' ? 'small' : 'medium'} appearance="subtle" icon={loading ? <Spinner size="tiny"/> : action.icon === undefined ? undefined : <span>{action.icon}</span>} disabled={loading || actionFlag(action.disabled, context.item)} aria-busy={loading || undefined} onClick={click(action)}>{action.label}</Button>;
        })}
    {menu.length > 0 && <Menu><MenuTrigger disableButtonEnhancement><Button appearance="subtle" icon={<MoreHorizontal20Regular />} aria-label={context.strings.moreActions}/></MenuTrigger>
      <MenuPopover><MenuList>{menu.map(action => {
                const loading = actionFlag(action.loading, context.item);
                return <MenuItem key={action.id} icon={loading ? <Spinner size="tiny"/> : action.icon === undefined ? undefined : <span>{action.icon}</span>} disabled={loading || actionFlag(action.disabled, context.item)} aria-busy={loading || undefined} onClick={click(action)}>{action.label}</MenuItem>;
            })}</MenuList></MenuPopover></Menu>}
  </div>;
}
function DataGridNativeCell<T extends object>({ context }: {
    context: DataGridCellRenderContext<T>;
}): React.ReactElement {
    const { column, value, item } = context;
    const styles = useDataGridCellStyles();
    if (value.status !== 'valid')
        return <span className={styles.text}>{value.status === 'invalid' ? column.options.invalidPlaceholder ?? context.strings.invalidValue : column.options.placeholder ?? (column.kind === 'sparkline' ? context.strings.noObservations : '—')}</span>;
    switch (value.kind) {
        case 'boolean': {
            const label = booleanLabel(context, value.value);
            return column.kind === 'boolean' && column.options.display === 'icon'
                ? <span className={styles.cluster} role="img" aria-label={label}>{value.value ? <CheckmarkCircle20Regular aria-hidden="true"/> : <DismissCircle20Regular aria-hidden="true"/>}</span>
                : <span className={styles.text}>{label}</span>;
        }
        case 'badge': {
            const { label, icon, ...props } = value.value;
            return <Badge {...props} appearance={props.appearance ?? 'tint'} icon={icon === undefined ? undefined : <span>{icon}</span>}>{label}</Badge>;
        }
        case 'icon': return <span role="img" aria-label={value.value.label}><span aria-hidden="true">{value.value.icon}</span></span>;
        case 'link': {
            if (column.kind !== 'link')
                return <></>;
            const href = typeof column.options.href === 'function' ? column.options.href(item, value.value) : column.options.href;
            if (href === undefined)
                return <span className={styles.text}>{value.value}</span>;
            return <Link {...column.options.linkProps} href={href} className={mergeClasses(styles.text, column.options.linkProps?.className)} onClick={event => {
                    event.stopPropagation();
                    column.options.onNavigate?.(event, { item, rowId: context.rowId, value: value.value, href });
                }}>{value.value}</Link>;
        }
        case 'person': return column.kind === 'person' ? <NativePerson person={value.value} column={column}/> : <></>;
        case 'people': return column.kind === 'people' ? <DataGridPeopleCell people={value.value} column={column}/> : <></>;
        case 'tags': return column.kind === 'tags' ? <DataGridTagsCell tags={value.value} maxVisible={column.options.maxVisible} strings={context.strings}/> : <></>;
        case 'progress': {
            if (column.kind !== 'progress')
                return <></>;
            const maximum = column.options.max ?? 1;
            const label = formatDataGridValue(column, value, context.locale, item);
            return <span className={styles.cluster}><ProgressBar {...column.options.progressBarProps} className={styles.grow} value={value.value} max={maximum} aria-label={label} aria-valuemin={0} aria-valuemax={maximum} aria-valuenow={value.value}/>{column.options.showValue && <span>{label}</span>}</span>;
        }
        case 'rating': {
            if (column.kind !== 'rating')
                return <></>;
            const count = typeof column.options.count === 'function' ? column.options.count({ item, value: value.value }) : column.options.count;
            if (count !== undefined && (!Number.isInteger(count) || count < 0))
                throw new TypeError('DataGrid rating count callback must return a non-negative whole number');
            const maximum = column.options.max ?? 5;
            const numberFormat = new Intl.NumberFormat(dataGridIntlLocale(column.options.locale ?? context.locale));
            const label = `${numberFormat.format(value.value)}/${numberFormat.format(maximum)}${count === undefined ? '' : `, ${context.strings.ratingCount(count, numberFormat.format(count))}`}`;
            // Fluent's slot contract requires null to suppress automatically
            // rendered numeric labels. Their meaning remains in the root label.
            return <RatingDisplay {...column.options.ratingDisplayProps}
                color={column.options.ratingDisplayProps?.color ?? 'marigold'}
                icon={column.options.ratingDisplayProps?.icon ?? Star20Filled}
                size={column.options.ratingDisplayProps?.size ?? 'large'}
                value={value.value} max={maximum} aria-label={label}
                valueText={null} countText={null}/>;
        }
        case 'sparkline': return column.kind === 'sparkline' ? <DataGridSparklineCell context={context} column={column} value={value.value}/> : <></>;
        case 'actions': return column.kind === 'actions' ? <DataGridActionsCell context={context} actions={column.options.actions}/> : <></>;
        case 'custom': return <></>;
        default: return <span className={styles.text}>{formatDataGridValue(column, value, context.locale, item)}</span>;
    }
}
/** Custom JSX is returned untouched; native cells own only presentation of already classified values. */
export function renderDataGridCell<T extends object>(context: DataGridCellRenderContext<T>): React.ReactNode {
    if (context.column.kind === 'custom' && context.value.status === 'valid')
        return context.column.renderCell(context.value.raw, context.item, context.rowId);
    return <DataGridNativeCell context={context}/>;
}
