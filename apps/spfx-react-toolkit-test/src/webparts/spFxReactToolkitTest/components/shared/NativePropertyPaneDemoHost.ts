import { PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import type { IPropertyPaneConfiguration, IPropertyPaneCustomFieldProps, IPropertyPaneField } from '@microsoft/sp-property-pane';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

/** Private source/host protocol. The initialized WebPart owns the instance and real lifecycle. */
interface NativePropertyPaneEnvironment<P extends object> {
  readonly instance: BaseClientSideWebPart<P>;
  readonly readProperties: () => P;
  readonly refresh: () => void;
}
interface NativePropertyPaneController {
  readonly getConfiguration: () => IPropertyPaneConfiguration;
  readonly fieldChanged: (propertyPath: string, oldValue: unknown, newValue: unknown) => void;
  readonly configurationComplete: () => void;
  readonly changesApplied: () => void;
  readonly dispose: () => void;
}
interface NativePropertyPanePreparation { readonly release: () => void; }
type NativePropertyPanePreparationResult =
  | { readonly status: 'prepared'; readonly lease: NativePropertyPanePreparation }
  | { readonly status: 'blocked'; readonly reason: string };
export interface DemoPropertyPanePort<P extends object> {
  readonly prepare: (create: (host: NativePropertyPaneEnvironment<P>) => NativePropertyPaneController) => NativePropertyPanePreparationResult;
}
interface NativeHostEnvironment<P extends object> extends NativePropertyPaneEnvironment<P> { readonly isOpen: () => boolean; }
interface Generation {
  controller?: NativePropertyPaneController;
  configuration?: IPropertyPaneConfiguration;
  readonly fields: Set<FieldCell>;
  withdrawn: boolean;
  committed: boolean;
}
interface FieldCell {
  generation?: Generation;
  render?: IPropertyPaneCustomFieldProps['onRender'];
  dispose?: IPropertyPaneCustomFieldProps['onDispose'];
  dispatch?: FieldDispatch;
  readonly key: string;
}
interface FieldDispatch {
  readonly render: (cell: FieldCell, args: Parameters<IPropertyPaneCustomFieldProps['onRender']>) => void;
  readonly dispose: (cell: FieldCell, element: HTMLElement) => void;
}
interface MountedField { readonly cell: FieldCell; readonly context: unknown; }
interface Candidate<P extends object> {
  create?: (host: NativePropertyPaneEnvironment<P>) => NativePropertyPaneController;
  generation?: Generation;
}
function clearField(cell: FieldCell): void {
  cell.render = undefined;
  cell.dispose = undefined;
  cell.generation = undefined;
  cell.dispatch = undefined;
}
// Issued SDK closures capture this small mutable cell only. Retirement clears its
// business callbacks and host dispatch; no descriptor/controller is captured here.
function fieldWrappers(cell: FieldCell): Pick<IPropertyPaneCustomFieldProps, 'onRender' | 'onDispose'> {
  return {
    onRender: (...args) => cell.dispatch?.render(cell, args),
    onDispose: element => cell.dispatch?.dispose(cell, element)
  };
}

/** Instance-owned bounded adapter for the pinned successful reactive native SDK route.
 * External SDK/React/host failure after getter return requires reload before further
 * native work: public hooks supply no reliable install/failure acknowledgement.
 */
export class NativePropertyPaneDemoHost<P extends object> {
  public readonly port: DemoPropertyPanePort<P>;
  private environment: NativeHostEnvironment<P> | undefined;
  private candidate: Candidate<P> | undefined;
  private produced: Generation | undefined;
  private nativeView: Generation | undefined;
  private readonly mounted = new Map<HTMLElement, MountedField>();
  private readonly witnesses = new Set<Generation>();
  private dispatching: Generation | undefined;
  private disposed = false;
  private readonly dispatch: FieldDispatch = {
    render: (cell, args) => this.renderField(cell, args),
    dispose: (cell, element) => this.disposeField(cell, element)
  };

  public constructor(environment: NativeHostEnvironment<P>, private observe?: (event: 'child-render' | 'field-dispose', fieldKey: string) => void) {
    this.environment = environment;
    this.port = { prepare: create => this.prepare(create) };
  }

  private reconcileClose(): void {
    if (this.environment?.isOpen() || this.mounted.size) return;
    this.witnesses.clear();
    if (this.produced) this.produced.withdrawn = true;
    if (this.candidate?.generation) this.candidate = undefined;
    if (this.nativeView?.committed) {
      const previous = this.nativeView;
      this.nativeView = undefined;
      if (this.produced === previous) this.produced = undefined;
      this.retire(previous);
    }
  }

  private prepare(create: (host: NativePropertyPaneEnvironment<P>) => NativePropertyPaneController): NativePropertyPanePreparationResult {
    if (this.disposed || !this.environment) return { status: 'blocked', reason: 'The native host is disposed.' };
    this.reconcileClose();
    if (this.environment.isOpen()) return { status: 'blocked', reason: 'Close the actual native property pane before preparing.' };
    if (this.mounted.size) return { status: 'blocked', reason: 'Native fields still own resources; wait for actual native disposal.' };
    if (this.candidate) this.candidate.create = undefined;
    const candidate: Candidate<P> = { create };
    this.candidate = candidate;
    return { status: 'prepared', lease: { release: () => {
      if (this.candidate !== candidate || candidate.generation) return;
      candidate.create = undefined;
      this.candidate = undefined;
    } } };
  }

  /** Called only by a real permitted getter; every source configuration reads the live SDK bag. */
  public getConfiguration(fallback: () => IPropertyPaneConfiguration): IPropertyPaneConfiguration {
    if (this.disposed || !this.environment) return fallback();
    if (!this.environment.isOpen()) { this.reconcileClose(); return fallback(); }
    const candidate = this.candidate;
    if (!candidate) return fallback();
    if (candidate.generation?.controller) {
      try { return this.configuration(candidate.generation); }
      catch { return candidate.generation.configuration ?? fallback(); }
    }
    if (!candidate.create) return fallback();
    // A successful cold successor can replace its uncommitted predecessor. A
    // committed warm view can coexist with one successor, but not two. The pinned
    // nonreentrant successful DTO/install route establishes this before user input.
    if (this.produced && this.nativeView && this.produced !== this.nativeView) return fallback();
    const generation: Generation = { fields: new Set(), withdrawn: false, committed: false };
    const environment = this.environment;
    const sourceEnvironment: NativePropertyPaneEnvironment<P> = {
      instance: environment.instance,
      readProperties: environment.readProperties,
      refresh: () => {
        if (!this.disposed && this.dispatching === generation && this.nativeView === generation && environment.isOpen()) environment.refresh();
      }
    };
    let configuration: IPropertyPaneConfiguration;
    try {
      generation.controller = candidate.create(sourceEnvironment);
      configuration = this.configuration(generation);
    } catch {
      this.retire(generation);
      candidate.create = undefined;
      this.candidate = undefined;
      return this.produced?.configuration ?? fallback();
    }
    candidate.create = undefined;
    candidate.generation = generation;
    const previous = this.produced;
    this.produced = generation;
    if (previous && previous !== this.nativeView) this.retire(previous);
    return configuration;
  }

  private configuration(generation: Generation): IPropertyPaneConfiguration {
    const controller = generation.controller;
    if (!controller) throw new Error('Native generation is retired.');
    const configuration = controller.getConfiguration();
    const first = configuration.pages[0]?.groups[0];
    if (configuration.pages.length !== 1 || configuration.showLoadingIndicator ||
      !first || !('groupFields' in first) || first.isGroupNameHidden !== true ||
      Object.prototype.hasOwnProperty.call(first, 'groupName') || Object.prototype.hasOwnProperty.call(first, 'isCollapsed') ||
      first.groupFields.length !== 1 || first.groupFields[0].type !== PropertyPaneFieldType.Custom ||
      configuration.pages[0].groups.some(group => !('groupFields' in group))) {
      throw new Error('Native scenarios require one plain page with one genuine unnamed first custom subject and synchronous configuration.');
    }
    // Clear superseded never-mounted originals. Native DTOs may retain their
    // wrappers; those retain only cleared cells. Mounted originals remain until
    // their actual same-DOM replacement or latest disposal.
    for (const cell of generation.fields) {
      if (![...this.mounted.values()].some(record => record.cell === cell)) {
        clearField(cell);
        generation.fields.delete(cell);
      }
    }
    const wrapped = { ...configuration, pages: configuration.pages.map(page => ({ ...page, groups: page.groups.map(group => {
      if (!('groupFields' in group)) throw new Error('Native scenarios require plain groups on one page.');
      return { ...group, groupFields: group.groupFields.map(field => this.wrapField(generation, field)) };
    }) })) };
    generation.configuration = wrapped;
    return wrapped;
  }

  private wrapField(generation: Generation, field: IPropertyPaneField<IPropertyPaneCustomFieldProps>): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    if (field.type !== PropertyPaneFieldType.Custom) return field;
    const cell: FieldCell = { generation, key: field.properties.key, render: field.properties.onRender, dispose: field.properties.onDispose, dispatch: this.dispatch };
    generation.fields.add(cell);
    return { ...field, properties: { ...field.properties, ...fieldWrappers(cell) } };
  }

  private renderField(cell: FieldCell, args: Parameters<IPropertyPaneCustomFieldProps['onRender']>): void {
    const generation = cell.generation;
    if (this.disposed || !generation?.controller) return;
    this.witnesses.add(generation);
    if (generation.withdrawn || !this.environment?.isOpen()) return;
    const [element, context] = args;
    const previous = this.mounted.get(element);
    if (previous && previous.cell !== cell && (previous.cell.generation !== generation || previous.cell.key !== cell.key)) this.disposeField(previous.cell, element);
    this.observe?.('child-render', cell.key);
    cell.render?.(...args);
    this.mounted.set(element, { cell, context });
    if (previous && previous.cell !== cell) {
      previous.cell.generation?.fields.delete(previous.cell);
      clearField(previous.cell);
    }
  }

  private disposeField(cell: FieldCell, element: HTMLElement): void {
    const mounted = this.mounted.get(element);
    if (!mounted || mounted.cell !== cell) return;
    this.mounted.delete(element);
    try { this.observe?.('field-dispose', cell.key); cell.dispose?.(element, mounted.context); }
    finally {
      // Native accordion collapse can unmount and later remount this very
      // descriptor. Keep its originals while the generation remains live; close,
      // a later descriptor issue, or terminal retirement clears them.
      if (this.disposed || cell.generation?.withdrawn || !this.environment?.isOpen()) {
        cell.generation?.fields.delete(cell);
        clearField(cell);
      }
    }
  }

  public rendered(): void {
    if (this.disposed) return;
    let witnessed: Generation | undefined;
    if (this.witnesses.size === 1) for (const generation of this.witnesses) witnessed = generation;
    this.witnesses.clear();
    if (!witnessed?.controller) return;
    const previous = this.nativeView;
    this.nativeView = witnessed;
    witnessed.committed = true;
    if (previous && previous !== witnessed && previous !== this.produced) this.retire(previous);
  }

  public fieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    const owner = this.nativeView;
    if (this.disposed || !owner?.controller || !this.environment?.isOpen()) return;
    this.dispatching = owner;
    try { owner.controller.fieldChanged(propertyPath, oldValue, newValue); }
    finally { this.dispatching = undefined; }
  }
  public configurationComplete(): void {
    if (this.disposed) return;
    this.nativeView?.controller?.configurationComplete();
    this.reconcileClose();
  }
  public changesApplied(): void { if (!this.disposed) this.nativeView?.controller?.changesApplied(); }

  private retire(generation: Generation): void {
    generation.withdrawn = true;
    const controller = generation.controller;
    generation.controller = undefined;
    generation.configuration = undefined;
    this.witnesses.delete(generation);
    for (const cell of generation.fields) clearField(cell);
    generation.fields.clear();
    controller?.dispose();
  }

  public dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    if (this.candidate) this.candidate.create = undefined;
    this.candidate = undefined;
    // Invalidate all wrapper admission before invoking real native cleanup.
    for (const generation of [this.nativeView, this.produced]) if (generation) generation.withdrawn = true;
    try {
      for (const [element, mounted] of [...this.mounted]) this.disposeField(mounted.cell, element);
    } finally {
      const generations = new Set([this.nativeView, this.produced]);
      this.nativeView = undefined;
      this.produced = undefined;
      this.dispatching = undefined;
      this.environment = undefined;
      this.observe = undefined;
      this.witnesses.clear();
      for (const generation of generations) if (generation) this.retire(generation);
    }
  }
}
