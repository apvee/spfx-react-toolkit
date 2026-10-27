// Private native SPFx module declarations. Never re-export from a public entrypoint.

declare module 'ApveeSpfxReactToolkitPaginationStrings' {
  const resource: import('../controls/pagination/localization/pagination-resource.types.internal').PaginationResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitDataGridStrings' {
  const resource: import('../controls/data-grid/localization/data-grid-resource.types.internal').DataGridResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitEditableTextStrings' {
  const resource: import('../controls/editable-text/localization/editable-text-resource.types.internal').EditableTextResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitWebPartShellStrings' {
  const resource: import('../controls/web-part-shell/localization/web-part-shell-resource.types.internal').WebPartShellResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitPickersStrings' {
  const resource: import('../controls/pickers/localization/pickers-resource.types.internal').PickersResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitFileFolderPickerStrings' {
  const resource: import('../controls/file-folder-picker/localization/file-folder-picker-resource.types.internal').FileFolderPickerResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitFormStrings' {
  const resource: import('../controls/form/localization/form-resource.types.internal').FormResource & { readonly __locale: string };
  export = resource;
}

declare module 'ApveeSpfxReactToolkitStageStrings' {
  const resource: import('../controls/stage/localization/stage-resource.types.internal').StageResource & { readonly __locale: string };
  export = resource;
}
