/** Public host properties. */
export interface IWebPartShellProperties { title: string; }
/** Public shell presentation. */
export interface WebPartShellProps<T extends IWebPartShellProperties> {
  properties: T;
  appearance?: 'plain' | 'card';
  size?: 'small' | 'large';
  titlePosition?: 'top' | 'bottom';
  padding?: 'none' | 'medium';
  surface?: 'auto' | 'section' | 'application';
  dir?: 'ltr' | 'rtl';
  strings?: { titleEditorLabel?: string; titlePlaceholder?: string };
  titleEditorStrings?: { bold?: string; colorNames?: { red?: string }; formattingShortcut?: (label: string, shortcut: string) => string };
  action?: import('react').ReactNode;
}
