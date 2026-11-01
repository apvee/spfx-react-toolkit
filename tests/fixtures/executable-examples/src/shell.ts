/** Public host properties. */
export interface IWebPartShellProperties { title: string; }
/** Public shell presentation. */
export interface WebPartShellProps<T extends IWebPartShellProperties> {
  properties: T;
  appearance?: 'plain' | 'card';
  size?: 'small' | 'large';
  titlePosition?: 'top' | 'bottom';
  padding?: 'none' | 'medium';
}
