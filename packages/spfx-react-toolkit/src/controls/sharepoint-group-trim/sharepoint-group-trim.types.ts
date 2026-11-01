/**
 * @packageDocumentation
 * @docPage controls/sharepoint-group-trim
 */
import type { ReactNode } from 'react';

/**
 * Exactly one SharePoint site-group reference. IDs are positive Int32 numbers;
 * names are exact, nonblank titles. Ambiguous/missing references are errors.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupSelector, SharePointGroupTrim } from '@apvee/spfx-react-toolkit';
 * const group: SharePointGroupSelector = { name: 'Project Readers' };
 * <SharePointGroupTrim groups={[group]}>Project</SharePointGroupTrim>;
 * ```
 */
export type SharePointGroupSelector =
  | {
      /**
       * Positive Int32 site-group ID.
       */ readonly id: number;
      /**
       * Excluded when selecting by ID.
       * Must be omitted or undefined in this branch.
       */ readonly name?: never;
    }
  | {
      /**
       * Exact nonblank group title.
       */ readonly name: string;
      /**
       * Excluded when selecting by name.
       * Must be omitted or undefined in this branch.
       */ readonly id?: never;
    };
/**
 * Current-user direct SharePoint group membership beneath a matching SPFx
 * provider. Remote-web current user is resolved in that web's site collection.
 * Entra principals contained in a SharePoint group are not expanded.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim, SharePointGroupTrimProps } from '@apvee/spfx-react-toolkit';
 * const props: SharePointGroupTrimProps = { groups: [{ id: 5 }], match: 'any' };
 * <SharePointGroupTrim {...props}>Team content</SharePointGroupTrim>;
 * ```
 */
export interface SharePointGroupTrimProps {
  /**
   * Required nonempty ID/name selectors; all must resolve successfully before any decision.
   */
  readonly groups: readonly SharePointGroupSelector[];
  /**
   * Absolute HTTP(S) web URL without query/fragment; omitted uses current web.
   */
  readonly webUrl?: string;
  /**
   * all (when omitted) requires direct membership in every selected group; any requires one.
   * @defaultValue 'all'
   */
  readonly match?: 'all' | 'any';
  /**
   * Protected content and unique direct family branches. Inactive content never mounts; when omitted, renders empty. ReactNode preserves legacy null input.
   */
  readonly children?: ReactNode;
}
/**
 * Structural SharePoint-group branch content; no DOM/style/ref props or wrapper.
 * Changing the active branch or its React marker key remounts its content;
 * unchanged branch/key rerenders preserve consumer state and child keys.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SharePointGroupTrim, SharePointGroupTrimBranchProps } from '@apvee/spfx-react-toolkit';
 * const branch: SharePointGroupTrimBranchProps = { children: 'Checking groups' };
 * <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
 *   <SharePointGroupTrim.Loading {...branch} />;
 * </SharePointGroupTrim>
 * ```
 */
export interface SharePointGroupTrimBranchProps {
  /**
   * Optional active branch children. When omitted, renders empty; ReactNode preserves legacy null input.
   */
  readonly children?: ReactNode;
}
