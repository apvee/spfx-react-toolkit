import type { ReactNode } from 'react';

/**
 * Entra group classification applied to every selector. All (default) adds no
 * metadata restriction; M365 requires Unified; Security requires securityEnabled
 * and excludes Unified, including mail-enabled and dynamic security groups.
 * @example const type: EntraGroupType = 'Security';
 */
export type EntraGroupType = 'All' | 'M365' | 'Security';
/**
 * member (default) uses Graph transitive membership; M365 membership is direct.
 * owner uses direct owners; memberOrOwner accepts either for each selected group.
 * Ownership checks fail for synchronized groups and non-M365 mail-enabled or
 * distribution groups, even when memberOrOwner could otherwise match membership.
 * @example const relationship: EntraGroupRelationship = 'memberOrOwner';
 */
export type EntraGroupRelationship = 'member' | 'owner' | 'memberOrOwner';
/**
 * Exactly one Entra group GUID or exact nonblank display name. groupType belongs
 * only on EntraGroupTrim, never on individual selectors. Missing/ambiguous names
 * and mismatched typed IDs are errors; name type filtering precedes ambiguity.
 * @example
 * ```tsx
 * const group: EntraGroupSelector = { name: 'Engineering' };
 * <EntraGroupTrim groups={[group]} groupType="M365">Team</EntraGroupTrim>;
 * ```
 */
export type EntraGroupSelector =
  | {
      /** Group GUID; compared case-insensitively. */ readonly id: string;
      /** Excluded when selecting by ID. */ readonly name?: never;
      /** Classification is control-level only. */ readonly groupType?: never;
    }
  | {
      /** Exact nonblank display name. */ readonly name: string;
      /** Excluded when selecting by name. */ readonly id?: never;
      /** Classification is control-level only. */ readonly groupType?: never;
    };
/**
 * Advisory current-user Entra membership/ownership check using the matching SPFx
 * provider's native Graph client. Graph permissions/consent and hidden membership
 * restrictions affect observable results; absent hidden membership cannot prove
 * real-world absence. Unsupported synchronized, Exchange-created mail-enabled
 * and distribution group owners fail as errors
 * when metadata identifies that restriction. All ID member checks avoid metadata.
 * @example
 * ```tsx
 * const props: EntraGroupTrimProps = {
 *   groups: [{ name: 'Engineering' }], groupType: 'M365', relationship: 'memberOrOwner',
 * };
 * <EntraGroupTrim {...props}>Team content</EntraGroupTrim>;
 * ```
 */
export interface EntraGroupTrimProps {
  /** Required nonempty GUID/name selectors. All selector syntax is validated before requests; names and metadata-required IDs resolve before decision. All-ID member checks cannot distinguish missing groups from nonmembership. */
  readonly groups: readonly EntraGroupSelector[];
  /** All (default), M365 or Security, case-sensitive, applied to every selector. */
  readonly groupType?: EntraGroupType;
  /** member (default), owner or memberOrOwner; errors in either requested relationship fail the whole check. */
  readonly relationship?: EntraGroupRelationship;
  /** all (default) requires every selected group; any requires one, after all reads succeed. */
  readonly match?: 'all' | 'any';
  /** Protected content and unique direct family branches. Inactive content never mounts; omission renders empty. ReactNode preserves legacy null input. */
  readonly children?: ReactNode;
}
/**
 * Structural Entra-group branch content; no wrapper, style, ref or DOM props.
 * Changing the active branch or its React marker key remounts its content;
 * unchanged branch/key rerenders preserve consumer state and child keys.
 * @example
 * ```tsx
 * const branch: EntraGroupTrimBranchProps = { children: 'Checking directory' };
 * <EntraGroupTrim.Loading {...branch} />;
 * ```
 */
export interface EntraGroupTrimBranchProps {
  /** Optional active branch children; omission renders empty. ReactNode preserves legacy null input. */
  readonly children?: ReactNode;
}
