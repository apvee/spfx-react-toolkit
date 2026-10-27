/** Translator data for the tooling regression fixture. */
export interface MiniResource {
  readonly title: string;
  readonly greetingTemplate: string;
  readonly calendar: {
    readonly months: readonly string[];
    readonly days: readonly string[];
  };
  readonly count: {
    readonly zero?: string;
    readonly one?: string;
    readonly two?: string;
    readonly few?: string;
    readonly many?: string;
    readonly other: string;
  };
}
