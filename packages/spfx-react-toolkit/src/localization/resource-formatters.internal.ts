/** Values are already chosen/formatted by the owning family; never invoke caller callbacks. */
export type ResourceTemplateValues = Readonly<Record<string, string | number>>;

/** Data-only cardinal forms. The generator validates the categories for every native culture. */
export interface ResourcePluralForms {
  readonly zero?: string;
  readonly one?: string;
  readonly two?: string;
  readonly few?: string;
  readonly many?: string;
  readonly other: string;
}

/** One replacement pass preserves caller dollar sequences, braces and token-like content literally. */
export function formatResourceTemplate(template: string, values: ResourceTemplateValues): string {
  return template.replace(/\{([^{}]+)\}/g, (_match: string, token: string): string => {
    if (!Object.prototype.hasOwnProperty.call(values, token)) {
      throw new Error(`Missing resource template value: ${token}`);
    }
    return String(values[token]);
  });
}

/** Native resource metadata selects grammar independently of caller number/data formatting. */
export function formatResourcePlural(
  resource: ResourcePluralForms,
  locale: string,
  count: number,
  values: ResourceTemplateValues = {},
): string {
  const category = new Intl.PluralRules(locale).select(count);
  const template = resource[category] ?? resource.other;
  return formatResourceTemplate(template, { count, ...values });
}
