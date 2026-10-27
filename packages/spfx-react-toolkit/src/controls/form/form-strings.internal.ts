import type { FormStrings, FormStringsOverride } from './form-strings.types';
import { nativeFormResource } from './localization/form-native.internal';
import { formatResourceTemplate } from '../../localization/resource-formatters.internal';

const resource = nativeFormResource;
const array = resource.arraySection;
const nativeStrings: FormStrings = {
    unexpectedFailure: resource.unexpectedFailure,
    datePicker: resource.datePicker,
    dateInputValidationMessage: resource.dateInputValidationMessage,
    timeInputValidationMessage: resource.timeInputValidationMessage,
    tagPicker: resource.tagPicker,
    fileInput: resource.fileInput,
    arraySection: {
        add: array.add, empty: array.empty,
        itemLabel: position => formatResourceTemplate(array.itemLabelTemplate, { position }),
        remove: label => formatResourceTemplate(array.removeTemplate, { label }),
        moveUp: label => formatResourceTemplate(array.moveUpTemplate, { label }),
        moveDown: label => formatResourceTemplate(array.moveDownTemplate, { label }),
        added: label => formatResourceTemplate(array.addedTemplate, { label }),
        removed: label => formatResourceTemplate(array.removedTemplate, { label }),
        moved: (label, position, total) => formatResourceTemplate(array.movedTemplate, { label, position, total }),
    },
};
/** Merge only known fields; undefined retains native defaults and empty text remains meaningful. */
export function mergeFormStringFields<T extends object>(defaults: T, overrides?: Partial<T>): T {
    const result = { ...defaults };
    for (const key of Object.keys(defaults) as Array<keyof T>) {
        const value = overrides?.[key];
        if (value !== undefined) result[key] = value;
    }
    return result;
}
/** Resolve instance UI text without changing parsing, values or validation state. */
export function resolveFormStrings(overrides?: FormStringsOverride): FormStrings {
    return {
        unexpectedFailure: overrides?.unexpectedFailure ?? nativeStrings.unexpectedFailure,
        dateInputValidationMessage: overrides?.dateInputValidationMessage ?? nativeStrings.dateInputValidationMessage,
        timeInputValidationMessage: overrides?.timeInputValidationMessage ?? nativeStrings.timeInputValidationMessage,
        datePicker: mergeFormStringFields(nativeStrings.datePicker, overrides?.datePicker),
        tagPicker: mergeFormStringFields(nativeStrings.tagPicker, overrides?.tagPicker),
        fileInput: mergeFormStringFields(nativeStrings.fileInput, overrides?.fileInput),
        arraySection: mergeFormStringFields(nativeStrings.arraySection, overrides?.arraySection),
    };
}
