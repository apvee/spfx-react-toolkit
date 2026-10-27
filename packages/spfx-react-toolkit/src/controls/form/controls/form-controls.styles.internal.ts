import { makeStyles } from '@griffel/react';

/** Logical intrinsic minima, truncation and native-input hiding have no composed descriptor. */
export const useFormControlsStyles = makeStyles({
  colorPicker: {
    '& .fui-ColorArea': { width: '100%', minWidth: 0, minHeight: 0, aspectRatio: '1 / 1' },
    '& .fui-ColorSlider, & .fui-AlphaSlider': { minWidth: 0, maxWidth: '100%' },
  },
  fieldset: { minInlineSize: '0px' },
  fileControls: { minInlineSize: '0px' },
  fileSelection: { maxInlineSize: '100%', minInlineSize: '0px' },
  fileSelectionText: { minInlineSize: '0px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  // Native file inputs must remain focusable for RHF error focus and label activation.
  fileHidden: { border: '0px', clip: 'rect(0 0 0 0)', clipPath: 'inset(50%)', height: '1px', margin: '-1px', overflow: 'hidden', padding: '0px', position: 'absolute', whiteSpace: 'nowrap', width: '1px' } });
