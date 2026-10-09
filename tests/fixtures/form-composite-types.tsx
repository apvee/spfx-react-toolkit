import * as React from 'react';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
import type { FormHsvColor } from '../../packages/spfx-react-toolkit/src/controls/form/form-value.types';
import { AlphaSlider, ColorArea, ColorSwatch } from '@fluentui/react-components';
const F = createForm<{ score: number; optional?: number; narrow: 1 | 2; nullable: number | null; color: FormHsvColor; alphaRequired: Required<FormHsvColor>; invalidColor: string; swatch: string; narrowSwatch: 'red' | 'blue' }>();
const root = React.createRef<HTMLDivElement>();
<F.Rating name="score" ref={root} onChange={(event, data) => data.value.toFixed()} />;
<F.Rating name="optional" />;
<F.ColorPicker name="color" ref={root} onColorChange={(event, data) => { data.color.h.toFixed(); data.color.a?.toFixed(); }}><ColorArea /><AlphaSlider /></F.ColorPicker>;
<F.SwatchPicker name="swatch" ref={root} onSelectionChange={(event, data) => data.selectedValue.toUpperCase()}><ColorSwatch value="blue" color="#0000ff" /></F.SwatchPicker>;
// @ts-expect-error complete number writer cannot target literal field
<F.Rating name="narrow" />;
// @ts-expect-error nullable read is incompatible with Rating
<F.Rating name="nullable" />;
// @ts-expect-error color is HSV, no hidden string conversion
<F.ColorPicker name="invalidColor" />;
// @ts-expect-error optional alpha writes cannot target required alpha
<F.ColorPicker name="alphaRequired" />;
// @ts-expect-error complete string writer cannot target literal field
<F.SwatchPicker name="narrowSwatch" />;
// @ts-expect-error public root ref is a div
<F.Rating name="score" ref={React.createRef<HTMLTextAreaElement>()} />;
// @ts-expect-error no unsupported codec
<F.ColorPicker name="color" codec={{}} />;
