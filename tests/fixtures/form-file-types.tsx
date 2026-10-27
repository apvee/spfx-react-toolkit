import * as React from 'react';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
const F = createForm<{ file: FileList | null; optional?: FileList | null; nonnullable: FileList; tooWide: FileList | null | string; onlyEmpty: null }>();
const ref = React.createRef<HTMLInputElement>();
<F.FileInput name="file" ref={ref} multiple accept="application/pdf" browseLabel="Choose" clearLabel="Remove" formatFiles={files => files.item(0)?.name} onChange={event => { const target: HTMLInputElement = event.currentTarget; void target; }} onBlur={event => { const target: HTMLInputElement = event.currentTarget; void target; }} />;
<F.FileInput name="optional" />;
// @ts-expect-error native clear writes null
<F.FileInput name="nonnullable" />;
// @ts-expect-error native input cannot read unrelated string values
<F.FileInput name="tooWide" />;
// @ts-expect-error native selection writes FileList
<F.FileInput name="onlyEmpty" />;
// @ts-expect-error public ref targets the actual file input
<F.FileInput name="file" ref={React.createRef<HTMLDivElement>()} />;
// @ts-expect-error RHF owns the live value
<F.FileInput name="file" value="file" />;
// @ts-expect-error RHF owns defaults
<F.FileInput name="file" defaultValue="file" />;
// @ts-expect-error native type is fixed
<F.FileInput name="file" type="text" />;
