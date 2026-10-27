import * as React from 'react';
import { PreviewCard } from '__TOOLKIT__';

export function Example(): React.ReactElement {
  return <PreviewCard title="**Literal** report" orientation="horizontal" image={{ src: "/cover.jpg", alt: "Cover" }} logo={{ as: "img", src: "/logo.svg", alt: "Brand" }} />;
}
