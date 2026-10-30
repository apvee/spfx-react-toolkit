import * as React from 'react';
import { Image } from '@fluentui/react-components';
import type { PreviewCardImage } from '../preview-card/preview-card.types';
import type { ContentSlot } from '../content-slot.types.internal';
import { renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import { useCardStyles } from './card.styles.internal';

/** Parent keys by src: other presentation changes cannot automatically retry a failed resource. */
export function CardImage({ image, fallback }: { image: PreviewCardImage; fallback?: ContentSlot<'div'> }): React.ReactElement {
  const [failed, setFailed] = React.useState(false);
  const sx = useSx(); const styles = useCardStyles();
  if (!failed) return <Image src={image.src} alt={image.alt} aria-hidden={image.alt.length === 0 ? true : undefined} fit="cover" className={sx(styles.media, styles.image)} onError={() => setFailed(true)} />;
  const meaningful = image.alt.trim().length > 0;
  return <div data-card-image-fallback="" className={sx(styles.media)} role={meaningful ? 'img' : undefined} aria-label={meaningful ? image.alt : undefined} aria-hidden={meaningful ? undefined : true}>
    {fallback === undefined ? meaningful ? <div className={sx(styles.fallback)}>{image.alt}</div> : undefined :
      renderContentSlot(fallback, 'div', sx, { shorthand: 'wrap', className: styles.fallback })}
  </div>;
}
