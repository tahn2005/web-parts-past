import * as React from 'react';
import { getIcon } from '../icons';
import type { IconSource } from '../models';
import { safeImageSrc } from '../urls';

export interface ITileIconProps {
  className: string;
  imageClassName: string;
  source: IconSource;
  iconName: string;
  imageUrl: string;
}

/**
 * The small coloured tile holding an icon or an uploaded image. Always
 * decorative: the card or section title next to it carries the meaning, so it
 * is hidden from assistive technology rather than announced twice.
 */
export function TileIcon({ className, imageClassName, source, iconName, imageUrl }: ITileIconProps): JSX.Element {
  const src = source === 'image' ? safeImageSrc(imageUrl) : undefined;
  const Icon = getIcon(iconName);

  return (
    <span className={className} aria-hidden="true">
      {src ? (
        <img className={imageClassName} src={src} alt="" loading="lazy" />
      ) : (
        <Icon aria-hidden="true" focusable="false" />
      )}
    </span>
  );
}
