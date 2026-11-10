import * as React from 'react';
import { ArrowRight, ChevronLeft, ChevronRight, Pencil, Trash2 } from 'lucide-react';
import styles from './ResourceCards.module.scss';
import { IMAGE_SIZES, type HeadingLevel, type IResourceCard } from '../common/model';
import { getPalette } from '../common/color';
import { getIcon } from '../common/icons';
import { safeImageUrl, safeUrl } from '../common/url';

/** Toolbar buttons the list may need to move focus to. */
export type CardControl = 'edit' | 'earlier' | 'later';

export interface IResourceCardProps {
  card: IResourceCard;
  /** Web part accent, used unless the card sets its own. */
  accentColor: string;
  headingLevel: HeadingLevel;
  isEditMode: boolean;
  isSelected: boolean;
  isFirst: boolean;
  isLast: boolean;
  onEdit: () => void;
  onMove: (delta: -1 | 1) => void;
  onDelete: () => void;
  /** Receives the toolbar buttons so the list can move focus to them. */
  buttonRef: (control: CardControl, element: HTMLButtonElement | undefined) => void;
}

/**
 * The icon tile, or the author's image in its place. If the image can't be
 * loaded (deleted, moved, no permission) readers see the icon instead of a
 * broken image, and authors also get a note saying so.
 */
function CardMedia(props: { card: IResourceCard; isEditMode: boolean }): React.ReactElement {
  const { card, isEditMode } = props;
  const imageUrl = card.mediaType === 'image' ? safeImageUrl(card.imageUrl) : '';
  const [failedUrl, setFailedUrl] = React.useState('');
  const { Icon } = getIcon(card.icon);

  const iconTile = (
    <div className={styles.iconTile}>
      <Icon className={styles.icon} aria-hidden="true" focusable="false" />
    </div>
  );

  if (!imageUrl || failedUrl === imageUrl) {
    const note =
      isEditMode && card.mediaType === 'image'
        ? imageUrl
          ? "The image couldn't be loaded. Check the link, or upload it again."
          : 'No image yet. Edit this card to upload one.'
        : '';
    return (
      <div className={styles.iconRow}>
        {iconTile}
        {note && <span className={styles.mediaNote}>{note}</span>}
      </div>
    );
  }

  const size = IMAGE_SIZES[card.imageSize] || IMAGE_SIZES.small;
  const alt = card.imageAlt.trim();
  return (
    <div className={styles.iconRow}>
      <img
        className={styles.image}
        src={imageUrl}
        // Empty alt marks the image decorative: the title right below already
        // names the card, so screen readers skip it unless the author wrote alt text.
        alt={alt}
        style={{ width: size, height: size }}
        loading="lazy"
        decoding="async"
        onError={() => setFailedUrl(imageUrl)}
      />
    </div>
  );
}

/**
 * One resource card. Mirrors ResourceCard.tsx in the original React app:
 * icon tile, title, description and an accent-coloured link with an arrow.
 *
 * Differences, all for accessibility, none visible at rest:
 * - It is a list item with a real heading, and the link is a real link.
 * - The link's accessible name includes the card title, so a screen reader
 *   user tabbing through hears "View Resource: Interviewing", not
 *   "View Resource" ten times.
 * - Keyboard focus gets a visible ring and the same highlight as hover.
 */
export default function ResourceCard(props: IResourceCardProps): React.ReactElement {
  const { card, accentColor, headingLevel, isEditMode, isSelected, isFirst, isLast } = props;

  const palette = getPalette(card.accentColor || accentColor);
  const url = safeUrl(card.url);
  const linkText = card.linkText.trim();
  const title = card.title.trim();
  const description = card.description.trim();
  const titleForLabels = title || 'Untitled card';
  const Heading = `h${headingLevel}` as 'h2';

  const cssVars = {
    '--rc-accent': palette.accent,
    '--rc-light': palette.light,
    '--rc-border': palette.border,
    '--rc-text': palette.text,
    '--rc-muted': palette.muted
  } as React.CSSProperties;

  const linkContent = (
    <>
      {linkText}
      <ArrowRight className={styles.arrow} aria-hidden="true" focusable="false" />
    </>
  );
  // Starts with the visible text (WCAG 2.5.3) and adds the card title, so a
  // screen reader user tabbing through hears "View Resource: Interviewing"
  // rather than "View Resource" on every card.
  const linkLabel = `${linkText}${title ? `: ${title}` : ''}${card.openInNewTab ? ' (opens in a new tab)' : ''}`;

  let link: React.ReactNode = null;
  if (linkText && url && !isEditMode) {
    link = (
      <a
        className={styles.link}
        href={url}
        target={card.openInNewTab ? '_blank' : undefined}
        rel={card.openInNewTab ? 'noopener noreferrer' : undefined}
        data-interception={card.openInNewTab ? 'off' : undefined}
        aria-label={linkLabel}
      >
        {linkContent}
      </a>
    );
  } else if (linkText && url) {
    // Edit mode: show the link exactly as it will look, but don't navigate
    // away from an unsaved page. Editing happens through the toolbar.
    link = (
      <span className={styles.link} aria-hidden="true">
        {linkContent}
      </span>
    );
  } else if (isEditMode) {
    link = (
      <span className={styles.missingLink}>
        {linkText ? 'No link yet. Edit this card to add a URL.' : 'No link text. Edit this card to add a link.'}
      </span>
    );
  }

  return (
    <li className={isSelected ? `${styles.card} ${styles.selected}` : styles.card} style={cssVars}>
      <CardMedia card={card} isEditMode={isEditMode} />

      {(title || isEditMode) && (
        <Heading className={styles.title}>{title || <span className={styles.placeholder}>Untitled card</span>}</Heading>
      )}
      <p className={styles.description}>{description}</p>
      {link}

      {isEditMode && (
        <div className={styles.toolbar} role="group" aria-label={`Card actions: ${titleForLabels}`}>
          <button
            type="button"
            className={styles.toolButton}
            onClick={props.onEdit}
            ref={(element) => props.buttonRef('edit', element || undefined)}
            aria-label={`Edit card: ${titleForLabels}`}
            title="Edit card"
          >
            <Pencil aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            onClick={() => props.onMove(-1)}
            ref={(element) => props.buttonRef('earlier', element || undefined)}
            disabled={isFirst}
            aria-label={`Move earlier: ${titleForLabels}`}
            title="Move earlier"
          >
            <ChevronLeft className={styles.flipInRtl} aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            onClick={() => props.onMove(1)}
            ref={(element) => props.buttonRef('later', element || undefined)}
            disabled={isLast}
            aria-label={`Move later: ${titleForLabels}`}
            title="Move later"
          >
            <ChevronRight className={styles.flipInRtl} aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            onClick={props.onDelete}
            aria-label={`Delete card: ${titleForLabels}`}
            title="Delete card"
          >
            <Trash2 aria-hidden="true" focusable="false" />
          </button>
        </div>
      )}
    </li>
  );
}
