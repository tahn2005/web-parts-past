import * as React from 'react';
import { ArrowRight, ChevronLeft, ChevronRight, Pencil, Trash2 } from 'lucide-react';
import * as strings from 'ResourceAccordionWebPartStrings';
import styles from './AccordionCards.module.scss';
import { TileIcon } from './TileIcon';
import type { IResourceCard } from '../models';
import type { CssVars } from '../theme';
import { safeHref } from '../urls';

export interface IResourceCardProps {
  card: IResourceCard;
  index: number;
  total: number;
  vars: CssVars;
  /** Card titles sit one level below the section heading. */
  headingTag: 'h3' | 'h4' | 'h5';
  defaultLinkText: string;
  isEditMode: boolean;
  isSelected: boolean;
  onSelect: (id: string) => void;
  onMove: (id: string, delta: number) => void;
  onDelete: (id: string) => void;
}

/**
 * A single resource card — the web part equivalent of ResourceCard.tsx in the
 * app. In edit mode it also carries a small toolbar and selects itself when
 * clicked, so each card behaves like its own web part on the canvas.
 */
export function ResourceCard(props: IResourceCardProps): JSX.Element {
  const { card, index, total, vars, headingTag: CardHeading, isEditMode, isSelected } = props;
  const href = safeHref(card.linkUrl);
  const linkText = card.linkText.trim() || props.defaultLinkText.trim() || 'View Resource';
  const displayTitle = card.title.trim();

  /** Toolbar clicks must not also bubble up and select the card. */
  const stop = (handler: () => void): React.MouseEventHandler => (event) => {
    event.stopPropagation();
    handler();
  };

  const className = [styles.card, isEditMode && isSelected ? styles.selected : ''].filter(Boolean).join(' ');

  return (
    <div
      className={className}
      style={vars as React.CSSProperties}
      onClick={isEditMode ? () => props.onSelect(card.id) : undefined}
    >
      {isEditMode && (
        <div
          className={styles.cardTools}
          role="group"
          aria-label={strings.CardOptions.replace('{0}', displayTitle || strings.UntitledCard)}
        >
          <button
            type="button"
            className={styles.toolButton}
            title={strings.EditCard}
            aria-label={strings.EditCard}
            aria-pressed={isSelected}
            data-focus-key={`edit-${card.id}`}
            onClick={stop(() => props.onSelect(card.id))}
          >
            <Pencil aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            title={strings.MoveCardEarlier}
            aria-label={strings.MoveCardEarlier}
            disabled={index === 0}
            data-focus-key={`earlier-${card.id}`}
            onClick={stop(() => props.onMove(card.id, -1))}
          >
            <ChevronLeft aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            title={strings.MoveCardLater}
            aria-label={strings.MoveCardLater}
            disabled={index === total - 1}
            data-focus-key={`later-${card.id}`}
            onClick={stop(() => props.onMove(card.id, 1))}
          >
            <ChevronRight aria-hidden="true" focusable="false" />
          </button>
          <button
            type="button"
            className={styles.toolButton}
            title={strings.DeleteCard}
            aria-label={strings.DeleteCard}
            onClick={stop(() => props.onDelete(card.id))}
          >
            <Trash2 aria-hidden="true" focusable="false" />
          </button>
        </div>
      )}

      <div className={styles.cardIconRow}>
        <TileIcon
          className={styles.cardIcon}
          imageClassName={styles.iconImage}
          source={card.iconSource}
          iconName={card.iconName}
          imageUrl={card.imageUrl}
        />
      </div>

      {(displayTitle || isEditMode) && (
        <CardHeading className={[styles.cardTitle, displayTitle ? '' : styles.placeholder].join(' ')}>
          {displayTitle || strings.UntitledCard}
        </CardHeading>
      )}

      {card.description.trim() && <p className={styles.cardDescription}>{card.description}</p>}

      {href ? (
        <a
          className={styles.cardLink}
          href={href}
          target={card.openInNewTab ? '_blank' : undefined}
          rel={card.openInNewTab ? 'noopener noreferrer' : undefined}
          data-interception={card.openInNewTab ? 'off' : undefined}
          onClick={isEditMode ? (event) => event.preventDefault() : undefined}
        >
          {linkText}
          {/* Visible text first, then context, so the accessible name still starts with what is on screen. */}
          {displayTitle && <span className={styles.srOnly}>: {displayTitle}</span>}
          {card.openInNewTab && <span className={styles.srOnly}> {strings.OpensInNewTab}</span>}
          <ArrowRight aria-hidden="true" focusable="false" />
        </a>
      ) : (
        isEditMode && <span className={styles.noLink}>{strings.NoLinkYet}</span>
      )}
    </div>
  );
}
