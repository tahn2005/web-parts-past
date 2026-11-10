import * as React from 'react';
import { Plus } from 'lucide-react';
import styles from './ResourceCards.module.scss';
import ResourceCard, { type CardControl } from './ResourceCard';
import type { HeadingLevel, IResourceCard } from '../common/model';
import { getPalette } from '../common/color';
import { CARD_FONT_FAMILY } from '../common/fonts';

export interface IResourceCardsProps {
  cards: IResourceCard[];
  accentColor: string;
  headingLevel: HeadingLevel;
  useInterFont: boolean;
  isEditMode: boolean;
  /** Card currently open in the property pane, highlighted on the canvas. */
  selectedCardId: string | undefined;
  /** Returns the new card's id. */
  onAddCard: () => string;
  onEditCard: (id: string) => void;
  onMoveCard: (id: string, delta: -1 | 1) => void;
  onDeleteCard: (id: string) => void;
}

/**
 * The card grid. One column on narrow widths, two from ~590px, three from
 * ~970px — the same points at which the original React app's grid changes, measured
 * on the web part rather than the browser window so it behaves correctly in
 * any SharePoint section layout.
 *
 * In edit mode it adds Quick Links-style controls: an "Add card" button and a
 * toolbar on each card to edit, reorder or delete it.
 */
// React renders nothing for `null`; there is no `undefined` equivalent.
// eslint-disable-next-line @rushstack/no-new-null
export default function ResourceCards(props: IResourceCardsProps): React.ReactElement | null {
  const { cards, accentColor, headingLevel, useInterFont, isEditMode, selectedCardId } = props;

  const [announcement, setAnnouncement] = React.useState('');
  // Toolbar buttons, keyed "<card id>:<control>".
  const toolButtons = React.useRef<{ [key: string]: HTMLButtonElement | undefined }>({});
  const addButton = React.useRef<HTMLButtonElement | undefined>(undefined);
  /** Where focus should go after the next render: a toolbar button key or 'add'. */
  const pendingFocus = React.useRef<string | undefined>(undefined);

  // Restore focus after a card is moved or deleted. Done after render because
  // the button to focus may only exist (or only be in its new place) then.
  React.useEffect(() => {
    const target = pendingFocus.current;
    if (target === undefined) {
      return;
    }
    pendingFocus.current = undefined;
    let element = target === 'add' ? addButton.current : toolButtons.current[target];
    // A card moved to either end disables that move button; keep focus on the
    // card by switching to the other one.
    if (element && element.disabled) {
      const split = target.lastIndexOf(':');
      const other = target.slice(split + 1) === 'earlier' ? 'later' : 'earlier';
      element = toolButtons.current[`${target.slice(0, split)}:${other}`];
    }
    if (element && !element.disabled) {
      element.focus();
    }
  });

  const announce = (message: string): void => {
    // Clear first so the same message twice in a row is still read out.
    setAnnouncement('');
    window.setTimeout(() => setAnnouncement(message), 50);
  };

  const describe = (card: IResourceCard): string => `"${card.title.trim() || 'Untitled card'}"`;

  const handleAdd = (): void => {
    const id = props.onAddCard();
    pendingFocus.current = `${id}:edit`;
    announce(`Card added. ${cards.length + 1} cards. The card's settings are open in the property pane.`);
  };

  const handleMove = (card: IResourceCard, index: number, delta: -1 | 1): void => {
    props.onMoveCard(card.id, delta);
    const control: CardControl = delta < 0 ? 'earlier' : 'later';
    pendingFocus.current = `${card.id}:${control}`;
    announce(`${describe(card)} moved to position ${index + delta + 1} of ${cards.length}.`);
  };

  const handleDelete = (card: IResourceCard, index: number): void => {
    const next = cards[index + 1] || cards[index - 1];
    props.onDeleteCard(card.id);
    pendingFocus.current = next ? `${next.id}:edit` : 'add';
    const remaining = cards.length - 1;
    announce(`${describe(card)} deleted. ${remaining} ${remaining === 1 ? 'card' : 'cards'} left.`);
  };

  if (cards.length === 0 && !isEditMode) {
    // Nothing to show readers; SharePoint collapses the empty web part.
    return null;
  }

  const rootStyle = {
    '--rc-font': useInterFont ? CARD_FONT_FAMILY : 'inherit',
    '--rc-ui-accent': getPalette(accentColor).text
  } as React.CSSProperties;

  return (
    <div className={styles.resourceCards} style={rootStyle}>
      {isEditMode && (
        <div className={styles.editBar}>
          <button
            type="button"
            className={styles.addButton}
            onClick={handleAdd}
            ref={(element) => {
              addButton.current = element || undefined;
            }}
          >
            <Plus aria-hidden="true" focusable="false" />
            Add card
          </button>
        </div>
      )}

      {cards.length > 0 ? (
        // role="list" keeps list semantics in Safari/VoiceOver, which drops
        // them from lists styled with list-style: none.
        <ul className={styles.grid} role="list">
          {cards.map((card, index) => (
            <ResourceCard
              key={card.id}
              card={card}
              accentColor={accentColor}
              headingLevel={headingLevel}
              isEditMode={isEditMode}
              isSelected={isEditMode && card.id === selectedCardId}
              isFirst={index === 0}
              isLast={index === cards.length - 1}
              onEdit={() => props.onEditCard(card.id)}
              onMove={(delta) => handleMove(card, index, delta)}
              onDelete={() => handleDelete(card, index)}
              buttonRef={(control, element) => {
                toolButtons.current[`${card.id}:${control}`] = element;
              }}
            />
          ))}
        </ul>
      ) : (
        <div className={styles.empty}>
          <p className={styles.emptyTitle}>No cards yet</p>
          <p className={styles.emptyText}>Select <strong>Add card</strong> to create your first resource card.</p>
        </div>
      )}

      <div className={styles.srOnly} role="status" aria-live="polite" aria-atomic="true">
        {announcement}
      </div>
    </div>
  );
}
