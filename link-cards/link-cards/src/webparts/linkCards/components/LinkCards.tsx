import * as React from 'react';
import { ChevronLeft, ChevronRight, Pencil, Plus, Trash2 } from 'lucide-react';
import * as strings from 'LinkCardsWebPartStrings';
import styles from './LinkCards.module.scss';
import { getIcon } from '../icons';
import { safeHref } from '../urls';
import type { ColumnSetting, ILinkCard } from '../models';
import type { CssVars } from '../theme';

export interface ILinkCardsProps {
  cards: ILinkCard[];
  cardVars: Record<string, CssVars>;
  columns: ColumnSetting;
  fontFamily: string;
  isEditMode: boolean;
  selectedCardId: string | undefined;
  onSelectCard: (id: string) => void;
  onAddCard: () => void;
  onMoveCard: (id: string, delta: number) => void;
  onDeleteCard: (id: string) => void;
}

/**
 * How many columns fit, from the web part's own width (so it works in any
 * section layout): two from 540px, three from 760px when allowed.
 */
function columnsFor(width: number, setting: ColumnSetting): 1 | 2 | 3 {
  const max = setting === 'auto' ? 2 : Number(setting);
  if (max >= 3 && width >= 760) return 3;
  if (max >= 2 && width >= 540) return 2;
  return 1;
}

/**
 * A grid of link cards. Each card with an address is one link — the whole card
 * is clickable — named by its visible text, plus "(opens in a new tab)" when
 * it does. A card without an address is shown as plain text, not a fake link.
 */
export function LinkCards(props: ILinkCardsProps): JSX.Element {
  const { isEditMode, selectedCardId } = props;
  // Readers never see an empty card: one needs text or a working address.
  const cards = isEditMode ? props.cards : props.cards.filter((c) => c.text.trim() !== '' || !!safeHref(c.url));
  const rootRef = React.useRef<HTMLDivElement>(null);
  const [width, setWidth] = React.useState(0);

  React.useLayoutEffect(() => {
    const element = rootRef.current;
    if (!element) return;
    const measure = (): void => setWidth(element.getBoundingClientRect().width);
    measure();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', measure);
      return () => window.removeEventListener('resize', measure);
    }
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  // After a move or delete, put keyboard focus somewhere sensible.
  const pendingFocus = React.useRef<string | undefined>(undefined);
  React.useEffect(() => {
    const key = pendingFocus.current;
    if (!key || !rootRef.current) return;
    pendingFocus.current = undefined;
    const target = rootRef.current.querySelector<HTMLButtonElement>(`[data-focus-key="${key}"]`);
    if (target && !target.disabled) target.focus();
    else rootRef.current.querySelector<HTMLElement>('[data-focus-key="add"]')?.focus();
  });

  const move = (id: string, delta: number): void => {
    pendingFocus.current = `${delta < 0 ? 'earlier' : 'later'}-${id}`;
    props.onMoveCard(id, delta);
  };
  const remove = (id: string): void => {
    const index = cards.findIndex((c) => c.id === id);
    const neighbour = cards[index + 1] || cards[index - 1];
    pendingFocus.current = neighbour ? `edit-${neighbour.id}` : 'add';
    props.onDeleteCard(id);
  };

  const columns = columnsFor(width, props.columns);
  const rootClass = [styles.root, columns === 3 ? styles.cols3 : columns === 2 ? styles.cols2 : '', isEditMode ? styles.editing : '']
    .filter(Boolean)
    .join(' ');

  if (!isEditMode && cards.length === 0) {
    return <div ref={rootRef} className={styles.root} />;
  }

  return (
    <div ref={rootRef} className={rootClass} style={{ '--lc-font': props.fontFamily } as React.CSSProperties}>
      <ul className={styles.grid}>
        {cards.map((card, index) => {
          const href = safeHref(card.url);
          const text = card.text.trim();
          const Icon = getIcon(card.iconName);
          const vars = (props.cardVars[card.id] ?? {}) as React.CSSProperties;
          // With no text, readers see the address itself, so the link always has a name.
          const label = text || (!isEditMode && href ? href : '');
          const content = (
            <>
              <Icon className={styles.icon} aria-hidden="true" focusable="false" />
              <span className={styles.text}>
                {label ? label : <span className={styles.placeholder}>{strings.TextPlaceholder}</span>}
                {href && card.openInNewTab && <span className={styles.srOnly}> {strings.OpensInNewTab}</span>}
                {isEditMode && !href && <span className={styles.hint}>{strings.NoLinkYet}</span>}
              </span>
            </>
          );

          return (
            <li
              key={card.id}
              className={[styles.item, isEditMode && card.id === selectedCardId ? styles.selected : ''].join(' ')}
              style={vars}
            >
              {href ? (
                <a
                  className={styles.card}
                  href={href}
                  target={card.openInNewTab ? '_blank' : undefined}
                  rel={card.openInNewTab ? 'noopener noreferrer' : undefined}
                  data-interception={card.openInNewTab ? 'off' : undefined}
                  onClick={
                    isEditMode
                      ? (event) => {
                          event.preventDefault();
                          props.onSelectCard(card.id);
                        }
                      : undefined
                  }
                >
                  {content}
                </a>
              ) : (
                <div className={styles.card} onClick={isEditMode ? () => props.onSelectCard(card.id) : undefined}>
                  {content}
                </div>
              )}

              {isEditMode && (
                <div className={styles.tools} role="group" aria-label={strings.CardOptions.replace('{0}', text || strings.UntitledCard)}>
                  <button
                    type="button"
                    className={styles.toolButton}
                    title={strings.EditCard}
                    aria-label={strings.EditCard}
                    aria-pressed={card.id === selectedCardId}
                    data-focus-key={`edit-${card.id}`}
                    onClick={() => props.onSelectCard(card.id)}
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
                    onClick={() => move(card.id, -1)}
                  >
                    <ChevronLeft aria-hidden="true" focusable="false" />
                  </button>
                  <button
                    type="button"
                    className={styles.toolButton}
                    title={strings.MoveCardLater}
                    aria-label={strings.MoveCardLater}
                    disabled={index === cards.length - 1}
                    data-focus-key={`later-${card.id}`}
                    onClick={() => move(card.id, 1)}
                  >
                    <ChevronRight aria-hidden="true" focusable="false" />
                  </button>
                  <button
                    type="button"
                    className={styles.toolButton}
                    title={strings.DeleteCard}
                    aria-label={strings.DeleteCard}
                    onClick={() => remove(card.id)}
                  >
                    <Trash2 aria-hidden="true" focusable="false" />
                  </button>
                </div>
              )}
            </li>
          );
        })}
        {isEditMode && (
          <li className={styles.item}>
            <button type="button" className={styles.addTile} data-focus-key="add" onClick={props.onAddCard}>
              <Plus aria-hidden="true" focusable="false" />
              {strings.AddCard}
            </button>
          </li>
        )}
      </ul>
    </div>
  );
}
