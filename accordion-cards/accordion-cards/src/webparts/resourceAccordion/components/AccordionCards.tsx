import * as React from 'react';
import { ChevronDown, Plus } from 'lucide-react';
import * as strings from 'ResourceAccordionWebPartStrings';
import styles from './AccordionCards.module.scss';
import { ResourceCard } from './ResourceCard';
import { TileIcon } from './TileIcon';
import type { HeadingLevel, IconSource, IResourceCard } from '../models';
import type { IResolvedTheme } from '../theme';

export interface IAccordionCardsProps {
  /** Unique per web part instance; used for the ids that wire up ARIA. */
  instanceId: string;
  isEditMode: boolean;
  title: string;
  description: string;
  headingLevel: HeadingLevel;
  iconSource: IconSource;
  iconName: string;
  imageUrl: string;
  showCount: boolean;
  startExpanded: boolean;
  cards: IResourceCard[];
  defaultLinkText: string;
  theme: IResolvedTheme;
  fontFamily: string;
  selectedCardId: string | undefined;
  onSelectCard: (id: string) => void;
  onAddCard: () => void;
  onMoveCard: (id: string, delta: number) => void;
  onDeleteCard: (id: string) => void;
}

/**
 * Card columns follow the web part's own width rather than the window's, so
 * the layout is right in any SharePoint section layout. The thresholds are the
 * accordion widths at which the app's grid switches columns (its sm/640px and
 * xl/1280px window breakpoints, minus the page padding and sidebar), which
 * keeps cards the same size as in the app. Measured in script because the
 * SPFx CSS pipeline strips CSS container queries.
 */
const TWO_COLUMNS_FROM = 592;
const THREE_COLUMNS_FROM = 968;

function columnsFor(width: number): 1 | 2 | 3 {
  if (width >= THREE_COLUMNS_FROM) return 3;
  if (width >= TWO_COLUMNS_FROM) return 2;
  return 1;
}

const CARD_HEADING: Record<HeadingLevel, 'h3' | 'h4' | 'h5'> = { '2': 'h3', '3': 'h4', '4': 'h5' };

/**
 * One collapsible section of resource cards: the accordion group from
 * LifecyclePageTemplate.tsx with its cards, as a single web part.
 *
 * Follows the WAI-ARIA accordion pattern: the header is a real heading
 * containing a button that reports aria-expanded and controls a labelled
 * region, and collapsed content is `hidden` so it leaves the tab order and
 * the accessibility tree.
 */
export function AccordionCards(props: IAccordionCardsProps): JSX.Element {
  const { instanceId, isEditMode, cards, theme, selectedCardId } = props;

  // Authors always land on an open section so the cards are there to edit.
  const [expanded, setExpanded] = React.useState<boolean>(isEditMode || props.startExpanded);
  React.useEffect(() => {
    setExpanded(isEditMode || props.startExpanded);
  }, [isEditMode]);

  React.useEffect(() => {
    if (selectedCardId) setExpanded(true);
  }, [selectedCardId]);

  // After a card is moved or deleted its toolbar button may be re-created or
  // gone; this puts keyboard focus back somewhere sensible.
  const rootRef = React.useRef<HTMLDivElement>(null);
  const pendingFocus = React.useRef<string | undefined>(undefined);
  React.useEffect(() => {
    const key = pendingFocus.current;
    if (!key || !rootRef.current) return;
    pendingFocus.current = undefined;
    const target = rootRef.current.querySelector<HTMLElement>(`[data-focus-key="${key}"]`);
    if (target && !(target as HTMLButtonElement).disabled) {
      target.focus();
    } else {
      rootRef.current.querySelector<HTMLElement>('[data-focus-key="add"]')?.focus();
    }
  });

  const [columns, setColumns] = React.useState<1 | 2 | 3>(1);
  React.useLayoutEffect(() => {
    const element = rootRef.current;
    if (!element) return;
    const measure = (): void => setColumns(columnsFor(element.getBoundingClientRect().width));
    measure();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', measure);
      return () => window.removeEventListener('resize', measure);
    }
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  const handleMove = (id: string, delta: number): void => {
    pendingFocus.current = `${delta < 0 ? 'earlier' : 'later'}-${id}`;
    props.onMoveCard(id, delta);
  };

  const handleDelete = (id: string): void => {
    const index = cards.findIndex((c) => c.id === id);
    const neighbour = cards[index + 1] || cards[index - 1];
    pendingFocus.current = neighbour ? `edit-${neighbour.id}` : 'add';
    props.onDeleteCard(id);
  };

  const handleAdd = (): void => {
    setExpanded(true);
    props.onAddCard();
  };

  const Heading = `h${props.headingLevel}` as 'h2' | 'h3' | 'h4';
  const triggerId = `ra-trigger-${instanceId}`;
  const titleId = `ra-title-${instanceId}`;
  const panelId = `ra-panel-${instanceId}`;
  const title = props.title.trim();
  const description = props.description.trim();
  const count = cards.length;
  const countLabel =
    count === 1 ? strings.ResourceCountLabelSingular : strings.ResourceCountLabel.replace('{0}', String(count));

  const rootClass = [
    styles.root,
    columns === 3 ? styles.cols3 : columns === 2 ? styles.cols2 : '',
    expanded ? styles.expanded : '',
    isEditMode ? styles.editing : ''
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div
      ref={rootRef}
      className={rootClass}
      style={{ ...theme.sectionVars, '--ra-font': props.fontFamily } as React.CSSProperties}
    >
      <div className={styles.section}>
        <Heading className={styles.heading}>
          <button
            type="button"
            id={triggerId}
            className={styles.trigger}
            aria-expanded={expanded}
            aria-controls={panelId}
            onClick={() => setExpanded((open) => !open)}
          >
            <span className={styles.triggerMain}>
              <TileIcon
                className={styles.headerIcon}
                imageClassName={styles.iconImage}
                source={props.iconSource}
                iconName={props.iconName}
                imageUrl={props.imageUrl}
              />
              <span className={styles.headerText}>
                <span id={titleId} className={[styles.title, title ? '' : styles.placeholder].join(' ')}>
                  {title || strings.UntitledSection}
                </span>
                {description && <span className={styles.description}>{description}</span>}
              </span>
              {props.showCount && (
                <span className={styles.badge}>
                  <span aria-hidden="true">{count}</span>
                  <span className={styles.srOnly}>{countLabel}</span>
                </span>
              )}
            </span>
            <ChevronDown className={styles.chevron} aria-hidden="true" focusable="false" />
          </button>
        </Heading>

        <div id={panelId} className={styles.panel} role="region" aria-labelledby={titleId} hidden={!expanded}>
          {count === 0 && !isEditMode ? (
            <p className={styles.empty}>{strings.EmptySection}</p>
          ) : (
            <ul className={styles.grid}>
              {cards.map((card, index) => (
                <li key={card.id} className={styles.gridItem}>
                  <ResourceCard
                    card={card}
                    index={index}
                    total={count}
                    vars={theme.cards[card.id]?.vars ?? {}}
                    headingTag={CARD_HEADING[props.headingLevel] ?? 'h3'}
                    defaultLinkText={props.defaultLinkText}
                    isEditMode={isEditMode}
                    isSelected={card.id === selectedCardId}
                    onSelect={props.onSelectCard}
                    onMove={handleMove}
                    onDelete={handleDelete}
                  />
                </li>
              ))}
              {isEditMode && (
                <li className={styles.gridItem}>
                  <button type="button" className={styles.addTile} data-focus-key="add" onClick={handleAdd}>
                    <Plus aria-hidden="true" focusable="false" />
                    {strings.AddCard}
                  </button>
                </li>
              )}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}
