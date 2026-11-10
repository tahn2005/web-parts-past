/**
 * Property bag for the Link Cards web part, its palettes and defaults.
 */

export type PaletteName = 'blue' | 'indigo' | 'green' | 'amber' | 'custom';
export type IconColorMode = 'section' | 'custom';
export type ColumnSetting = 'auto' | '1' | '2' | '3';

/** One link card. */
export interface ILinkCard {
  /** Stable id, used as the React key and to track the card being edited. */
  id: string;
  text: string;
  /** Empty means the card is shown but isn't a link (e.g. "coming soon"). */
  url: string;
  openInNewTab: boolean;
  /** Name from the icon registry (icons.ts). */
  iconName: string;
  iconColorMode: IconColorMode;
  /** Only used when iconColorMode is 'custom'. */
  iconColor: string;
}

export interface ILinkCardsWebPartProps {
  palette: PaletteName;
  /** The custom icon colour when palette is 'custom'. */
  iconColor: string;
  textColor: string;
  cardBackground: string;
  columns: ColumnSetting;
  useInterFont: boolean;
  cards: ILinkCard[];
}

export interface IPalette {
  accent: string;
  tint: string;
  border: string;
}

/** Icon colour plus the hover tint and border that go with it. */
export const PALETTES: Record<Exclude<PaletteName, 'custom'>, IPalette> = {
  blue: { accent: '#2563eb', tint: '#eff6ff', border: '#bfdbfe' },
  indigo: { accent: '#4f46e5', tint: '#eef2ff', border: '#c7d2fe' },
  green: { accent: '#059669', tint: '#ecfdf5', border: '#a7f3d0' },
  amber: { accent: '#d97706', tint: '#fffbeb', border: '#fde68a' }
};

export const DEFAULT_TEXT_COLOR = '#0f1c36';
export const DEFAULT_CARD_BACKGROUND = '#fafbfe';
export const DEFAULT_ICON = 'ExternalLink';

export function createCard(id: string, overrides: Partial<ILinkCard> = {}): ILinkCard {
  return {
    id,
    text: 'Link title',
    url: '',
    openInNewTab: true,
    iconName: DEFAULT_ICON,
    iconColorMode: 'section',
    iconColor: '',
    ...overrides
  };
}

const DEFAULTS: Omit<ILinkCardsWebPartProps, 'cards'> = {
  palette: 'blue',
  iconColor: PALETTES.blue.accent,
  textColor: DEFAULT_TEXT_COLOR,
  cardBackground: DEFAULT_CARD_BACKGROUND,
  columns: 'auto',
  useInterFont: true
};

/** Fills in anything missing so partial or older configurations never break the render. */
export function normalizeProperties(props: ILinkCardsWebPartProps): void {
  const bag = props as unknown as Record<string, unknown>;
  (Object.keys(DEFAULTS) as Array<keyof typeof DEFAULTS>).forEach((key) => {
    if (bag[key] === undefined || bag[key] === null) bag[key] = DEFAULTS[key];
  });
  if (!Array.isArray(props.cards)) props.cards = [];
  props.cards = props.cards
    .filter((card) => card && typeof card === 'object')
    .map((card, i) => ({ ...createCard(card.id || `card-${i}`), ...card }))
    .map((card) => ({
      ...card,
      text: typeof card.text === 'string' ? card.text : '',
      url: typeof card.url === 'string' ? card.url : ''
    }));
}
