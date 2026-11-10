/**
 * Property bag shapes for the Accordion Cards web part, plus the palettes
 * and defaults that reproduce the Manager Hub app out of the box.
 */

export type IconSource = 'icon' | 'image';
export type HeadingLevel = '2' | '3' | '4';
export type PaletteName = 'blue' | 'indigo' | 'green' | 'amber' | 'custom';
export type CardAccentMode = 'section' | 'custom';

/** One resource card inside the accordion. */
export interface IResourceCard {
  /** Stable id, used as the React key and to track the card being edited. */
  id: string;
  title: string;
  description: string;
  iconSource: IconSource;
  /** Name from the icon registry (icons.ts). */
  iconName: string;
  /** Server-relative or absolute image URL, used when iconSource is 'image'. */
  imageUrl: string;
  linkUrl: string;
  /** Empty means "use the section's default link text". */
  linkText: string;
  openInNewTab: boolean;
  accentMode: CardAccentMode;
  /** Only used when accentMode is 'custom'. */
  accentColor: string;
}

export interface IAccordionCardsWebPartProps {
  // Header
  title: string;
  description: string;
  headingLevel: HeadingLevel;
  iconSource: IconSource;
  iconName: string;
  imageUrl: string;
  showCount: boolean;
  startExpanded: boolean;

  // Section colours
  palette: PaletteName;
  accentColor: string;
  tintColor: string;
  borderColor: string;
  headerBackground: string;
  headerHoverBackground: string;
  titleColor: string;
  descriptionColor: string;
  chevronColor: string;
  bodyBackground: string;

  // Card defaults
  cardBackground: string;
  cardTitleColor: string;
  cardDescriptionColor: string;
  defaultLinkText: string;

  useInterFont: boolean;

  cards: IResourceCard[];
}

export interface IPalette {
  accent: string;
  tint: string;
  border: string;
}

/** The four lifecycle-stage palettes from the Manager Hub app, verbatim. */
export const PALETTES: Record<Exclude<PaletteName, 'custom'>, IPalette> = {
  blue: { accent: '#2563eb', tint: '#eff6ff', border: '#bfdbfe' },
  indigo: { accent: '#4f46e5', tint: '#eef2ff', border: '#c7d2fe' },
  green: { accent: '#059669', tint: '#ecfdf5', border: '#a7f3d0' },
  amber: { accent: '#d97706', tint: '#fffbeb', border: '#fde68a' }
};

/** Neutral colours the app hard-codes; used whenever a colour field is blank. */
export const DEFAULT_COLORS = {
  headerBackground: '#ffffff',
  headerHoverBackground: '#f8fafc',
  titleColor: '#0f1c36',
  descriptionColor: '#5a6a8a',
  chevronColor: '#5a6a8a',
  bodyBackground: '#ffffff',
  cardBackground: '#fafbfe',
  cardTitleColor: '#0f1c36',
  cardDescriptionColor: '#5a6a8a'
};

export const DEFAULT_LINK_TEXT = 'View Resource';
export const DEFAULT_ICON = 'FileText';

export function createCard(id: string, overrides: Partial<IResourceCard> = {}): IResourceCard {
  return {
    id,
    title: 'Resource title',
    description: 'A short description of this resource.',
    iconSource: 'icon',
    iconName: DEFAULT_ICON,
    imageUrl: '',
    linkUrl: '',
    linkText: '',
    openInNewTab: false,
    accentMode: 'section',
    accentColor: '',
    ...overrides
  };
}

/**
 * Fills in anything missing from a stored property bag, so older or partial
 * configurations (or a hand-edited page) never crash the render.
 */
export function normalizeProperties(props: IAccordionCardsWebPartProps): void {
  const defaults: Partial<IAccordionCardsWebPartProps> = {
    title: '',
    description: '',
    headingLevel: '2',
    iconSource: 'icon',
    iconName: 'Search',
    imageUrl: '',
    showCount: true,
    startExpanded: false,
    palette: 'blue',
    accentColor: PALETTES.blue.accent,
    tintColor: PALETTES.blue.tint,
    borderColor: PALETTES.blue.border,
    ...DEFAULT_COLORS,
    defaultLinkText: DEFAULT_LINK_TEXT,
    useInterFont: true
  };

  const bag = props as unknown as Record<string, unknown>;
  Object.keys(defaults).forEach((key) => {
    if (bag[key] === undefined || bag[key] === null) {
      bag[key] = (defaults as Record<string, unknown>)[key];
    }
  });

  if (!Array.isArray(props.cards)) {
    props.cards = [];
  }
  props.cards = props.cards
    .filter((card) => card && typeof card === 'object')
    .map((card, i) => ({ ...createCard(card.id || `card-${i}-${Date.now()}`), ...card }));
}
