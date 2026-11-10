/**
 * Data model for the Resource Cards web part.
 *
 * Everything an author can change is stored on the web part's properties, so
 * SharePoint persists it with the page — no list or external storage needed.
 */

export type CardMedia = 'icon' | 'image';
export type ImageSize = 'small' | 'medium' | 'large';

/** Image box size for each option. Small matches the icon tile exactly. */
export const IMAGE_SIZES: { [size in ImageSize]: string } = {
  small: '2rem',
  medium: '3rem',
  large: '4rem'
};

export interface IResourceCard {
  /** Stable id, used as the React key so cards keep focus when reordered. */
  id: string;
  /** Show a Lucide icon or a custom image at the top of the card. */
  mediaType: CardMedia;
  /** Lucide icon name; see ./icons.ts for the available set. */
  icon: string;
  /** Image address (uploaded to Site Assets or pasted). Used when mediaType is 'image'. */
  imageUrl: string;
  /** Alternative text. Empty means decorative (the card title already says what it is). */
  imageAlt: string;
  imageSize: ImageSize;
  title: string;
  description: string;
  /** Visible link label. Defaults to "View Resource". */
  linkText: string;
  /** Destination. Leave empty to hide the link on the published page. */
  url: string;
  openInNewTab: boolean;
  /** Optional per-card accent (#rrggbb). Empty means "use the web part accent". */
  accentColor: string;
}

export type HeadingLevel = 2 | 3 | 4 | 5 | 6;

export interface IResourceCardsWebPartProps {
  cards: IResourceCard[];
  /** Web part accent (#rrggbb). Each card can override it. */
  accentColor: string;
  /**
   * Heading level for card titles. Purely semantic — every level looks the
   * same — so authors can keep the page outline in order.
   */
  headingLevel: HeadingLevel;
  /** Use the Resource Cards typeface (Inter) rather than the site font. */
  useInterFont: boolean;
}

export const DEFAULT_LINK_TEXT = 'View Resource';
export const DEFAULT_ICON = 'FileText';

export function createCardId(): string {
  return `card-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function createCard(): IResourceCard {
  return {
    id: createCardId(),
    mediaType: 'icon',
    icon: DEFAULT_ICON,
    imageUrl: '',
    imageAlt: '',
    imageSize: 'small',
    title: 'New card',
    description: 'Describe what people will find when they follow this link.',
    linkText: DEFAULT_LINK_TEXT,
    url: '',
    openInNewTab: false,
    accentColor: ''
  };
}

/**
 * Fills in anything missing from stored cards, so content saved by an older
 * version of the web part (or edited by hand) always renders.
 */
export function normalizeCards(cards: unknown): IResourceCard[] {
  if (!Array.isArray(cards)) {
    return [];
  }
  return cards
    .filter((card): card is Partial<IResourceCard> => !!card && typeof card === 'object')
    .map((card) => ({
      id: typeof card.id === 'string' && card.id ? card.id : createCardId(),
      mediaType: card.mediaType === 'image' ? 'image' : 'icon',
      icon: typeof card.icon === 'string' && card.icon ? card.icon : DEFAULT_ICON,
      imageUrl: typeof card.imageUrl === 'string' ? card.imageUrl : '',
      imageAlt: typeof card.imageAlt === 'string' ? card.imageAlt : '',
      imageSize: card.imageSize === 'medium' || card.imageSize === 'large' ? card.imageSize : 'small',
      title: typeof card.title === 'string' ? card.title : '',
      description: typeof card.description === 'string' ? card.description : '',
      linkText: typeof card.linkText === 'string' ? card.linkText : DEFAULT_LINK_TEXT,
      url: typeof card.url === 'string' ? card.url : '',
      openInNewTab: !!card.openInNewTab,
      accentColor: typeof card.accentColor === 'string' ? card.accentColor : ''
    }));
}
