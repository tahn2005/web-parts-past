/**
 * Property bag for the Custom Width Header web part, its palettes and defaults.
 */

export type HeaderLayout = 'page' | 'resource';
export type HeaderHeight = 'auto' | 'medium' | 'large';
/** 'auto' picks 1 or 2 from the page around it; see headingLevel.ts. */
export type HeadingLevel = 'auto' | '1' | '2';
export type IconSource = 'none' | 'icon' | 'image';
export type BackgroundType = 'color' | 'image';
export type ImageTextStyle = 'overlay' | 'panel';
export type ImagePosition = 'center' | 'top' | 'bottom';
export type PaletteName = 'blue' | 'indigo' | 'green' | 'amber' | 'custom';

export interface ICustomWidthHeaderWebPartProps {
  // Layout
  layout: HeaderLayout;
  height: HeaderHeight;
  headingLevel: HeadingLevel;

  // Text
  showBackLink: boolean;
  backLinkText: string;
  /** Empty means "this site's home page". */
  backLinkUrl: string;
  eyebrow: string;
  title: string;
  description: string;

  // Icon
  iconSource: IconSource;
  iconName: string;
  iconImageUrl: string;

  // Background
  backgroundType: BackgroundType;
  palette: PaletteName;
  backgroundColor: string;
  useGradient: boolean;
  backgroundImageUrl: string;
  imagePosition: ImagePosition;
  imageTextStyle: ImageTextStyle;
  /** Overlay colour in "overlay" mode; the panel colour in "panel" mode. */
  overlayColor: string;
  /** 0–100. A floor is enforced so text always meets contrast (see theme.ts). */
  overlayOpacity: number;

  textColor: string;
  useManagerHubFonts: boolean;
}

/** Hero colours of the four lifecycle stages in the Manager Hub app. */
export const PALETTES: Record<Exclude<PaletteName, 'custom'>, string> = {
  blue: '#2563eb',
  indigo: '#4f46e5',
  green: '#059669',
  amber: '#d97706'
};

export const DEFAULTS: ICustomWidthHeaderWebPartProps = {
  layout: 'page',
  height: 'auto',
  headingLevel: 'auto',
  showBackLink: true,
  backLinkText: 'Home',
  backLinkUrl: '',
  eyebrow: '',
  title: 'Page title',
  description: 'A short description of what this page covers.',
  iconSource: 'icon',
  iconName: 'Layers',
  iconImageUrl: '',
  backgroundType: 'color',
  palette: 'blue',
  backgroundColor: PALETTES.blue,
  useGradient: false,
  backgroundImageUrl: '',
  imagePosition: 'center',
  imageTextStyle: 'overlay',
  overlayColor: '#0f1c36',
  overlayOpacity: 55,
  textColor: '#ffffff',
  useManagerHubFonts: true
};

/** Fills in anything missing so partial or older configurations never break the render. */
export function normalizeProperties(props: ICustomWidthHeaderWebPartProps): void {
  const bag = props as unknown as Record<string, unknown>;
  (Object.keys(DEFAULTS) as Array<keyof ICustomWidthHeaderWebPartProps>).forEach((key) => {
    if (bag[key] === undefined || bag[key] === null) {
      bag[key] = DEFAULTS[key];
    }
  });
  const opacity = Number(props.overlayOpacity);
  props.overlayOpacity = Number.isFinite(opacity) ? Math.min(100, Math.max(0, opacity)) : DEFAULTS.overlayOpacity;
}
