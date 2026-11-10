/**
 * Turns the author's colour choices into CSS custom properties, checking each
 * against the backgrounds it can appear on. A card's background switches to a
 * tint on hover and focus, so text, icon and focus ring must pass on both.
 *
 *  - Text: 4.5:1 (WCAG 2.2 AA, normal-size text).
 *  - Icon and focus ring: 3:1 (WCAG 2.2 AA, non-text contrast).
 *
 * Anything short is nudged darker (or lighter on dark cards) by the smallest
 * amount that passes; colours that already pass are untouched. Each change is
 * reported so the property pane can explain it.
 */

import { colorOr, deriveTints, ensureContrast } from './colors';
import { DEFAULT_CARD_BACKGROUND, DEFAULT_TEXT_COLOR, ILinkCardsWebPartProps, IPalette, PALETTES } from './models';

/** Small margin so values rounded for CSS still pass. */
const TEXT_CONTRAST = 4.5 + 0.05;
const GRAPHIC_CONTRAST = 3 + 0.05;

export type CssVars = Record<string, string>;

export interface IContrastAdjustment {
  label: string;
  from: string;
  to: string;
}

export interface ILinkCardsTheme {
  cards: Record<string, CssVars>;
  adjustments: IContrastAdjustment[];
}

export function resolveSectionPalette(props: ILinkCardsWebPartProps): IPalette {
  if (props.palette !== 'custom' && PALETTES[props.palette]) return PALETTES[props.palette];
  const accent = colorOr(props.iconColor, PALETTES.blue.accent);
  return { accent, ...deriveTints(accent) };
}

export function resolveTheme(props: ILinkCardsWebPartProps): ILinkCardsTheme {
  const adjustments: IContrastAdjustment[] = [];
  const seen = new Set<string>();
  const report = (label: string, from: string, to: string): string => {
    const key = `${label}|${from}|${to}`;
    if (from.toLowerCase() !== to.toLowerCase() && !seen.has(key)) {
      seen.add(key);
      adjustments.push({ label, from, to });
    }
    return to;
  };

  const section = resolveSectionPalette(props);
  const cardBg = colorOr(props.cardBackground, DEFAULT_CARD_BACKGROUND);
  const textRaw = colorOr(props.textColor, DEFAULT_TEXT_COLOR);

  const cards: Record<string, CssVars> = {};
  (props.cards || []).forEach((card, index) => {
    const custom = card.iconColorMode === 'custom' && !!card.iconColor;
    const palette = custom ? { accent: colorOr(card.iconColor, section.accent), ...deriveTints(colorOr(card.iconColor, section.accent)) } : section;
    const surfaces = [cardBg, palette.tint];
    const who = custom ? `“${card.text.trim() || `Card ${index + 1}`}” icon` : 'Icons';

    cards[card.id] = {
      '--lc-bg': cardBg,
      '--lc-hover-bg': palette.tint,
      '--lc-hover-border': palette.border,
      '--lc-text': report('Card text', textRaw, ensureContrast(textRaw, surfaces, TEXT_CONTRAST)),
      '--lc-icon': report(who, palette.accent, ensureContrast(palette.accent, surfaces, GRAPHIC_CONTRAST)),
      // The ring sits just outside the card, over the page (white) as well as the card.
      '--lc-focus': ensureContrast(palette.accent, surfaces.concat('#ffffff'), GRAPHIC_CONTRAST),
      '--lc-hint': ensureContrast('#5a6a8a', surfaces, TEXT_CONTRAST)
    };
  });

  return { cards, adjustments };
}
