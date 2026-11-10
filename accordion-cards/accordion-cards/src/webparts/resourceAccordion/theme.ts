/**
 * Turns the author's colour choices into the CSS custom properties the
 * component styles read, applying the contrast safeguards in colors.ts on the
 * way. Every adjustment is reported so the property pane can tell the author
 * what was changed and why.
 */

import { colorOr, deriveTints, ensureContrast } from './colors';
import { DEFAULT_COLORS, IAccordionCardsWebPartProps, IResourceCard, PALETTES } from './models';

/** WCAG 2.2 AA: 4.5:1 for normal-size text, 3:1 for meaningful graphics and focus indicators. */
const TEXT_CONTRAST = 4.5;
const GRAPHIC_CONTRAST = 3;

export interface IContrastAdjustment {
  label: string;
  from: string;
  to: string;
}

export type CssVars = Record<string, string>;

export interface IResolvedCardTheme {
  vars: CssVars;
}

export interface IResolvedTheme {
  sectionVars: CssVars;
  cards: Record<string, IResolvedCardTheme>;
  adjustments: IContrastAdjustment[];
}

function track(
  adjustments: IContrastAdjustment[],
  label: string,
  original: string,
  adjusted: string
): string {
  if (original.toLowerCase() !== adjusted.toLowerCase()) {
    adjustments.push({ label, from: original, to: adjusted });
  }
  return adjusted;
}

export function resolveSectionPalette(props: IAccordionCardsWebPartProps): {
  accent: string;
  tint: string;
  border: string;
} {
  if (props.palette !== 'custom' && PALETTES[props.palette]) {
    return PALETTES[props.palette];
  }
  const accent = colorOr(props.accentColor, PALETTES.blue.accent);
  const derived = deriveTints(accent);
  return {
    accent,
    tint: colorOr(props.tintColor, derived.tint),
    border: colorOr(props.borderColor, derived.border)
  };
}

export function resolveTheme(props: IAccordionCardsWebPartProps): IResolvedTheme {
  const adjustments: IContrastAdjustment[] = [];
  const { accent, tint, border } = resolveSectionPalette(props);

  const headerBg = colorOr(props.headerBackground, DEFAULT_COLORS.headerBackground);
  const headerHoverBg = colorOr(props.headerHoverBackground, DEFAULT_COLORS.headerHoverBackground);
  const bodyBg = colorOr(props.bodyBackground, DEFAULT_COLORS.bodyBackground);
  const cardBg = colorOr(props.cardBackground, DEFAULT_COLORS.cardBackground);
  const headerBgs = [headerBg, headerHoverBg];

  const titleRaw = colorOr(props.titleColor, DEFAULT_COLORS.titleColor);
  const descRaw = colorOr(props.descriptionColor, DEFAULT_COLORS.descriptionColor);
  const chevronRaw = colorOr(props.chevronColor, DEFAULT_COLORS.chevronColor);

  const sectionVars: CssVars = {
    '--ra-accent': accent,
    '--ra-tint': tint,
    '--ra-border': border,
    '--ra-header-bg': headerBg,
    '--ra-header-hover-bg': headerHoverBg,
    '--ra-body-bg': bodyBg,
    '--ra-title': track(adjustments, 'Section title', titleRaw, ensureContrast(titleRaw, headerBgs, TEXT_CONTRAST)),
    '--ra-description': track(
      adjustments,
      'Section description',
      descRaw,
      ensureContrast(descRaw, headerBgs, TEXT_CONTRAST)
    ),
    '--ra-chevron': track(
      adjustments,
      'Expand/collapse arrow',
      chevronRaw,
      ensureContrast(chevronRaw, headerBgs, GRAPHIC_CONTRAST)
    ),
    '--ra-badge-text': track(
      adjustments,
      'Resource count',
      accent,
      ensureContrast(accent, [tint], TEXT_CONTRAST)
    ),
    '--ra-empty': ensureContrast(descRaw, [bodyBg], TEXT_CONTRAST),
    '--ra-focus': ensureContrast(accent, [headerBg, headerHoverBg, bodyBg, cardBg], GRAPHIC_CONTRAST)
  };

  const cards: Record<string, IResolvedCardTheme> = {};
  (props.cards || []).forEach((card: IResourceCard, index: number) => {
    let cardAccent = accent;
    let cardTint = tint;
    let cardBorder = border;
    if (card.accentMode === 'custom' && card.accentColor) {
      cardAccent = colorOr(card.accentColor, accent);
      const derived = deriveTints(cardAccent);
      cardTint = derived.tint;
      cardBorder = derived.border;
    }

    // A card's background switches to its tint on hover, so text must pass on both.
    const surfaces = [cardBg, cardTint];
    // Shared colours are reported once for all cards; only a card with its own
    // accent is named individually.
    const isCustom = card.accentMode === 'custom' && !!card.accentColor;
    const name = isCustom ? (card.title ? `“${card.title}”` : `Card ${index + 1}`) : 'Card';
    const titleColor = colorOr(props.cardTitleColor, DEFAULT_COLORS.cardTitleColor);
    const descColor = colorOr(props.cardDescriptionColor, DEFAULT_COLORS.cardDescriptionColor);

    cards[card.id] = {
      vars: {
        '--rc-accent': cardAccent,
        '--rc-tint': cardTint,
        '--rc-border': cardBorder,
        '--rc-bg': cardBg,
        '--rc-title': track(
          adjustments,
          `${name} title`,
          titleColor,
          ensureContrast(titleColor, surfaces, TEXT_CONTRAST)
        ),
        '--rc-description': track(
          adjustments,
          `${name} description`,
          descColor,
          ensureContrast(descColor, surfaces, TEXT_CONTRAST)
        ),
        '--rc-link': track(
          adjustments,
          isCustom ? `${name} link` : 'Card links',
          cardAccent,
          ensureContrast(cardAccent, surfaces, TEXT_CONTRAST)
        ),
        '--rc-focus': ensureContrast(cardAccent, surfaces.concat(bodyBg), GRAPHIC_CONTRAST)
      }
    };
  });

  return { sectionVars, cards, adjustments: dedupe(adjustments) };
}

/** Card text colours are shared, so the same adjustment can be reported once per card. */
function dedupe(list: IContrastAdjustment[]): IContrastAdjustment[] {
  const seen = new Set<string>();
  return list.filter((item) => {
    const key = `${item.label}|${item.from}|${item.to}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
