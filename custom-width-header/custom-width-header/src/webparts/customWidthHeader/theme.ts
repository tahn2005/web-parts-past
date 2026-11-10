/**
 * Works out the header's colours so that every piece of text meets WCAG 2.2 AA
 * contrast, whatever the author picks — including text over a photo.
 *
 * Text over an image is the classic failure: contrast depends on whichever
 * pixels end up behind the letters. Here the overlay is checked against the
 * two worst cases — a pure white and a pure black image — and made just strong
 * enough to pass both. Every real image lies between those extremes, so the
 * text passes over any photo, at any crop and any screen size.
 *
 * Order of adjustments, least visible first:
 *  1. Secondary text (description, label, back link) is drawn at the app's
 *     partial opacity; raise that opacity if needed.
 *  2. Then darken (or lighten) the background colour, or strengthen the
 *     image overlay.
 *  3. Last resort: shift the text colour itself.
 * Each change is reported so the property pane can show it.
 */

import { colorOr, contrastRatio, ensureContrast, IRgb, mix, parseColor, toHex } from './colors';
import { DEFAULTS, ICustomWidthHeaderWebPartProps, PALETTES } from './models';

const WHITE: IRgb = { r: 255, g: 255, b: 255 };
const BLACK: IRgb = { r: 0, g: 0, b: 0 };

/**
 * Large text (the title is 25.6–28px bold) needs 3:1; everything else 4.5:1.
 * Targets carry a small margin so values rounded for CSS still pass.
 */
const MARGIN = 0.05;
const LARGE_TEXT = 3 + MARGIN;
const NORMAL_TEXT = 4.5 + MARGIN;

/** The app draws its gradient as the accent at 0xee and 0xcc alpha. */
const GRADIENT_START_ALPHA = 0xee / 255;
const GRADIENT_END_ALPHA = 0xcc / 255;

/** The app's opacities for secondary text, relative to the text colour. */
const SECONDARY = {
  description: 0.82,
  eyebrow: 0.8,
  backLink: 0.75
};

export interface IContrastAdjustment {
  label: string;
  from: string;
  to: string;
}

export type CssVars = Record<string, string>;

export interface IHeaderTheme {
  vars: CssVars;
  adjustments: IContrastAdjustment[];
  /** Overlay opacity actually used (0–1), after any raise for contrast. */
  overlayOpacity: number;
}

/** `fg` drawn at `alpha` over `bg`. */
function over(fg: IRgb, alpha: number, bg: IRgb): IRgb {
  return mix(bg, fg, alpha);
}

function rgba(c: IRgb, alpha: number): string {
  const a = Math.round(alpha * 1000) / 1000;
  return `rgba(${Math.round(c.r)}, ${Math.round(c.g)}, ${Math.round(c.b)}, ${a})`;
}

function worst(fg: IRgb, alpha: number, backgrounds: IRgb[]): number {
  return Math.min(...backgrounds.map((bg) => contrastRatio(over(fg, alpha, bg), bg)));
}

function isLight(c: IRgb): boolean {
  return contrastRatio(c, BLACK) > contrastRatio(c, WHITE);
}

export function resolveBaseColor(props: ICustomWidthHeaderWebPartProps): string {
  if (props.palette !== 'custom' && PALETTES[props.palette]) return PALETTES[props.palette];
  return colorOr(props.backgroundColor, PALETTES.blue);
}

export function resolveHeaderTheme(props: ICustomWidthHeaderWebPartProps): IHeaderTheme {
  const adjustments: IContrastAdjustment[] = [];
  let text = parseColor(props.textColor) ?? WHITE;
  const textLight = isLight(text);
  /** Direction that increases contrast with the text. */
  const away = textLight ? BLACK : WHITE;

  const isImage = props.backgroundType === 'image';
  const isPanel = isImage && props.imageTextStyle === 'panel';

  // ---- the colours text can end up on ---------------------------------
  let base = parseColor(resolveBaseColor(props)) ?? parseColor(PALETTES.blue)!;
  let overlay = parseColor(props.overlayColor) ?? parseColor(DEFAULTS.overlayColor)!;
  let overlayAlpha = isPanel ? 1 : Math.min(1, Math.max(0, props.overlayOpacity / 100));

  const surfaces = (): IRgb[] => {
    if (isImage) {
      // Panel: solid. Overlay: whatever the photo is, it lies between these two.
      return isPanel ? [overlay] : [over(overlay, overlayAlpha, WHITE), over(overlay, overlayAlpha, BLACK)];
    }
    // Colour backgrounds are flattened onto white, so the result never depends
    // on the SharePoint section colour behind the web part.
    return props.useGradient
      ? [over(base, GRADIENT_START_ALPHA, WHITE), over(base, GRADIENT_END_ALPHA, WHITE)]
      : [base];
  };

  // Only text that is actually shown sets the bar: a header with just a title
  // (large text, 3:1) keeps colours that small text would have needed changed.
  const hasSmallText =
    props.description.trim() !== '' ||
    props.eyebrow.trim() !== '' ||
    (props.showBackLink && props.backLinkText.trim() !== '');
  const passes = (alpha: number, ratio: number): boolean => worst(text, alpha, surfaces()) >= ratio;
  const allPass = (): boolean => passes(1, hasSmallText ? NORMAL_TEXT : LARGE_TEXT);

  // ---- step 2: background / overlay, only if even solid text fails --------
  if (!allPass()) {
    if (isImage && !isPanel) {
      const from = overlayAlpha;
      while (overlayAlpha < 1 && !allPass()) overlayAlpha = Math.min(1, overlayAlpha + 0.01);
      adjustments.push({
        label: 'Image overlay strength',
        from: `${Math.round(from * 100)}%`,
        to: `${Math.round(overlayAlpha * 100)}%`
      });
      if (!allPass()) {
        const original = toHex(overlay);
        const start = overlay;
        for (let step = 1; step <= 50 && !allPass(); step++) overlay = mix(start, away, step / 50);
        adjustments.push({ label: 'Image overlay color', from: original, to: toHex(overlay) });
      }
    } else if (isPanel) {
      const original = toHex(overlay);
      const start = overlay;
      for (let step = 1; step <= 50 && !allPass(); step++) overlay = mix(start, away, step / 50);
      adjustments.push({ label: 'Text panel color', from: original, to: toHex(overlay) });
    } else {
      const original = toHex(base);
      const start = base;
      for (let step = 1; step <= 50 && !allPass(); step++) base = mix(start, away, step / 50);
      adjustments.push({ label: 'Background color', from: original, to: toHex(base) });
    }
  }

  // ---- step 3: last resort, the text colour itself ------------------------
  if (!allPass()) {
    const original = toHex(text);
    text = parseColor(ensureContrast(original, surfaces().map(toHex), hasSmallText ? NORMAL_TEXT : LARGE_TEXT)) ?? text;
    adjustments.push({ label: 'Text color', from: original, to: toHex(text) });
  }

  // ---- step 1: secondary text opacity -------------------------------------
  const secondary = (label: string, appAlpha: number): number => {
    let alpha = appAlpha;
    while (alpha < 1 && !passes(alpha, NORMAL_TEXT)) alpha = Math.min(1, Math.round((alpha + 0.01) * 100) / 100);
    if (alpha !== appAlpha && hasSmallText) {
      adjustments.push({ label, from: `${Math.round(appAlpha * 100)}% opacity`, to: `${Math.round(alpha * 100)}% opacity` });
    }
    return alpha;
  };
  const descriptionAlpha = secondary('Description text', SECONDARY.description);
  const eyebrowAlpha = secondary('Label text', SECONDARY.eyebrow);
  const backAlpha = secondary('Back link', SECONDARY.backLink);

  const [first, second] = surfaces();
  const background = isImage
    ? 'transparent'
    : props.useGradient
      ? `linear-gradient(135deg, ${toHex(first)} 0%, ${toHex(second)} 100%)`
      : toHex(base);

  return {
    overlayOpacity: overlayAlpha,
    adjustments,
    vars: {
      '--ph-background': background,
      // A solid colour underneath too, so tools (and forced-colors fallbacks)
      // always have something definite to measure against.
      '--ph-background-color': toHex(isImage ? (isPanel ? overlay : over(overlay, overlayAlpha, BLACK)) : first),
      '--ph-overlay': isImage && !isPanel ? rgba(overlay, overlayAlpha) : 'transparent',
      '--ph-panel': isPanel ? toHex(overlay) : 'transparent',
      '--ph-text': toHex(text),
      '--ph-description': rgba(text, descriptionAlpha),
      '--ph-eyebrow': rgba(text, eyebrowAlpha),
      '--ph-back': rgba(text, backAlpha),
      '--ph-icon-tile': rgba(text, 0.2)
    }
  };
}

/** Exposed for tests: the colours a given theme's text sits on. */
export function contrastReport(props: ICustomWidthHeaderWebPartProps): { worstTitle: number; worstSmall: number } {
  const theme = resolveHeaderTheme(props);
  const text = parseColor(theme.vars['--ph-text'])!;
  const isImage = props.backgroundType === 'image';
  const isPanel = isImage && props.imageTextStyle === 'panel';
  const bgs = isImage
    ? isPanel
      ? [parseColor(theme.vars['--ph-panel'])!]
      : [WHITE, BLACK].map((b) => over(parseColor(theme.vars['--ph-overlay'].replace(/rgba\((\d+), (\d+), (\d+).*/, 'rgb($1,$2,$3)'))!, theme.overlayOpacity, b))
    : (theme.vars['--ph-background'].match(/#[0-9a-f]{6}/g) || []).map((h) => parseColor(h)!);
  const alphaOf = (v: string): number => Number(/, ([\d.]+)\)$/.exec(v)![1]);
  const small = Math.min(
    ...['--ph-description', '--ph-eyebrow', '--ph-back'].map((k) => worst(text, alphaOf(theme.vars[k]), bgs))
  );
  return { worstTitle: worst(text, 1, bgs), worstSmall: small };
}
