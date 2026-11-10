/**
 * Colour maths for the web part.
 *
 * Authors can pick any colour, so every colour that carries text or meaning is
 * checked against the backgrounds it can appear on and, if it falls short of
 * WCAG 2.2 AA, nudged darker (or lighter, on dark backgrounds) by the smallest
 * amount that passes. Colours that already pass are returned untouched, which
 * is why the default palettes render exactly as designed.
 */

export interface IRgb {
  r: number;
  g: number;
  b: number;
}

const WHITE: IRgb = { r: 255, g: 255, b: 255 };
const BLACK: IRgb = { r: 0, g: 0, b: 0 };

/** Parses #rgb, #rrggbb, #rrggbbaa, rgb() and rgba(). Returns undefined otherwise. */
export function parseColor(value: string | undefined): IRgb | undefined {
  const input = (value ?? '').trim().toLowerCase();
  if (!input) return undefined;

  const hex = /^#?([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/.exec(input);
  if (hex) {
    let digits = hex[1];
    if (digits.length === 3) {
      digits = digits.split('').map((d) => d + d).join('');
    }
    return {
      r: parseInt(digits.slice(0, 2), 16),
      g: parseInt(digits.slice(2, 4), 16),
      b: parseInt(digits.slice(4, 6), 16)
    };
  }

  const rgb = /^rgba?\(\s*(\d{1,3})[\s,]+(\d{1,3})[\s,]+(\d{1,3})/.exec(input);
  if (rgb) {
    return {
      r: Math.min(255, Number(rgb[1])),
      g: Math.min(255, Number(rgb[2])),
      b: Math.min(255, Number(rgb[3]))
    };
  }

  return undefined;
}

export function toHex({ r, g, b }: IRgb): string {
  const part = (n: number): string => Math.round(n).toString(16).padStart(2, '0');
  return `#${part(r)}${part(g)}${part(b)}`;
}

/** Returns a valid hex colour, falling back when the input is blank or unparseable. */
export function colorOr(value: string | undefined, fallback: string): string {
  const parsed = parseColor(value);
  return parsed ? toHex(parsed) : fallback;
}

/** Mixes `color` toward `target`; amount 0 = color, 1 = target. */
export function mix(color: IRgb, target: IRgb, amount: number): IRgb {
  return {
    r: color.r + (target.r - color.r) * amount,
    g: color.g + (target.g - color.g) * amount,
    b: color.b + (target.b - color.b) * amount
  };
}

function luminance({ r, g, b }: IRgb): number {
  const channel = (c: number): number => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

export function contrastRatio(a: IRgb, b: IRgb): number {
  const la = luminance(a);
  const lb = luminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

function worstContrast(fg: IRgb, backgrounds: IRgb[]): number {
  return Math.min(...backgrounds.map((bg) => contrastRatio(fg, bg)));
}

/**
 * Returns `foreground` unchanged if it meets `minRatio` against every
 * background; otherwise the closest shade of it that does.
 */
export function ensureContrast(foreground: string, backgrounds: string[], minRatio: number): string {
  const fg = parseColor(foreground);
  const bgs = backgrounds.map(parseColor).filter((c): c is IRgb => c !== undefined);
  if (!fg || bgs.length === 0) return foreground;
  if (worstContrast(fg, bgs) >= minRatio) return toHex(fg);

  for (let step = 1; step <= 50; step++) {
    const amount = step / 50;
    const darker = mix(fg, BLACK, amount);
    if (worstContrast(darker, bgs) >= minRatio) return toHex(darker);
    const lighter = mix(fg, WHITE, amount);
    if (worstContrast(lighter, bgs) >= minRatio) return toHex(lighter);
  }

  return worstContrast(BLACK, bgs) >= worstContrast(WHITE, bgs) ? '#000000' : '#ffffff';
}

/** Soft background and border tints derived from a single accent colour. */
export function deriveTints(accent: string): { tint: string; border: string } {
  const base = parseColor(accent) ?? { r: 37, g: 99, b: 235 };
  return {
    tint: toHex(mix(base, WHITE, 0.92)),
    border: toHex(mix(base, WHITE, 0.7))
  };
}
