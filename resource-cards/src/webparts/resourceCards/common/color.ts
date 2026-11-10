/**
 * Colour helpers.
 *
 * The original app gives every lifecycle page an accent plus two derived tints:
 * `light` (hover background, icon tile) and `border` (hover border). The four
 * page palettes are reproduced exactly; any other accent gets tints mixed from
 * it the same way.
 *
 * Text colours are checked against WCAG 2.2 contrast (4.5:1 for normal text)
 * and darkened only as far as needed, so an author can pick any accent without
 * producing an unreadable link.
 */

export interface IAccentPalette {
  /** Icon colour — the accent exactly as chosen. */
  accent: string;
  /** Icon tile and hover background. */
  light: string;
  /** Hover border. */
  border: string;
  /** Link text and focus ring: the accent, darkened if needed to reach 4.5:1. */
  text: string;
  /** Card description text, same treatment. */
  muted: string;
}

export interface IAccentPreset {
  name: string;
  accent: string;
  light: string;
  border: string;
}

/** Palettes of the original lifecycle pages, plus a navy. */
export const ACCENT_PRESETS: IAccentPreset[] = [
  { name: 'Blue (Hire & Onboard)', accent: '#2563eb', light: '#eff6ff', border: '#bfdbfe' },
  { name: 'Indigo (Lead & Develop)', accent: '#4f46e5', light: '#eef2ff', border: '#c7d2fe' },
  { name: 'Green (Support & Retain)', accent: '#059669', light: '#ecfdf5', border: '#a7f3d0' },
  { name: 'Amber (Exit & Transition)', accent: '#d97706', light: '#fffbeb', border: '#fde68a' },
  { name: 'Navy', accent: '#1a3461', light: '#e8edf7', border: '#c3cee3' }
];

export const DEFAULT_ACCENT = ACCENT_PRESETS[0].accent;

/** Resting card background, from the React app. */
export const CARD_BACKGROUND = '#fafbfe';
/** Card description colour, from the React app. */
export const MUTED_TEXT = '#5a6a8a';
/** WCAG 2.2 SC 1.4.3 minimum for normal-size text. */
export const MIN_TEXT_CONTRAST = 4.5;

type Rgb = [number, number, number];

const HEX_PATTERN = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i;

/** Normalises "#abc", "abc", "#AABBCC" to "#aabbcc"; returns '' if invalid. */
export function normalizeHex(value: string | undefined): string {
  const match = HEX_PATTERN.exec((value || '').trim());
  if (!match) {
    return '';
  }
  let hex = match[1].toLowerCase();
  if (hex.length === 3) {
    hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
  }
  return `#${hex}`;
}

function toRgb(hex: string): Rgb {
  const value = parseInt(normalizeHex(hex).slice(1), 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function toHex(rgb: Rgb): string {
  return `#${rgb
    .map((channel) => {
      const clamped = Math.max(0, Math.min(255, Math.round(channel)));
      return (clamped < 16 ? '0' : '') + clamped.toString(16);
    })
    .join('')}`;
}

/** Mixes `amount` (0–1) of `other` into `color`. */
function mix(color: string, other: string, amount: number): string {
  const a = toRgb(color);
  const b = toRgb(other);
  return toHex([
    a[0] + (b[0] - a[0]) * amount,
    a[1] + (b[1] - a[1]) * amount,
    a[2] + (b[2] - a[2]) * amount
  ]);
}

function relativeLuminance(hex: string): number {
  const channels = toRgb(hex).map((channel) => {
    const c = channel / 255;
    return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

export function contrastRatio(foreground: string, background: string): number {
  const a = relativeLuminance(foreground);
  const b = relativeLuminance(background);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

/**
 * Returns `color` unchanged if it already meets `minimum` against every
 * background, otherwise the least-darkened version of it that does.
 */
export function ensureContrast(color: string, backgrounds: string[], minimum: number = MIN_TEXT_CONTRAST): string {
  const passes = (candidate: string): boolean =>
    backgrounds.every((background) => contrastRatio(candidate, background) >= minimum);

  if (passes(color)) {
    return color;
  }
  for (let step = 1; step <= 50; step++) {
    const candidate = mix(color, '#000000', step / 50);
    if (passes(candidate)) {
      return candidate;
    }
  }
  return '#000000';
}

/**
 * Builds the full palette for an accent. Unknown or empty values fall back to
 * the default accent so a typo can never break rendering.
 */
export function getPalette(accentColor: string | undefined): IAccentPalette {
  const accent = normalizeHex(accentColor) || DEFAULT_ACCENT;
  const preset = ACCENT_PRESETS.filter((p) => p.accent === accent)[0];

  // Tailwind's 50 and 200 shades sit at roughly 7% and 27% of the base colour.
  const light = preset ? preset.light : mix(accent, '#ffffff', 0.93);
  const border = preset ? preset.border : mix(accent, '#ffffff', 0.73);

  const backgrounds = [CARD_BACKGROUND, light];
  return {
    accent,
    light,
    border,
    text: ensureContrast(accent, backgrounds),
    muted: ensureContrast(MUTED_TEXT, backgrounds)
  };
}
