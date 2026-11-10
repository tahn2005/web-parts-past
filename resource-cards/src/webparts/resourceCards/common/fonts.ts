/**
 * Registers the Resource Cards typeface (Inter, 400 and 600) from files bundled
 * with the web part, so nothing is fetched from a third-party font CDN.
 *
 * The family is registered under its own name so it can't clash with an
 * "Inter" the tenant or another web part may already have loaded.
 */

import inter400 from '../fonts/inter-latin-400-normal.woff2';
import inter600 from '../fonts/inter-latin-600-normal.woff2';

export const CARD_FONT_FAMILY = "'Resource Cards Inter', Inter, 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";

const STYLE_ID = 'resource-cards-inter-font';

export function registerCardFont(): void {
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) {
    return;
  }
  const face = (url: string, weight: number): string =>
    `@font-face{font-family:'Resource Cards Inter';font-style:normal;font-weight:${weight};` +
    `font-display:swap;src:url('${url}') format('woff2');}`;

  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = face(inter400, 400) + face(inter600, 600);
  document.head.appendChild(style);
}
