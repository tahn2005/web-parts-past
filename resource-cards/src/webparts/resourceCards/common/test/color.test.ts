import { ACCENT_PRESETS, CARD_BACKGROUND, contrastRatio, getPalette, MIN_TEXT_CONTRAST, normalizeHex } from '../color';

describe('getPalette', () => {
  it('reproduces the original page palettes exactly', () => {
    const blue = getPalette('#2563eb');
    expect(blue.light).toBe('#eff6ff');
    expect(blue.border).toBe('#bfdbfe');
    expect(blue.text).toBe('#2563eb');
  });

  it('keeps link and description text at 4.5:1 or better for every preset', () => {
    ACCENT_PRESETS.forEach((preset) => {
      const palette = getPalette(preset.accent);
      [CARD_BACKGROUND, palette.light].forEach((background) => {
        expect(contrastRatio(palette.text, background)).toBeGreaterThanOrEqual(MIN_TEXT_CONTRAST);
        expect(contrastRatio(palette.muted, background)).toBeGreaterThanOrEqual(MIN_TEXT_CONTRAST);
      });
    });
  });

  it('darkens only text, never the icon colour, for light custom colours', () => {
    const palette = getPalette('#fde047');
    expect(palette.accent).toBe('#fde047');
    expect(palette.text).not.toBe('#fde047');
    expect(contrastRatio(palette.text, palette.light)).toBeGreaterThanOrEqual(MIN_TEXT_CONTRAST);
  });

  it('falls back to the default accent for invalid input', () => {
    expect(getPalette('not a colour').accent).toBe(ACCENT_PRESETS[0].accent);
    expect(normalizeHex('#ABC')).toBe('#aabbcc');
  });
});
