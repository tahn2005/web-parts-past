import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { ArrowRight, FileText } from 'lucide-react';
import styles from './PropertyPaneControls.module.scss';
import { ACCENT_PRESETS, CARD_BACKGROUND, getPalette, normalizeHex } from '../common/color';

export interface IAccentColorPickerProps {
  label: string;
  /** "#rrggbb", or '' to inherit when `inheritLabel` is set. */
  value: string;
  onChange: (value: string) => void;
  /** When set, offers an extra first choice that stores '' (inherit). */
  inheritLabel?: string;
  /** The colour inherited when value is '' (used for the preview). */
  inheritedColor?: string;
}

let instanceCount = 0;

/**
 * Accent colour chooser: the original lifecycle page palettes as one-click
 * swatches, plus any custom colour by hex code or the system colour picker.
 *
 * A live preview shows the link as it will appear. If the chosen colour is
 * too light for readable text, the preview (and the card) use a darkened
 * shade for the link, and the pane says so.
 */
export default function AccentColorPicker(props: IAccentColorPickerProps): React.ReactElement {
  const [groupName] = React.useState(() => `rc-accent-${++instanceCount}`);
  const value = normalizeHex(props.value);
  const isInherit = !!props.inheritLabel && !value;
  const isPreset = ACCENT_PRESETS.some((preset) => preset.accent === value);
  const effective = value || normalizeHex(props.inheritedColor) || ACCENT_PRESETS[0].accent;
  const palette = getPalette(effective);

  const [customText, setCustomText] = React.useState(value);
  React.useEffect(() => {
    setCustomText(value);
  }, [value]);

  const customError = customText && !normalizeHex(customText) ? 'Enter a hex colour such as #2563eb.' : '';

  const commitCustom = (text: string): void => {
    setCustomText(text);
    const hex = normalizeHex(text);
    if (hex) {
      props.onChange(hex);
    }
  };

  return (
    <fieldset className={styles.fieldset}>
      <legend className={styles.legend}>{props.label}</legend>

      <div className={styles.swatchRow}>
        {props.inheritLabel && (
          <label className={styles.swatchOption} title={props.inheritLabel}>
            <input
              type="radio"
              className={styles.visuallyHiddenInput}
              name={groupName}
              checked={isInherit}
              onChange={() => props.onChange('')}
            />
            <span
              className={`${styles.swatch} ${styles.swatchInherit} ${isInherit ? styles.swatchSelected : ''}`}
              style={{ backgroundColor: normalizeHex(props.inheritedColor) || ACCENT_PRESETS[0].accent }}
              aria-hidden="true"
            />
            <span className={styles.srOnly}>{props.inheritLabel}</span>
          </label>
        )}
        {ACCENT_PRESETS.map((preset) => {
          const checked = value === preset.accent;
          return (
            <label key={preset.accent} className={styles.swatchOption} title={preset.name}>
              <input
                type="radio"
                className={styles.visuallyHiddenInput}
                name={groupName}
                checked={checked}
                onChange={() => props.onChange(preset.accent)}
              />
              <span
                className={`${styles.swatch} ${checked ? styles.swatchSelected : ''}`}
                style={{ backgroundColor: preset.accent }}
                aria-hidden="true"
              />
              <span className={styles.srOnly}>{preset.name}</span>
            </label>
          );
        })}
      </div>
      <p className={styles.helpText}>
        {isInherit
          ? props.inheritLabel
          : isPreset
            ? ACCENT_PRESETS.filter((preset) => preset.accent === value)[0].name
            : `Custom colour ${value}`}
      </p>

      <div className={styles.customRow}>
        <div className={styles.customText}>
          <TextField
            label="Custom colour (hex)"
            value={customText}
            placeholder={effective}
            onChange={(_event, newValue) => commitCustom(newValue || '')}
            errorMessage={customError}
            autoComplete="off"
            spellCheck={false}
          />
        </div>
        <label className={styles.nativeColor}>
          <span className={styles.nativeColorLabel}>Pick colour</span>
          <input type="color" value={effective} onChange={(event) => commitCustom(event.target.value)} />
        </label>
      </div>

      <div className={styles.preview} aria-hidden="true">
        <span className={styles.previewLabel}>Preview</span>
        <span className={styles.previewCard} style={{ backgroundColor: CARD_BACKGROUND }}>
          <span className={styles.previewTile} style={{ backgroundColor: palette.light, color: palette.accent }}>
            <FileText focusable="false" />
          </span>
          <span className={styles.previewLink} style={{ color: palette.text }}>
            View Resource
            <ArrowRight focusable="false" />
          </span>
        </span>
      </div>
      {palette.text !== palette.accent && (
        <p className={styles.helpText}>
          This colour is too light for readable text, so card links use a slightly darker shade ({palette.text}). Icons keep
          the exact colour.
        </p>
      )}
    </fieldset>
  );
}
