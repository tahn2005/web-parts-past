import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import styles from './PropertyPaneControls.module.scss';
import { getIcon, searchIcons } from '../common/icons';

export interface IIconPickerProps {
  label: string;
  value: string;
  onChange: (iconName: string) => void;
}

let instanceCount = 0;

/**
 * Searchable icon grid. Built from native radio buttons so it works with the
 * keyboard (Tab into the group, arrow keys to move, selection follows) and is
 * announced as "radio button, 3 of 92" by screen readers with no extra ARIA.
 */
export default function IconPicker(props: IIconPickerProps): React.ReactElement {
  const [query, setQuery] = React.useState('');
  const [groupName] = React.useState(() => `rc-icon-picker-${++instanceCount}`);
  const selected = getIcon(props.value);
  const results = searchIcons(query);
  const SelectedIcon = selected.Icon;

  return (
    <fieldset className={styles.fieldset}>
      <legend className={styles.legend}>{props.label}</legend>

      <div className={styles.selectedIcon}>
        <span className={styles.selectedIconTile}>
          <SelectedIcon aria-hidden="true" focusable="false" />
        </span>
        <span>
          Selected: <strong>{selected.label}</strong>
        </span>
      </div>

      <TextField
        label="Search icons"
        value={query}
        onChange={(_event, newValue) => setQuery(newValue || '')}
        placeholder="For example: calendar, people, money"
        autoComplete="off"
      />

      <p className={styles.resultCount} role="status" aria-live="polite">
        {results.length === 0 ? 'No icons match your search.' : `${results.length} icons`}
      </p>

      {results.length > 0 && (
        <div className={styles.iconGrid}>
          {results.map((option) => {
            const Icon = option.Icon;
            const isSelected = option.name === selected.name;
            return (
              <label key={option.name} className={styles.iconOption} title={option.label}>
                <input
                  type="radio"
                  className={styles.visuallyHiddenInput}
                  name={groupName}
                  value={option.name}
                  checked={isSelected}
                  onChange={() => props.onChange(option.name)}
                />
                <span className={`${styles.iconSwatch} ${isSelected ? styles.iconSwatchSelected : ''}`}>
                  <Icon aria-hidden="true" focusable="false" />
                </span>
                <span className={styles.srOnly}>{option.label}</span>
              </label>
            );
          })}
        </div>
      )}
    </fieldset>
  );
}
