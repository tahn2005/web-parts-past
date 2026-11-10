import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Label } from '@fluentui/react/lib/Label';
import { getIcon, ICON_NAMES, iconLabel } from '../icons';
import styles from './IconPicker.module.scss';

export interface IIconPickerProps {
  label: string;
  value: string;
  onChange: (iconName: string) => void;
}

/**
 * A searchable grid of the icons the web part ships with. Each icon is a
 * toggle button with its name as the accessible label, so the picker works
 * the same with a keyboard or screen reader as with a mouse.
 */
export function IconPicker({ label, value, onChange }: IIconPickerProps): JSX.Element {
  const [query, setQuery] = React.useState('');
  const [selected, setSelected] = React.useState(value);
  const labelId = React.useMemo(() => `icon-picker-${Math.random().toString(36).slice(2)}`, []);

  React.useEffect(() => setSelected(value), [value]);

  const needle = query.trim().toLowerCase();
  const visible = ICON_NAMES.filter(
    (name) => !needle || name.toLowerCase().includes(needle) || iconLabel(name).toLowerCase().includes(needle)
  );
  const Current = getIcon(selected);

  const pick = (name: string): void => {
    setSelected(name);
    onChange(name);
  };

  return (
    <div className={styles.picker}>
      <Label id={labelId}>{label}</Label>
      <div className={styles.current}>
        <span className={styles.currentIcon} aria-hidden="true">
          <Current focusable="false" />
        </span>
        <span>{iconLabel(selected || 'FileText')}</span>
      </div>
      <SearchBox
        placeholder="Search icons"
        ariaLabel="Search icons"
        value={query}
        onChange={(_, next) => setQuery(next ?? '')}
        onClear={() => setQuery('')}
      />
      <div className={styles.grid} role="group" aria-labelledby={labelId}>
        {visible.map((name) => {
          const Icon = getIcon(name);
          const isSelected = name === selected;
          return (
            <button
              key={name}
              type="button"
              className={[styles.option, isSelected ? styles.selected : ''].join(' ')}
              title={iconLabel(name)}
              aria-label={iconLabel(name)}
              aria-pressed={isSelected}
              onClick={() => pick(name)}
            >
              <Icon aria-hidden="true" focusable="false" />
            </button>
          );
        })}
      </div>
      <div className={styles.status} role="status">
        {visible.length === 0 ? 'No icons match your search.' : ''}
      </div>
    </div>
  );
}
