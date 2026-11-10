import { DisplayMode, Version } from '@microsoft/sp-core-library';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  type IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import {
  type IFilePickerResult,
  PropertyFieldFilePicker
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import styles from './QuickLinkCardsWebPart.module.scss';
import * as strings from 'QuickLinkCardsWebPartStrings';

export interface IQuickLinkCard {
  id: string;
  title: string;
  description: string;
  buttonText: string;
  buttonUrl: string;
  imageUrl: string;
  openInNewTab: boolean;
  // When false the card uses the web part's default colors.
  customColors: boolean;
  accentColor?: string;
  buttonColor?: string;
  buttonTextColor?: string;
  cardBackground?: string;
  titleColor?: string;
}

// Which way the cards flow: from the left edge, centered, or from the right edge
// (first card on the far right).
export type CardAlignment = 'ltr' | 'center' | 'rtl';

export interface IQuickLinkCardsWebPartProps {
  cards: IQuickLinkCard[];
  columns: number;
  alignment: CardAlignment;
  imageWidth: number;
  accentColor: string;
  buttonColor: string;
  buttonTextColor: string;
  cardBackground: string;
  titleColor: string;
  descriptionColor: string;
  borderColor: string;
}

const DEFAULTS = {
  columns: 3,
  alignment: 'ltr' as CardAlignment,
  imageWidth: 120,
  accentColor: '#c8102e',
  buttonColor: '#181b23',
  buttonTextColor: '#ffffff',
  cardBackground: '#ffffff',
  titleColor: '#111111',
  descriptionColor: '#6d6d6d',
  borderColor: '#d6dbe1'
};

// A one-column SharePoint section is about 1200px wide, which fits four cards of the
// ~280px minimum the stylesheet allows; more per row could never actually show.
const MAX_COLUMNS = 4;
const IMAGE_WIDTH = { min: 60, max: 240 };

const FLOW: Record<CardAlignment, { direction: string; justify: string }> = {
  ltr: { direction: 'row', justify: 'flex-start' },
  center: { direction: 'row', justify: 'center' },
  // With row-reverse, "flex-start" is the right edge, so each row fills from the right.
  rtl: { direction: 'row-reverse', justify: 'flex-start' }
};

// Small inline icons for the edit-mode toolbar, so no icon font has to load.
const svg = (d: string): string =>
  `<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" focusable="false" fill="none" stroke="currentColor" stroke-width="1.25" stroke-linecap="round" stroke-linejoin="round"><path d="${d}"/></svg>`;
const ICONS = {
  add: svg('M8 3v10M3 8h10'),
  edit: svg('M10.5 2.5l3 3L6 13H3v-3zM9 4l3 3'),
  left: svg('M13 8H3M7 4L3 8l4 4'),
  right: svg('M3 8h10M9 4l4 4-4 4'),
  remove: svg('M2.5 4h11M6 4V2.5h4V4M4 4l.7 9.5h6.6L12 4M6.5 6.5v5M9.5 6.5v5'),
  image: svg('M2.5 3h11v10h-11zM2.5 11l3.5-3.5 3 3 2-2 2.5 2.5M10.5 5.5h.01')
};

// Folder (inside the site's "Site Assets" library) that images uploaded from the
// property pane are saved to.
const UPLOAD_FOLDER = 'QuickLinkCards';

// Property pane controls that only drive the UI (buttons, the image picker) still
// get a "target property"; these keys start with this prefix and are never saved.
const UI_ONLY_PREFIX = 'qlc';

const format = (template: string, ...values: (string | number)[]): string =>
  values.reduce<string>((text, value, i) => text.split(`{${i}}`).join(String(value)), template);

// How long after a card click a pane "start" still counts as this web part's own request.
const OWN_PANE_REQUEST_MS = 1500;

// Only one Quick Link Cards instance on a page edits a card at a time.
let activeInstance: QuickLinkCardsWebPart | undefined;

export default class QuickLinkCardsWebPart extends BaseClientSideWebPart<IQuickLinkCardsWebPartProps> {

  // Card being edited from its toolbar. While set, the property pane shows only that
  // card's settings; otherwise it shows the web part's layout and default colors.
  private _selectedCardId: string | undefined;
  // When this web part last asked for the pane itself (see onPropertyPaneConfigurationStart).
  private _paneRequestedAt: number | undefined;
  private _closeCheckTimer: number | undefined;
  private _focusLayoutPane: boolean = false;
  // Image upload status per card id, shown in that card's settings.
  private _uploadStatus: Map<string, string> = new Map();
  // Cards with an image upload still running, and the latest image choice per card:
  // an upload that finishes after a newer choice is ignored.
  private _uploading: Set<string> = new Set();
  private _imageChoice: Map<string, number> = new Map();
  private _imageChoiceCount: number = 0;
  // An upload that finished after the page left edit mode; applied when editing resumes.
  private _pendingImages: Map<string, string> = new Map();

  protected onInit(): Promise<void> {
    // Pages built by hand or by scripts may hold cards without ids, or copied cards that
    // share one; selection and image uploads need every id to be unique.
    this.properties.cards = (Array.isArray(this.properties.cards) ? this.properties.cards : [])
      .filter(card => !!card && typeof card === 'object');
    const seen = new Set<string>();
    for (const card of this.properties.cards) {
      if (typeof card.id !== 'string' || !card.id || seen.has(card.id)) {
        card.id = this._newId();
      }
      seen.add(card.id);
    }
    // Values saved by older versions (e.g. 6 per row, 'left') are shown as the
    // nearest current option, so the property pane always has one selected.
    this.properties.columns = this._columns();
    this.properties.imageWidth = this._imageWidth();
    this.properties.alignment = this._alignment();

    // Delegated listeners survive every re-render of the HTML below.
    this.domElement.addEventListener('click', event => this._onCanvasClick(event));
    this.domElement.addEventListener('keydown', event => this._onCanvasKeyDown(event));
    return super.onInit();
  }

  protected onDispose(): void {
    window.clearTimeout(this._closeCheckTimer);
    if (activeInstance === this) {
      activeInstance = undefined;
    }
    super.onDispose();
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    // An image whose upload finished after the page was published is saved the next
    // time the page is edited (and shown as saved only then).
    if (this.displayMode === DisplayMode.Edit && this._pendingImages.size) {
      this._pendingImages.forEach((url, cardId) => {
        const card = (this.properties.cards || []).find(c => c.id === cardId);
        if (card) {
          card.imageUrl = url;
        }
      });
      this._pendingImages.clear();
    }
    if (this.displayMode !== DisplayMode.Edit) {
      this._endCardEditing();
    }
    super.onDisplayModeChanged(oldDisplayMode);
  }

  // Section backgrounds can use a theme variant (for example "Strong", where the
  // background is the theme color), so edit-mode accents follow the section's colors.
  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    const colors = theme?.semanticColors;
    const set = (name: string, value: string | undefined): void => {
      if (value) {
        this.domElement.style.setProperty(name, value);
      } else {
        this.domElement.style.removeProperty(name);
      }
    };
    set('--qlc-theme-primary', colors?.link);
    set('--qlc-theme-neutral', colors?.inputBorder);
    set('--qlc-theme-hover', colors?.buttonBackgroundHovered);
    set('--qlc-theme-subtext', colors?.bodySubtext);
  }

  public render(): void {
    const cards = this.properties.cards || [];
    const isEditing = this.displayMode === DisplayMode.Edit;
    const focus = this._captureFocus();

    if (cards.length === 0) {
      this.domElement.innerHTML = isEditing
        ? `<div class="${styles.quickLinkCards}">
             <div class="${styles.placeholder}">
               <div>${escape(strings.PlaceholderText)}</div>
               <button type="button" data-qlc-action="add" data-qlc-key="add">${escape(strings.AddCardButton)}</button>
             </div>
           </div>`
        : '';
      this._restoreFocus(focus);
      return;
    }

    const p = this.properties;
    const flow = FLOW[this._alignment()];
    const rootVars = [
      `--qlc-columns:${this._columns()}`,
      `--qlc-direction:${flow.direction}`,
      `--qlc-align:${flow.justify}`,
      `--qlc-image-width:${this._imageWidth()}px`,
      `--qlc-description:${this._color(p.descriptionColor, DEFAULTS.descriptionColor)}`,
      `--qlc-border:${this._color(p.borderColor, DEFAULTS.borderColor)}`
    ].join(';');

    // The grid is always laid out left to right (see the stylesheet); each card keeps
    // the page's own text direction, so Hebrew or Arabic text still reads correctly.
    const textDir = getComputedStyle(this.domElement).direction === 'rtl' ? 'rtl' : 'ltr';

    const addBar = isEditing
      ? `<div class="${styles.addBar}">
           <button type="button" class="${styles.addButton}" data-qlc-action="add" data-qlc-key="add">${ICONS.add}<span>${escape(strings.AddCardButton)}</span></button>
         </div>`
      : '';

    this.domElement.innerHTML = `
      <div class="${styles.quickLinkCards} ${isEditing ? styles.editing : ''}" style="${escape(rootVars)}">
        ${addBar}
        <ul class="${styles.grid}" role="list">
          ${cards.map((card, i) => this._renderCard(card, i, cards.length, isEditing, textDir)).join('')}
        </ul>
      </div>`;
    this._restoreFocus(focus);
  }

  private _renderCard(card: IQuickLinkCard, index: number, count: number, isEditing: boolean, textDir: string): string {
    const p = this.properties;
    const c = card.customColors ? card : undefined;
    const cardVars = [
      `--qlc-accent:${this._color(c?.accentColor, p.accentColor, DEFAULTS.accentColor)}`,
      `--qlc-button-bg:${this._color(c?.buttonColor, p.buttonColor, DEFAULTS.buttonColor)}`,
      `--qlc-button-text:${this._color(c?.buttonTextColor, p.buttonTextColor, DEFAULTS.buttonTextColor)}`,
      `--qlc-card-bg:${this._color(c?.cardBackground, p.cardBackground, DEFAULTS.cardBackground)}`,
      `--qlc-title:${this._color(c?.titleColor, p.titleColor, DEFAULTS.titleColor)}`
    ].join(';');
    const title = card.title || '';
    const name = title || strings.UntitledCard;
    const heading = title ? `<h3 class="${styles.cardTitle}">${escape(title)}</h3>` : '';
    // Visible on the card itself, so authors know to wait before publishing.
    const uploading = isEditing && this._uploading.has(card.id)
      ? `<div class="${styles.uploadingNote}" role="status">${escape(strings.UploadingImageCard)}</div>`
      : '';

    let image = '';
    if (card.imageUrl) {
      // The title right next to it already says what the card is, so the image is decorative.
      image = `<div class="${styles.cardImage}">
                 <img src="${escape(this._safeUrl(card.imageUrl))}" alt="" loading="lazy" />
               </div>`;
    } else if (isEditing) {
      // Shows authors where the image goes; readers just see a text-only card.
      image = `<div class="${styles.cardImage} ${styles.imagePlaceholder}">${ICONS.image}<span>${escape(strings.AddImageHint)}</span></div>`;
    }
    const description = card.description
      ? `<div class="${styles.cardDescription}">${escape(card.description)}</div>`
      : '';

    let button = '';
    const url = this._safeUrl(card.buttonUrl);
    if (card.buttonText && url) {
      const target = card.openInNewTab ? ' target="_blank" rel="noopener noreferrer"' : '';
      // The hidden title tells screen-reader users which "Learn more" this is.
      const context = title ? `<span class="${styles.srOnly}">: ${escape(title)}</span>` : '';
      button = `<a class="${styles.cardButton}" href="${escape(url)}"${target} data-interception="off">${escape(card.buttonText)}${context}</a>`;
    } else if (card.buttonText && isEditing) {
      // Readers never see a button that goes nowhere; authors get a reminder instead.
      button = `<span class="${styles.cardButton} ${styles.cardButtonInactive}">${escape(card.buttonText)}</span>
                <span class="${styles.noLinkHint}">${escape(strings.NoLinkHint)}</span>`;
    }

    const isSelected = isEditing && this._selectedCardId !== undefined && card.id === this._selectedCardId;
    let editAttrs = '';
    let toolbar = '';
    if (isEditing) {
      // In edit mode the card itself is focusable and selectable (click, Enter or Space),
      // so keyboard users can reach cards that have no button.
      editAttrs = ` data-qlc-action="edit" data-qlc-index="${index}" data-qlc-key="card" tabindex="0"` +
        ` aria-label="${escape(format(strings.CardLabel, index + 1, count, name))}"` +
        (isSelected ? ' aria-current="true"' : '');
      toolbar = `<div class="${styles.toolbar}">
           ${this._toolButton('edit', 'edit', index, strings.EditCardButton, format(strings.EditCardLabel, name), ICONS.edit)}
           ${this._moveButton(index, count, -1, name)}
           ${this._moveButton(index, count, 1, name)}
           ${this._toolButton('delete', 'delete', index, strings.RemoveCardButton, format(strings.RemoveCardLabel, name), ICONS.remove)}
         </div>`;
    }

    return `
      <li class="${styles.card} ${isSelected ? styles.selected : ''}" style="${escape(cardVars)}" dir="${textDir}" data-qlc-card-id="${escape(card.id)}"${editAttrs}>
        ${toolbar}
        ${image}
        <div class="${styles.cardContent}">
          ${heading}
          <div class="${styles.accentLine}"></div>
          ${description}
          ${button}
          ${uploading}
        </div>
      </li>`;
  }

  private _toolButton(action: string, key: string, index: number, tooltip: string, label: string, icon: string, extraAttrs: string = ''): string {
    return `<button type="button" class="${styles.toolButton}" data-qlc-action="${action}" data-qlc-key="${key}" data-qlc-index="${index}"${extraAttrs} title="${escape(tooltip)}" aria-label="${escape(label)}">${icon}</button>`;
  }

  // Arrow buttons move a card visually left (-1) or right (+1). In right-to-left
  // layouts the first card is on the right, so "left" means later in the list.
  private _moveButton(index: number, count: number, visualStep: -1 | 1, name: string): string {
    const offset = this._alignment() === 'rtl' ? -visualStep : visualStep;
    const target = index + offset;
    if (target < 0 || target >= count) {
      return '';
    }
    const offsetAttr = ` data-qlc-offset="${offset}"`;
    return visualStep < 0
      ? this._toolButton('move', 'left', index, strings.MoveLeftButton, format(strings.MoveLeftLabel, name), ICONS.left, offsetAttr)
      : this._toolButton('move', 'right', index, strings.MoveRightButton, format(strings.MoveRightLabel, name), ICONS.right, offsetAttr);
  }

  // Remembers which card control had keyboard focus, because render() replaces the HTML.
  private _captureFocus(): { cardId?: string; key?: string; index: number } | undefined {
    const active = document.activeElement as HTMLElement | null;
    if (!active || active === this.domElement || !this.domElement.contains(active)) {
      return undefined;
    }
    const cardElement = active.closest<HTMLElement>('[data-qlc-card-id]');
    // Position on screen before the change, used if the card is about to disappear.
    const cardElements = Array.from(this.domElement.querySelectorAll('[data-qlc-card-id]'));
    return {
      cardId: cardElement?.dataset.qlcCardId,
      key: active.dataset.qlcKey,
      index: cardElement ? Math.max(0, cardElements.indexOf(cardElement)) : 0
    };
  }

  // Puts focus back on the same control (or the nearest sensible one) after a re-render,
  // only when focus was inside the web part, so typing in the property pane is never
  // interrupted.
  private _restoreFocus(focus: { cardId?: string; key?: string; index: number } | undefined): void {
    if (!focus) {
      return;
    }
    const cardElements = Array.from(this.domElement.querySelectorAll<HTMLElement>('[data-qlc-card-id]'));
    const cardElement = cardElements.find(element => element.dataset.qlcCardId === focus.cardId);
    let target: HTMLElement | null | undefined;
    if (cardElement) {
      target = (focus.key && focus.key !== 'card'
        ? cardElement.querySelector<HTMLElement>(`[data-qlc-key="${focus.key}"]`)
        : undefined) || cardElement;
    } else if (focus.cardId !== undefined && cardElements.length) {
      // The card was deleted: continue with the card that took its place.
      target = cardElements[Math.min(focus.index, cardElements.length - 1)];
    } else {
      target = this.domElement.querySelector<HTMLElement>(`[data-qlc-key="${focus.key || 'add'}"]`) ||
        this.domElement.querySelector<HTMLElement>('[data-qlc-key="add"]');
    }
    target?.focus();
  }

  // Older drafts stored 'left' / 'right'; anything unknown falls back to the default.
  private _alignment(): CardAlignment {
    const value = this.properties.alignment as string;
    if (value === 'left') {
      return 'ltr';
    }
    if (value === 'right') {
      return 'rtl';
    }
    return Object.prototype.hasOwnProperty.call(FLOW, value) ? value as CardAlignment : DEFAULTS.alignment;
  }

  // Numbers that end up in inline styles are clamped, so stored values can't inject markup.
  private _columns(): number {
    return this._clamp(this.properties.columns, 1, MAX_COLUMNS, DEFAULTS.columns);
  }

  private _imageWidth(): number {
    return this._clamp(this.properties.imageWidth, IMAGE_WIDTH.min, IMAGE_WIDTH.max, DEFAULTS.imageWidth);
  }

  private _clamp(value: unknown, min: number, max: number, fallback: number): number {
    const number = Math.round(Number(value));
    return Number.isFinite(number) && number > 0 ? Math.min(max, Math.max(min, number)) : fallback;
  }

  private _onCanvasClick(event: MouseEvent): void {
    if (this.displayMode !== DisplayMode.Edit) {
      return;
    }
    const target = (event.target as HTMLElement).closest<HTMLElement>('[data-qlc-action]');
    if (!target || !this.domElement.contains(target)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    this._runAction(target);
  }

  private _onCanvasKeyDown(event: KeyboardEvent): void {
    const target = event.target as HTMLElement;
    // Only the focused card itself; buttons and links inside keep their own keys.
    if (this.displayMode !== DisplayMode.Edit || target.dataset.qlcKey !== 'card' ||
      (event.key !== 'Enter' && event.key !== ' ')) {
      return;
    }
    event.preventDefault();
    this._runAction(target);
  }

  private _runAction(target: HTMLElement): void {
    const index = Number(target.dataset.qlcIndex);
    switch (target.dataset.qlcAction) {
      case 'add':
        this._addCard();
        break;
      case 'edit':
        this._selectCard(this.properties.cards[index]?.id);
        break;
      case 'move':
        this._moveCard(index, Number(target.dataset.qlcOffset));
        break;
      case 'delete':
        this._removeCard(index);
        break;
    }
  }

  // Returns the first value that looks like a CSS color, so a typo in the color
  // box can't break the page's styles.
  private _color(...candidates: (string | undefined)[]): string {
    for (const value of candidates) {
      if (typeof value === 'string' && /^[#a-zA-Z0-9(),.\s%]+$/.test(value)) {
        return value.trim();
      }
    }
    return '';
  }

  // Only allow web, relative, mail and phone links (blocks "javascript:" links).
  private _safeUrl(url: string | undefined): string {
    const value = (typeof url === 'string' ? url : '').trim();
    if (!value) {
      return '';
    }
    return /^(https?:|mailto:|tel:|\/|#|\?)/i.test(value) ? value : `https://${value}`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _newId(): string {
    return `card-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
  }

  private _newCard(): IQuickLinkCard {
    return {
      id: this._newId(),
      title: strings.NewCardTitle,
      description: strings.NewCardDescription,
      buttonText: strings.NewCardButtonText,
      buttonUrl: '',
      imageUrl: '',
      openInNewTab: false,
      customColors: false
    };
  }

  private _addCard(): void {
    const card = this._newCard();
    this.properties.cards = [...(this.properties.cards || []), card];
    this._selectCard(card.id);
  }

  // Opens the property pane on one card's settings, like editing a Quick Links item.
  // open() (not refresh()) also takes the pane over when it shows a different web part
  // or section. It runs on the next tick so it is the last pane request for this click,
  // after anything SharePoint's canvas does for the same click.
  private _selectCard(id: string | undefined): void {
    if (activeInstance && activeInstance !== this) {
      activeInstance._endCardEditing();
    }
    activeInstance = this;
    window.clearTimeout(this._closeCheckTimer);
    this._selectedCardId = id;
    this.render();
    this._paneRequestedAt = Date.now();
    window.setTimeout(() => this.context.propertyPane.open(), 0);
  }

  private _showLayoutSettings(): void {
    this._endCardEditing();
    // The button that was pressed disappears with the card settings; move keyboard
    // focus to the first layout setting instead of losing it.
    this._focusLayoutPane = true;
    this.context.propertyPane.refresh();
  }

  // Back to "no card selected". refresh() only redraws the pane if it is showing this
  // web part, so it is safe to call whatever the pane is doing.
  private _endCardEditing(): void {
    window.clearTimeout(this._closeCheckTimer);
    if (activeInstance === this) {
      activeInstance = undefined;
    }
    if (this._selectedCardId === undefined) {
      return;
    }
    this._selectedCardId = undefined;
    for (const cardId of Array.from(this._uploadStatus.keys())) {
      if (!this._uploading.has(cardId)) {
        this._uploadStatus.delete(cardId);
      }
    }
    this.render();
    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneConfigurationStart(): void {
    // The pane started showing this web part. If this web part didn't ask for it (the
    // author used the web part's own edit button, or selected the web part), show the
    // layout settings, as SharePoint's own Quick Links does.
    // Anything within a moment of a card click counts as that click: SharePoint's canvas
    // may send its own pane request for the same click, in either order.
    if (this._paneRequestedAt !== undefined && Date.now() - this._paneRequestedAt < OWN_PANE_REQUEST_MS) {
      return;
    }
    this._endCardEditing();
  }

  protected onPropertyPaneConfigurationComplete(): void {
    // SharePoint also raises this 5 seconds after the last change, when the author
    // presses Alt+P, and when open() re-targets a pane that is already open. Card
    // editing ends only once the pane has settled and is closed or belongs to
    // something else (another web part, a section, the page).
    window.clearTimeout(this._closeCheckTimer);
    this._closeCheckTimer = window.setTimeout(() => {
      this._closeCheckTimer = undefined;
      const pane = this.context.propertyPane;
      if (this._selectedCardId !== undefined && (!pane.isPropertyPaneOpen() || !pane.isRenderedByWebPart())) {
        this._endCardEditing();
      }
    }, 500);
  }

  private _moveCard(index: number, offset: number): void {
    const cards = [...this.properties.cards];
    const target = index + offset;
    if (target < 0 || target >= cards.length) {
      return;
    }
    [cards[index], cards[target]] = [cards[target], cards[index]];
    this.properties.cards = cards;
    this._refresh();
  }

  private _removeCard(index: number): void {
    const card = this.properties.cards[index];
    if (!card || !confirm(format(strings.RemoveCardConfirm, card.title || strings.UntitledCard))) {
      return;
    }
    this.properties.cards = this.properties.cards.filter((_, i) => i !== index);
    if (card.id === this._selectedCardId) {
      this._endCardEditing();
    }
    this._refresh();
  }

  private _refresh(): void {
    this.context.propertyPane.refresh();
    this.render();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    // SharePoint writes a value for every control, including buttons and the image
    // picker; only the real settings are kept.
    if (propertyPath.indexOf(UI_ONLY_PREFIX) === 0) {
      delete (this.properties as unknown as Record<string, unknown>)[propertyPath];
    }

    const match = /^cards\[(\d+)\]\.(\w+)$/.exec(propertyPath);
    if (match) {
      const card = this.properties.cards[Number(match[1])];
      // Turning on custom colors starts from the current defaults, so nothing jumps.
      if (card && match[2] === 'customColors' && newValue === true) {
        card.accentColor = card.accentColor || this.properties.accentColor || DEFAULTS.accentColor;
        card.buttonColor = card.buttonColor || this.properties.buttonColor || DEFAULTS.buttonColor;
        card.buttonTextColor = card.buttonTextColor || this.properties.buttonTextColor || DEFAULTS.buttonTextColor;
        card.cardBackground = card.cardBackground || this.properties.cardBackground || DEFAULTS.cardBackground;
        card.titleColor = card.titleColor || this.properties.titleColor || DEFAULTS.titleColor;
      }
      // The color fields appear or disappear with this toggle.
      if (match[2] === 'customColors') {
        this.context.propertyPane.refresh();
      }
      // Typing or clearing the image address is the author's newest image choice.
      if (card && match[2] === 'imageUrl') {
        this._imageChoice.set(card.id, ++this._imageChoiceCount);
        this._uploading.delete(card.id);
        this._uploadStatus.delete(card.id);
      }
    }
    // Re-render so changes made above (e.g. filled-in custom colors) show immediately.
    this.render();
  }

  protected onBeforeSerialize(): void {
    // Last line of defence: never save UI-only values such as raw image-picker results,
    // which can include the author's OneDrive paths.
    for (const key of Object.keys(this.properties)) {
      if (key.indexOf(UI_ONLY_PREFIX) === 0) {
        delete (this.properties as unknown as Record<string, unknown>)[key];
      }
    }
    super.onBeforeSerialize();
  }

  // Images from this site, stock images and pasted links are used as they are. Images
  // uploaded from the computer, and files browsed from OneDrive or another site, are
  // copied into this site's Site Assets library so every reader can see them.
  private async _onImagePicked(cardId: string, result: IFilePickerResult): Promise<void> {
    const findCard = (): IQuickLinkCard | undefined => (this.properties.cards || []).find(card => card.id === cardId);
    if (!findCard() || !result) {
      return;
    }
    const choice = ++this._imageChoiceCount;
    this._imageChoice.set(cardId, choice);
    const isLatest = (): boolean => this._imageChoice.get(cardId) === choice;
    try {
      let url = result.fileAbsoluteUrl;
      if (!url || this._isFileOutsideThisSite(result)) {
        this._uploading.add(cardId);
        this._uploadStatus.set(cardId, strings.UploadingImage);
        this._refresh();
        const file = await result.downloadFileContent();
        if (!file) {
          throw new Error(strings.ImageReadFailed);
        }
        url = await this._uploadToSiteAssets(file);
      } else if (Object.prototype.hasOwnProperty.call(result, 'spItemUrl')) {
        url = this._encodePath(url);
      }
      if (!isLatest()) {
        return;
      }
      this._uploading.delete(cardId);
      this._uploadStatus.delete(cardId);
      if (this.displayMode !== DisplayMode.Edit) {
        // The page was published while the upload ran: those properties are already
        // saved, so keep the image for the next edit instead of pretending it's saved.
        this._pendingImages.set(cardId, url);
        this.render();
        return;
      }
      // Look the card up again: it may have moved while the upload was running.
      const card = findCard();
      if (card) {
        card.imageUrl = url;
      }
    } catch (error) {
      if (!isLatest()) {
        return;
      }
      this._uploading.delete(cardId);
      this._uploadStatus.set(cardId, `${strings.UploadFailed} ${this._errorText(error)}`);
    }
    this._refresh();
  }

  // Files browsed in this site come back with a plain (unencoded) path, so names with
  // spaces, '#' or '%' must be encoded to work as an image address.
  private _encodePath(url: string): string {
    const match = /^(https?:\/\/[^/]+)(\/.*)?$/i.exec(url);
    if (!match || !match[2]) {
      return url;
    }
    return match[1] + match[2].split('/').map(segment => encodeURIComponent(segment)).join('/');
  }

  // Files chosen by browsing (the OneDrive and site file tabs) carry spItemUrl.
  private _isFileOutsideThisSite(result: IFilePickerResult): boolean {
    if (!Object.prototype.hasOwnProperty.call(result, 'spItemUrl') || !result.fileAbsoluteUrl) {
      return false;
    }
    const site = this.context.pageContext.site.absoluteUrl.toLowerCase().replace(/\/$/, '');
    const url = result.fileAbsoluteUrl.toLowerCase();
    return url !== site && url.indexOf(`${site}/`) !== 0;
  }

  // Turns a SharePoint REST error body into the plain sentence inside it.
  private _errorText(error: unknown): string {
    const message = error instanceof Error ? error.message : String(error);
    try {
      const body = JSON.parse(message);
      const text = body?.error?.message ?? body?.['odata.error']?.message?.value;
      if (typeof text === 'string' && text) {
        return text;
      }
    } catch {
      // Not JSON; use the message as it is.
    }
    return message.length > 200 ? `${message.slice(0, 200)}...` : message;
  }

  private async _uploadToSiteAssets(file: File): Promise<string> {
    const web = this.context.pageContext.web.absoluteUrl;
    const client = this.context.spHttpClient;
    const config = SPHttpClient.configurations.v1;

    // Creates the Site Assets library if this site doesn't have one yet.
    const listResponse = await client.post(
      `${web}/_api/web/lists/EnsureSiteAssetsLibrary()?$select=Id`, config, {});
    if (!listResponse.ok) {
      throw new Error(await listResponse.text());
    }
    const list = await listResponse.json();
    const folderResponse = await client.get(
      `${web}/_api/web/lists(guid'${list.Id}')/RootFolder?$select=ServerRelativeUrl`, config);
    if (!folderResponse.ok) {
      throw new Error(await folderResponse.text());
    }
    const rootFolder = await folderResponse.json();
    const folderUrl = `${rootFolder.ServerRelativeUrl}/${UPLOAD_FOLDER}`;

    // Make sure the upload folder exists (no-op if it already does).
    await client.post(`${web}/_api/web/folders/addUsingPath(decodedurl='${this._odata(folderUrl)}')`, config, {});

    const fileName = `${Date.now()}-${file.name.replace(/[^\w.-]+/g, '_')}`;
    const uploadResponse = await client.post(
      `${web}/_api/web/GetFolderByServerRelativePath(decodedurl='${this._odata(folderUrl)}')` +
      `/Files/AddUsingPath(decodedurl='${this._odata(fileName)}',overwrite=true)?$select=ServerRelativeUrl`,
      config,
      { body: file });
    if (!uploadResponse.ok) {
      throw new Error(await uploadResponse.text());
    }
    const uploaded = await uploadResponse.json();
    return `${new URL(web).origin}${uploaded.ServerRelativeUrl}`;
  }

  // Escapes a value for use inside a quoted OData string in a REST URL.
  private _odata(value: string): string {
    return encodeURIComponent(value.replace(/'/g, "''"));
  }

  private _colorField(path: string, label: string, color: string | undefined, key: string): IPropertyPaneField<unknown> {
    return PropertyFieldColorPicker(path, {
      label,
      selectedColor: color,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      alphaSliderHidden: true,
      style: PropertyFieldColorPickerStyle.Inline,
      iconName: 'Precipitation',
      key
    });
  }

  private _cardGroups(card: IQuickLinkCard, index: number): IPropertyPaneGroup[] {
    const path = `cards[${index}]`;
    const imageFields: IPropertyPaneField<unknown>[] = [
      PropertyFieldFilePicker(`${UI_ONLY_PREFIX}ImagePicker${card.id}`, {
        context: this.context as never,
        label: card.imageUrl ? strings.ImageChangeLabel : strings.ImageLabel,
        buttonLabel: card.imageUrl ? strings.ImageChangeButton : strings.ImageButton,
        buttonIcon: 'Photo2',
        accepts: ['.gif', '.jpg', '.jpeg', '.png', '.svg', '.webp'],
        filePickerResult: undefined as unknown as IFilePickerResult,
        properties: this.properties,
        // The picker stores its raw result under its own key; only imageUrl is kept.
        onPropertyChange: (propertyPath: string) => {
          delete (this.properties as unknown as Record<string, unknown>)[propertyPath];
        },
        onSave: (result: IFilePickerResult) => { this._onImagePicked(card.id, result).catch(() => undefined); },
        hideWebSearchTab: true,
        key: `imagePicker${card.id}`
      }),
      PropertyPaneTextField(`${path}.imageUrl`, {
        label: strings.ImageUrlLabel,
        description: strings.ImageUrlDescription
      })
    ];
    const status = this._uploadStatus.get(card.id);
    if (status) {
      imageFields.push(PropertyPaneLabel(`${UI_ONLY_PREFIX}UploadMessage`, { text: status }));
    }

    const colorFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle(`${path}.customColors`, {
        label: strings.CustomColorsLabel,
        onText: strings.CustomColorsOn,
        offText: strings.CustomColorsOff
      })
    ];
    if (card.customColors) {
      colorFields.push(
        this._colorField(`${path}.accentColor`, strings.AccentColorLabel, card.accentColor, `accent${card.id}`),
        this._colorField(`${path}.titleColor`, strings.TitleColorLabel, card.titleColor, `title${card.id}`),
        this._colorField(`${path}.buttonColor`, strings.ButtonColorLabel, card.buttonColor, `button${card.id}`),
        this._colorField(`${path}.buttonTextColor`, strings.ButtonTextColorLabel, card.buttonTextColor, `buttonText${card.id}`),
        this._colorField(`${path}.cardBackground`, strings.CardBackgroundLabel, card.cardBackground, `bg${card.id}`)
      );
    }

    return [
      {
        groupName: strings.TextAndLinkGroup,
        groupFields: [
          PropertyPaneTextField(`${path}.title`, { label: strings.TitleLabel }),
          PropertyPaneTextField(`${path}.description`, { label: strings.DescriptionLabel, multiline: true, rows: 2 }),
          PropertyPaneTextField(`${path}.buttonText`, { label: strings.ButtonTextLabel }),
          PropertyPaneTextField(`${path}.buttonUrl`, {
            label: strings.ButtonUrlLabel,
            placeholder: 'https://',
            description: strings.ButtonUrlDescription
          }),
          PropertyPaneToggle(`${path}.openInNewTab`, { label: strings.OpenInNewTabLabel })
        ]
      },
      { groupName: strings.ImageGroup, groupFields: imageFields },
      { groupName: strings.CardColorsGroup, groupFields: colorFields },
      {
        groupName: strings.MoreSettingsGroup,
        groupFields: [
          PropertyPaneButton(`${UI_ONLY_PREFIX}BackToLayout`, {
            text: strings.BackToLayoutButton,
            icon: 'Back',
            buttonType: PropertyPaneButtonType.Command,
            onClick: () => this._showLayoutSettings()
          })
        ]
      }
    ];
  }

  private _focusIfRequested(field: IPropertyPaneField<unknown>): IPropertyPaneField<unknown> {
    if (this._focusLayoutPane) {
      this._focusLayoutPane = false;
      field.shouldFocus = true;
    }
    return field;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const cards = this.properties.cards || [];
    const p = this.properties;

    const selectedIndex = this._selectedCardId === undefined
      ? -1
      : cards.findIndex(card => card.id === this._selectedCardId);
    if (selectedIndex !== -1) {
      return {
        pages: [
          {
            header: { description: format(strings.CardPaneDescription, selectedIndex + 1, cards.length) },
            groups: this._cardGroups(cards[selectedIndex], selectedIndex)
          }
        ]
      };
    }

    return {
      pages: [
        {
          header: { description: strings.StylePageDescription },
          groups: [
            {
              groupName: strings.LayoutGroup,
              groupFields: [
                PropertyPaneLabel(`${UI_ONLY_PREFIX}Help`, { text: strings.EditCardsHelp }),
                this._focusIfRequested(PropertyPaneDropdown('columns', {
                  label: strings.ColumnsLabel,
                  options: Array.from({ length: MAX_COLUMNS }, (_, i) => ({ key: i + 1, text: String(i + 1) }))
                })),
                PropertyPaneLabel(`${UI_ONLY_PREFIX}ColumnsHelp`, { text: strings.ColumnsHelp }),
                PropertyPaneChoiceGroup('alignment', {
                  label: strings.AlignmentLabel,
                  options: [
                    { key: 'ltr', text: strings.AlignLtr, iconProps: { officeFabricIconFontName: 'AlignLeft' } },
                    { key: 'center', text: strings.AlignCenter, iconProps: { officeFabricIconFontName: 'AlignCenter' } },
                    { key: 'rtl', text: strings.AlignRtl, iconProps: { officeFabricIconFontName: 'AlignRight' } }
                  ]
                }),
                PropertyPaneSlider('imageWidth', {
                  label: strings.ImageWidthLabel,
                  min: IMAGE_WIDTH.min,
                  max: IMAGE_WIDTH.max,
                  step: 10
                })
              ]
            },
            {
              groupName: strings.DefaultColorsGroup,
              groupFields: [
                PropertyPaneLabel(`${UI_ONLY_PREFIX}ColorsHelp`, { text: strings.DefaultColorsHelp }),
                this._colorField('accentColor', strings.AccentColorLabel, p.accentColor, 'accentColor'),
                this._colorField('titleColor', strings.TitleColorLabel, p.titleColor, 'titleColor'),
                this._colorField('descriptionColor', strings.DescriptionColorLabel, p.descriptionColor, 'descriptionColor'),
                this._colorField('buttonColor', strings.ButtonColorLabel, p.buttonColor, 'buttonColor'),
                this._colorField('buttonTextColor', strings.ButtonTextColorLabel, p.buttonTextColor, 'buttonTextColor'),
                this._colorField('cardBackground', strings.CardBackgroundLabel, p.cardBackground, 'cardBackground'),
                this._colorField('borderColor', strings.BorderColorLabel, p.borderColor, 'borderColor')
              ]
            }
          ]
        }
      ]
    };
  }
}
