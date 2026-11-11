import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ResourceCardsWebPartStrings';
import ResourceCards, { type IResourceCardsProps } from './components/ResourceCards';
import IconPicker from './propertyPane/IconPicker';
import AccentColorPicker from './propertyPane/AccentColorPicker';
import ImagePicker from './propertyPane/ImagePicker';
import { reactPropertyPaneField } from './propertyPane/reactField';
import {
  createCard,
  normalizeCards,
  type HeadingLevel,
  type IResourceCard,
  type IResourceCardsWebPartProps
} from './common/model';
import { DEFAULT_ACCENT, normalizeHex } from './common/color';
import { checkUrl } from './common/url';
import { uploadImageToSiteAssets } from './common/upload';
import { registerCardFont } from './common/fonts';

/** How long after a "configuration complete" event the pane is checked (see onPropertyPaneConfigurationComplete). */
const PANE_SETTLE_MS = 500;
/** How long after this web part asks for the pane a pane "start" still counts as its own request. */
const OWN_PANE_REQUEST_MS = 1500;
/** A "start" this soon after a "complete" means the pane never left this web part. */
const PANE_RETARGET_MS = 100;
/** Raised on the document when a web part puts one of its cards in the pane, so any other one stops editing its card. */
const CARD_PANE_CLAIMED = 'card-settings-pane-claimed';

export default class ResourceCardsWebPart extends BaseClientSideWebPart<IResourceCardsWebPartProps> {
  private _root: Root | undefined;
  /** Card whose settings are showing in the property pane, if any. */
  private _selectedCardId: string | undefined;
  private _closeCheckTimer: number | undefined;
  private _openTimer: number | undefined;
  private _verifyTimer: number | undefined;
  /** When this web part last asked for the pane itself, and when the pane last reported "complete". */
  private _paneRequestedAt: number | undefined;
  private _paneCompletedAt: number | undefined;

  protected onInit(): Promise<void> {
    this._applyDefaults();
    registerCardFont();
    document.addEventListener(CARD_PANE_CLAIMED, this._onPaneClaimed);
    return super.onInit();
  }

  public render(): void {
    const isEditMode = this.displayMode === DisplayMode.Edit;

    const element = React.createElement<IResourceCardsProps>(ResourceCards, {
      cards: this.properties.cards,
      accentColor: this.properties.accentColor,
      headingLevel: this.properties.headingLevel,
      useInterFont: this.properties.useInterFont,
      isEditMode,
      selectedCardId: isEditMode ? this._selectedCardId : undefined,
      onAddCard: this._addCard,
      onEditCard: this._editCard,
      onMoveCard: this._moveCard,
      onDeleteCard: this._deleteCard
    });

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(element);
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    window.clearTimeout(this._verifyTimer);
    this._selectedCardId = undefined;
    super.onDisplayModeChanged(oldDisplayMode);
    this.render();
  }

  protected onDispose(): void {
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    window.clearTimeout(this._verifyTimer);
    document.removeEventListener(CARD_PANE_CLAIMED, this._onPaneClaimed);
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // ---- Card operations (shared by the canvas toolbar and the property pane) --

  private _addCard = (): string => {
    const card = createCard();
    this.properties.cards = [...this.properties.cards, card];
    this._selectedCardId = card.id;
    this._afterCardsChanged(true);
    return card.id;
  };

  private _editCard = (id: string): void => {
    this._selectedCardId = id;
    this._afterCardsChanged(true);
  };

  private _moveCard = (id: string, delta: -1 | 1): void => {
    const cards = [...this.properties.cards];
    const from = this._indexOf(id);
    const to = from + delta;
    if (from === -1 || to < 0 || to >= cards.length) {
      return;
    }
    const [card] = cards.splice(from, 1);
    cards.splice(to, 0, card);
    this.properties.cards = cards;
    this._afterCardsChanged(false);
  };

  private _deleteCard = (id: string): void => {
    this.properties.cards = this.properties.cards.filter((card) => card.id !== id);
    if (this._selectedCardId === id) {
      this._selectedCardId = undefined;
    }
    this._afterCardsChanged(false);
  };

  private _showAllSettings = (): void => {
    this._selectedCardId = undefined;
    this._afterCardsChanged(false);
  };

  private _indexOf(id: string | undefined): number {
    for (let i = 0; i < this.properties.cards.length; i++) {
      if (this.properties.cards[i].id === id) {
        return i;
      }
    }
    return -1;
  }

  /** Re-renders the canvas and keeps the property pane in step with it. */
  private _afterCardsChanged(openPane: boolean): void {
    window.clearTimeout(this._closeCheckTimer);
    this.render();
    this.context.propertyPane.refresh();
    if (openPane) {
      this._claimPane();
    }
  }

  /**
   * A property pane button writes its onClick result to its target property.
   * Each one here acts through the methods above (which update `cards`
   * themselves), then hands back the up-to-date list so nothing is
   * overwritten. The pane is refreshed once SharePoint has applied the write.
   */
  private _cardsButton(action: () => void): () => IResourceCard[] {
    return () => {
      action();
      window.setTimeout(() => this.context.propertyPane.refresh(), 0);
      return this.properties.cards;
    };
  }

  // ---- Property pane ---------------------------------------------------------

  /**
   * Puts the pane on the selected card and keeps it there. open() (not just
   * refresh()) also takes the pane over when it shows something else, and
   * marks the pane as opened by a web part, which is how the check in
   * onPropertyPaneConfigurationComplete knows it still belongs here. It runs
   * on the next tick so it is the last pane request for the click that
   * selected the card, after anything SharePoint's canvas does for it.
   */
  private _claimPane(): void {
    document.dispatchEvent(new CustomEvent<string>(CARD_PANE_CLAIMED, { detail: this.instanceId }));
    this._paneRequestedAt = Date.now();
    window.clearTimeout(this._openTimer);
    window.clearTimeout(this._verifyTimer);
    this._openTimer = window.setTimeout(() => {
      // The author may have left the card, or edit mode, before this runs.
      if (this._selectedCardId !== undefined && this.displayMode === DisplayMode.Edit) {
        this.context.propertyPane.open();
      }
    }, 0);
  }

  private _onPaneClaimed = (event: Event): void => {
    if ((event as CustomEvent<string>).detail !== this.instanceId) {
      this._endCardEditing(false);
    }
  };

  /** Back to "no card selected": the pane closed or moved on, or the author asked for the settings. */
  private _endCardEditing(refreshPane: boolean): void {
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    window.clearTimeout(this._verifyTimer);
    if (this._selectedCardId === undefined) {
      return;
    }
    this._selectedCardId = undefined;
    this.render();
    if (refreshPane) {
      this.context.propertyPane.refresh();
    }
  }

  /**
   * The pane started showing this web part. While a card is being edited that
   * is either this web part's own request, or SharePoint re-drawing a pane
   * that was already here; the card stays in both cases. Otherwise the author
   * opened the pane afresh, for example with the web part's edit button, and
   * gets the general settings.
   */
  protected onPropertyPaneConfigurationStart(): void {
    window.clearTimeout(this._closeCheckTimer);
    if (this._selectedCardId === undefined) {
      return;
    }
    const now = Date.now();
    const ownRequest = this._paneRequestedAt !== undefined && now - this._paneRequestedAt < OWN_PANE_REQUEST_MS;
    const alreadyHere = this._paneCompletedAt !== undefined && now - this._paneCompletedAt < PANE_RETARGET_MS;
    if (!ownRequest && !alreadyHere) {
      this._endCardEditing(true);
      return;
    }
    // SharePoint's canvas can send its own request for the pane around the same
    // click. If it had the last word the pane is no longer marked as opened by
    // a web part, so ask for it once more.
    window.clearTimeout(this._verifyTimer);
    this._verifyTimer = window.setTimeout(() => {
      if (this._selectedCardId !== undefined && !this.context.propertyPane.isRenderedByWebPart()) {
        this._claimPane();
      }
    }, 0);
  }

  /**
   * SharePoint raises this when the pane closes or moves to another web part,
   * but also 5 seconds after the author's last change and whenever focus
   * leaves the pane, while it is still open on this web part. Only the first
   * two end card editing: a card's settings stay up until the author presses
   * "Done". The check waits a moment so the pane has settled
   * before it is asked; a pane that SharePoint moved elsewhere is no longer
   * marked as opened by a web part.
   */
  protected onPropertyPaneConfigurationComplete(): void {
    this._paneCompletedAt = Date.now();
    window.clearTimeout(this._closeCheckTimer);
    this._closeCheckTimer = window.setTimeout(() => {
      this._closeCheckTimer = undefined;
      const pane = this.context.propertyPane;
      if (this.displayMode === DisplayMode.Edit && pane.isPropertyPaneOpen() && pane.isRenderedByWebPart()) {
        return;
      }
      this._endCardEditing(false);
    }, PANE_SETTLE_MS);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // The colour fields render their own previews, and switching between icon
    // and image swaps which fields are shown, so rebuild the pane for these.
    // (The image picker tracks its own value, so an upload doesn't rebuild the
    // pane and interrupt its "Uploaded" announcement.)
    if (propertyPath === 'accentColor' || /\.(accentColor|mediaType)$/.test(propertyPath)) {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const index = this._indexOf(this._selectedCardId);
    if (index === -1) {
      this._selectedCardId = undefined;
      return { pages: [{ header: { description: strings.PaneDescription }, groups: this._settingsGroups() }] };
    }
    return {
      pages: [
        {
          header: { description: strings.CardPaneDescription.replace('{0}', String(index + 1)).replace('{1}', String(this.properties.cards.length)) },
          groups: this._cardGroups(index)
        }
      ]
    };
  }

  private _settingsGroups(): IPropertyPaneGroup[] {
    const cards = this.properties.cards;
    const count = cards.length === 1 ? strings.OneCard : strings.ManyCards.replace('{0}', String(cards.length));

    return [
      {
        groupName: strings.CardsGroupName,
        groupFields: [
          PropertyPaneLabel('cardsSummary', { text: `${count} ${strings.CardsHelp}` }),
          ...cards.map((card, i) =>
            PropertyPaneButton('cards', {
              text: `${i + 1}. ${card.title.trim() || strings.UntitledCard}`,
              ariaLabel: `${strings.EditCard}: ${card.title.trim() || strings.UntitledCard}`,
              buttonType: PropertyPaneButtonType.Command,
              icon: 'Edit',
              onClick: this._cardsButton(() => this._editCard(card.id))
            })
          ),
          PropertyPaneButton('cards', {
            text: strings.AddCard,
            buttonType: PropertyPaneButtonType.Primary,
            icon: 'Add',
            onClick: this._cardsButton(() => {
              this._addCard();
            })
          })
        ]
      },
      {
        groupName: strings.AppearanceGroupName,
        groupFields: [
          reactPropertyPaneField('accentColor', 'accent-color', (onChange) =>
            React.createElement(AccentColorPicker, {
              label: strings.AccentLabel,
              value: this.properties.accentColor,
              onChange: (value: string) => onChange(value || DEFAULT_ACCENT)
            })
          ),
          PropertyPaneToggle('useInterFont', {
            label: strings.UseInterFontLabel,
            onText: strings.UseInterFontOn,
            offText: strings.UseInterFontOff
          }),
          PropertyPaneDropdown('headingLevel', {
            label: strings.HeadingLevelLabel,
            options: [2, 3, 4, 5, 6].map((level) => ({ key: level, text: `${strings.HeadingLevelOption} ${level}` }))
          }),
          PropertyPaneLabel('headingLevelHelp', { text: strings.HeadingLevelHelp })
        ]
      }
    ];
  }

  private _cardGroups(index: number): IPropertyPaneGroup[] {
    const card = this.properties.cards[index];
    const path = (field: keyof IResourceCard): string => `cards[${index}].${field}`;
    const isFirst = index === 0;
    const isLast = index === this.properties.cards.length - 1;

    return [
      {
        groupName: strings.CardContentGroupName,
        groupFields: [
          PropertyPaneTextField(path('title'), {
            label: strings.TitleLabel,
            onGetErrorMessage: (value: string) => ((value || '').trim() ? '' : strings.TitleRequired),
            deferredValidationTime: 400
          }),
          PropertyPaneTextField(path('description'), {
            label: strings.DescriptionLabel,
            multiline: true,
            rows: 4,
            resizable: true
          }),
          PropertyPaneTextField(path('linkText'), {
            label: strings.LinkTextLabel,
            description: strings.LinkTextHelp
          }),
          PropertyPaneTextField(path('url'), {
            label: strings.UrlLabel,
            description: strings.UrlHelp,
            placeholder: 'https://',
            onGetErrorMessage: (value: string) => {
              const result = checkUrl(value);
              return result.ok ? '' : result.reason;
            },
            deferredValidationTime: 400
          }),
          PropertyPaneToggle(path('openInNewTab'), {
            label: strings.NewTabLabel,
            onText: strings.On,
            offText: strings.Off
          })
        ]
      },
      {
        groupName: strings.IconGroupName,
        groupFields: [
          PropertyPaneChoiceGroup(path('mediaType'), {
            label: strings.MediaTypeLabel,
            options: [
              { key: 'icon', text: strings.MediaTypeIcon },
              { key: 'image', text: strings.MediaTypeImage }
            ]
          }),
          ...(card.mediaType === 'image'
            ? [
                reactPropertyPaneField(path('imageUrl'), `image-${card.id}`, (onChange) =>
                  React.createElement(ImagePicker, {
                    label: strings.ImageLabel,
                    value: card.imageUrl,
                    onChange: (url: string) => onChange(url),
                    upload: (file: File) => uploadImageToSiteAssets(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, file)
                  })
                ),
                PropertyPaneTextField(path('imageAlt'), {
                  label: strings.ImageAltLabel,
                  description: strings.ImageAltHelp,
                  multiline: true,
                  rows: 2
                }),
                PropertyPaneDropdown(path('imageSize'), {
                  label: strings.ImageSizeLabel,
                  options: [
                    { key: 'small', text: strings.ImageSizeSmall },
                    { key: 'medium', text: strings.ImageSizeMedium },
                    { key: 'large', text: strings.ImageSizeLarge }
                  ]
                })
              ]
            : [
                reactPropertyPaneField(path('icon'), `icon-${card.id}`, (onChange) =>
                  React.createElement(IconPicker, {
                    label: strings.IconLabel,
                    value: card.icon,
                    onChange: (name: string) => onChange(name)
                  })
                )
              ]),
          reactPropertyPaneField(path('accentColor'), `accent-${card.id}`, (onChange) =>
            React.createElement(AccentColorPicker, {
              label: strings.CardAccentLabel,
              value: card.accentColor,
              inheritLabel: strings.CardAccentInherit,
              inheritedColor: this.properties.accentColor,
              onChange: (value: string) => onChange(value)
            })
          )
        ]
      },
      {
        groupName: strings.CardActionsGroupName,
        groupFields: [
          PropertyPaneButton('cards', {
            text: strings.MoveEarlier,
            buttonType: PropertyPaneButtonType.Normal,
            icon: 'Back',
            disabled: isFirst,
            onClick: this._cardsButton(() => this._moveCard(card.id, -1))
          }),
          PropertyPaneButton('cards', {
            text: strings.MoveLater,
            buttonType: PropertyPaneButtonType.Normal,
            icon: 'Forward',
            disabled: isLast,
            onClick: this._cardsButton(() => this._moveCard(card.id, 1))
          }),
          PropertyPaneButton('cards', {
            text: strings.DeleteCard,
            buttonType: PropertyPaneButtonType.Normal,
            icon: 'Delete',
            onClick: this._cardsButton(() => this._deleteCard(card.id))
          }),
          PropertyPaneButton('cards', {
            text: strings.BackToSettings,
            buttonType: PropertyPaneButtonType.Primary,
            onClick: this._cardsButton(this._showAllSettings)
          })
        ]
      }
    ];
  }

  /** Fills in anything missing, e.g. on a web part added before a property existed. */
  private _applyDefaults(): void {
    const p = this.properties;
    p.cards = normalizeCards(p.cards);
    p.accentColor = normalizeHex(p.accentColor) || DEFAULT_ACCENT;
    // Older versions could draw a panel behind the cards; it no longer exists.
    delete (p as { showPanel?: boolean }).showPanel;
    // Earlier versions saved this setting as "useHubFont"; carry it over.
    const legacy = p as { useHubFont?: boolean };
    if (p.useInterFont === undefined && legacy.useHubFont !== undefined) {
      p.useInterFont = legacy.useHubFont;
    }
    delete legacy.useHubFont;
    p.useInterFont = p.useInterFont !== false;
    const level = Number(p.headingLevel);
    p.headingLevel = (level >= 2 && level <= 6 ? level : 3) as HeadingLevel;
  }
}
