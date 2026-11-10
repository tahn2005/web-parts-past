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

export default class ResourceCardsWebPart extends BaseClientSideWebPart<IResourceCardsWebPartProps> {
  private _root: Root | undefined;
  /** Card whose settings are showing in the property pane, if any. */
  private _selectedCardId: string | undefined;

  protected onInit(): Promise<void> {
    this._applyDefaults();
    registerCardFont();
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
    this._selectedCardId = undefined;
    super.onDisplayModeChanged(oldDisplayMode);
    this.render();
  }

  protected onDispose(): void {
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
    this.render();
    if (openPane && !this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.open();
    } else if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
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

  protected onPropertyPaneConfigurationComplete(): void {
    // Pane closed: drop the selection highlight on the canvas.
    if (this._selectedCardId !== undefined) {
      this._selectedCardId = undefined;
      this.render();
    }
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
