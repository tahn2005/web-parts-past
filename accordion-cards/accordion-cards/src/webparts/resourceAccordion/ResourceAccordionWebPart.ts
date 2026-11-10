import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Guid, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { cloneDeep, set } from '@microsoft/sp-lodash-subset';
import type { IFilePickerResult } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import * as strings from 'ResourceAccordionWebPartStrings';

import { AccordionCards, IAccordionCardsProps } from './components/AccordionCards';
import { createCard, IAccordionCardsWebPartProps, IResourceCard, normalizeProperties, PALETTES } from './models';
import { resolveSectionPalette, resolveTheme } from './theme';
import { deriveTints } from './colors';
import { validateImageUrl, validateLink } from './urls';

const INTER_STYLESHEET_ID = 'accordion-cards-inter-font';
const INTER_STYLESHEET =
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap';
const INTER_STACK = "'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";
const SHAREPOINT_STACK =
  "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";

/** Property-pane buttons need a target property; this one is never persisted. */
const PANE_ACTION = 'paneAction';
/** The PnP file picker writes its raw result here; it is removed straight away. */
const PICKER_SCRATCH = 'imagePickerResult';

type PaneControls = typeof import('./propertyPane/paneControls');


export default class AccordionCardsWebPart extends BaseClientSideWebPart<IAccordionCardsWebPartProps> {
  /** The card whose settings the property pane is showing, if any. */
  private _selectedCardId: string | undefined;
  private _uploadStatus: string = '';
  private _refreshTimer: number | undefined;
  private _pane: PaneControls | undefined;

  protected onInit(): Promise<void> {
    normalizeProperties(this.properties);
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public render(): void {
    normalizeProperties(this.properties);
    const p = this.properties;
    if (p.useInterFont) this._ensureInterFont();

    const isEditMode = this.displayMode === DisplayMode.Edit;
    const element = React.createElement<IAccordionCardsProps>(AccordionCards, {
      instanceId: this.instanceId,
      isEditMode,
      title: p.title,
      description: p.description,
      headingLevel: p.headingLevel,
      iconSource: p.iconSource,
      iconName: p.iconName,
      imageUrl: p.imageUrl,
      showCount: p.showCount,
      startExpanded: p.startExpanded,
      cards: p.cards,
      defaultLinkText: p.defaultLinkText,
      theme: resolveTheme(p),
      fontFamily: p.useInterFont ? INTER_STACK : SHAREPOINT_STACK,
      selectedCardId: isEditMode ? this._selectedCardId : undefined,
      onSelectCard: this._selectCard,
      onAddCard: this._addCard,
      onMoveCard: this._moveCard,
      onDeleteCard: this._deleteCard
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    window.clearTimeout(this._refreshTimer);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // Card operations — shared by the canvas toolbar and the property pane --

  private _selectCard = (id: string | undefined): void => {
    this._selectedCardId = id;
    this._uploadStatus = '';
    this.render();
    if (id && !this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.open();
    } else {
      this.context.propertyPane.refresh();
    }
  };

  private _addCard = (): void => {
    const card = createCard(Guid.newGuid().toString());
    this.properties.cards = [...this.properties.cards, card];
    this._selectCard(card.id);
  };

  private _duplicateCard = (id: string): void => {
    const index = this._indexOf(id);
    if (index < 0) return;
    const copy: IResourceCard = { ...cloneDeep(this.properties.cards[index]), id: Guid.newGuid().toString() };
    const cards = [...this.properties.cards];
    cards.splice(index + 1, 0, copy);
    this.properties.cards = cards;
    this._selectCard(copy.id);
  };

  private _moveCard = (id: string, delta: number): void => {
    const from = this._indexOf(id);
    const to = from + delta;
    if (from < 0 || to < 0 || to >= this.properties.cards.length) return;
    const cards = [...this.properties.cards];
    const [card] = cards.splice(from, 1);
    cards.splice(to, 0, card);
    this.properties.cards = cards;
    this.render();
    this.context.propertyPane.refresh();
  };

  private _deleteCard = (id: string): void => {
    const card = this.properties.cards[this._indexOf(id)];
    if (!card) return;
    const name = card.title.trim() || strings.UntitledCard;
    // eslint-disable-next-line no-alert
    if (!window.confirm(strings.ConfirmDeleteCard.replace('{0}', name))) return;
    this.properties.cards = this.properties.cards.filter((c) => c.id !== id);
    if (this._selectedCardId === id) this._selectedCardId = undefined;
    this.render();
    this.context.propertyPane.refresh();
  };

  private _indexOf(id: string | undefined): number {
    return this.properties.cards.findIndex((c) => c.id === id);
  }

  // Property pane ----------------------------------------------------------

  protected async loadPropertyPaneResources(): Promise<void> {
    this._pane = await import(
      /* webpackChunkName: 'resource-accordion-property-pane' */
      './propertyPane/paneControls'
    );
  }

  /** Set once loadPropertyPaneResources() has run, which SPFx does before asking for the pane. */
  private get pane(): PaneControls {
    if (!this._pane) throw new Error('Property pane resources have not loaded.');
    return this._pane;
  }

  protected onPropertyPaneConfigurationComplete(): void {
    this._selectedCardId = undefined;
    this._uploadStatus = '';
    this.render();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    const bag = this.properties as unknown as Record<string, unknown>;
    if (propertyPath === PANE_ACTION) {
      delete bag[PANE_ACTION];
      this.context.propertyPane.refresh();
      return;
    }

    // Switching to "Custom" starts from the palette that was showing, not from scratch.
    if (propertyPath === 'palette' && newValue === 'custom' && typeof oldValue === 'string') {
      const previous = PALETTES[oldValue as keyof typeof PALETTES];
      if (previous) {
        this.properties.accentColor = previous.accent;
        this.properties.tintColor = previous.tint;
        this.properties.borderColor = previous.border;
      }
    }

    // Picking a section colour also sets the soft tint and border that go with
    // it, so the count badge and card hover never clash; both can still be
    // fine-tuned afterwards.
    if (propertyPath === 'accentColor' && typeof newValue === 'string' && newValue) {
      const derived = deriveTints(newValue);
      this.properties.tintColor = derived.tint;
      this.properties.borderColor = derived.border;
    }

    // A card switched to a custom accent starts from the section's accent.
    const accentMode = /^cards\[(\d+)\]\.accentMode$/.exec(propertyPath);
    if (accentMode && newValue === 'custom') {
      const card = this.properties.cards[Number(accentMode[1])];
      if (card && !card.accentColor) card.accentColor = resolveSectionPalette(this.properties).accent;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();

    // Fields that change which other fields are shown need the pane rebuilt now.
    // Colour changes only affect the Accessibility notes, so they wait until the
    // author pauses — rebuilding mid-drag would interrupt the colour picker.
    if (/(^|\.)(palette|iconSource|accentMode)$/.test(propertyPath)) {
      this.context.propertyPane.refresh();
    } else if (/(Color|Background)$/.test(propertyPath)) {
      window.clearTimeout(this._refreshTimer);
      this._refreshTimer = window.setTimeout(() => this.context.propertyPane.refresh(), 1200);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    normalizeProperties(this.properties);
    const index = this._indexOf(this._selectedCardId);
    if (index >= 0) return this._cardPane(index);
    this._selectedCardId = undefined;
    return this._sectionPane();
  }

  private _button(text: string, icon: string, action: () => void, primary = false): IPropertyPaneField<unknown> {
    return PropertyPaneButton(PANE_ACTION, {
      text,
      icon,
      buttonType: primary ? PropertyPaneButtonType.Primary : PropertyPaneButtonType.Normal,
      onClick: () => {
        action();
        return undefined;
      }
    });
  }

  private _color(path: string, label: string, value: string): IPropertyPaneField<unknown> {
    return this.pane.PropertyFieldColorPicker(path, {
      key: `color-${path}`,
      label,
      selectedColor: value,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      alphaSliderHidden: true,
      showPreview: true,
      style: this.pane.PropertyFieldColorPickerStyle.Inline,
      iconName: 'Precipitation'
    });
  }

  /** Icon-or-image fields, shared by the section header and every card. */
  private _iconFields(prefix: string, source: string, iconName: string, keySuffix: string): IPropertyPaneField<unknown>[] {
    const fields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup(`${prefix}iconSource`, {
        label: strings.IconSourceField,
        options: [
          { key: 'icon', text: strings.IconSourceIcon },
          { key: 'image', text: strings.IconSourceImage }
        ]
      })
    ];

    if (source === 'image') {
      fields.push(
        this.pane.PropertyFieldFilePicker(PICKER_SCRATCH, {
          key: `image-${keySuffix}`,
          context: this.context as never,
          label: strings.ImageField,
          buttonLabel: strings.ChooseImage,
          accepts: this.pane.IMAGE_EXTENSIONS,
          filePickerResult: undefined as unknown as IFilePickerResult,
          properties: this.properties,
          hideWebSearchTab: true,
          // Default is to remember the tab in the page URL's #hash.
          storeLastActiveTab: false,
          onSave: (result: IFilePickerResult) => this._applyPickedImage(result, `${prefix}imageUrl`),
          onPropertyChange: () => {
            delete (this.properties as unknown as Record<string, unknown>)[PICKER_SCRATCH];
          }
        }),
        PropertyPaneTextField(`${prefix}imageUrl`, {
          label: strings.ImageUrlField,
          onGetErrorMessage: validateImageUrl,
          deferredValidationTime: 400
        })
      );
      if (this._uploadStatus) fields.push(PropertyPaneLabel(PANE_ACTION, { text: this._uploadStatus }));
    } else {
      fields.push(
        this.pane.PropertyPaneIconPicker(`${prefix}iconName`, {
          key: `icon-${keySuffix}`,
          label: strings.IconField,
          value: iconName
        })
      );
    }
    return fields;
  }

  private async _applyPickedImage(result: IFilePickerResult, path: string): Promise<void> {
    try {
      let url = result.fileAbsoluteUrl;
      if (!url) {
        this._uploadStatus = strings.Uploading;
        this.context.propertyPane.refresh();
        const file = await result.downloadFileContent();
        url = await this.pane.uploadImage(
          this.context.spHttpClient,
          this.context.pageContext.web.absoluteUrl,
          file,
          result.fileName || file.name || 'image'
        );
      }
      set(this.properties, path, url);
      this._uploadStatus = '';
    } catch (error) {
      this._uploadStatus = strings.UploadFailed.replace('{0}', (error as Error).message);
    }
    this.render();
    this.context.propertyPane.refresh();
  }

  private _sectionPane(): IPropertyPaneConfiguration {
    const p = this.properties;
    const theme = resolveTheme(p);

    // The section's own colour sits with the section's other settings, so it's
    // as easy to find as each card's colour. Cards set to "Match the section"
    // follow it.
    const sectionColorFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('palette', {
        label: strings.PaletteField,
        options: [
          { key: 'blue', text: strings.PaletteBlue },
          { key: 'indigo', text: strings.PaletteIndigo },
          { key: 'green', text: strings.PaletteGreen },
          { key: 'amber', text: strings.PaletteAmber },
          { key: 'custom', text: strings.PaletteCustom }
        ]
      })
    ];
    if (p.palette === 'custom') {
      sectionColorFields.push(
        this._color('accentColor', strings.AccentColorField, p.accentColor),
        this._color('tintColor', strings.TintColorField, p.tintColor),
        this._color('borderColor', strings.BorderColorField, p.borderColor)
      );
    }
    sectionColorFields.push(PropertyPaneLabel(PANE_ACTION, { text: strings.SectionColorHelp }));

    const colorFields: IPropertyPaneField<unknown>[] = [];
    colorFields.push(
      this._color('headerBackground', strings.HeaderBackgroundField, p.headerBackground),
      this._color('headerHoverBackground', strings.HeaderHoverBackgroundField, p.headerHoverBackground),
      this._color('titleColor', strings.TitleColorField, p.titleColor),
      this._color('descriptionColor', strings.DescriptionColorField, p.descriptionColor),
      this._color('chevronColor', strings.ChevronColorField, p.chevronColor),
      this._color('bodyBackground', strings.BodyBackgroundField, p.bodyBackground)
    );

    const cardFields: IPropertyPaneField<unknown>[] = [PropertyPaneLabel(PANE_ACTION, { text: strings.CardsHelp })];
    p.cards.forEach((card, i) => {
      cardFields.push(
        this._button(
          strings.EditCardButton.replace('{0}', String(i + 1)).replace('{1}', card.title.trim() || strings.UntitledCard),
          'Edit',
          () => this._selectCard(card.id)
        )
      );
    });
    cardFields.push(
      this._button(strings.AddCard, 'Add', this._addCard, true),
      PropertyPaneTextField('defaultLinkText', { label: strings.DefaultLinkTextField }),
      this._color('cardBackground', strings.CardBackgroundField, p.cardBackground),
      this._color('cardTitleColor', strings.CardTitleColorField, p.cardTitleColor),
      this._color('cardDescriptionColor', strings.CardDescriptionColorField, p.cardDescriptionColor)
    );

    const groups: IPropertyPaneGroup[] = [
      {
        groupName: strings.HeaderGroup,
        groupFields: [
          PropertyPaneTextField('title', { label: strings.TitleField }),
          PropertyPaneTextField('description', { label: strings.DescriptionField, multiline: true, rows: 2 }),
          ...this._iconFields('', p.iconSource, p.iconName, 'header'),
          ...sectionColorFields,
          PropertyPaneToggle('showCount', { label: strings.ShowCountField }),
          PropertyPaneToggle('startExpanded', { label: strings.StartExpandedField }),
          PropertyPaneDropdown('headingLevel', {
            label: strings.HeadingLevelField,
            options: [
              { key: '2', text: 'Heading 2' },
              { key: '3', text: 'Heading 3' },
              { key: '4', text: 'Heading 4' }
            ]
          }),
          PropertyPaneLabel(PANE_ACTION, { text: strings.HeadingLevelHelp })
        ]
      },
      { groupName: strings.CardsGroup, groupFields: cardFields },
      { groupName: strings.ColorsGroup, isCollapsed: true, groupFields: colorFields },
      {
        groupName: strings.FontGroup,
        isCollapsed: true,
        groupFields: [PropertyPaneToggle('useInterFont', { label: strings.UseInterFontField })]
      },
      {
        groupName: strings.AccessibilityGroup,
        isCollapsed: theme.adjustments.length === 0,
        groupFields:
          theme.adjustments.length === 0
            ? [PropertyPaneLabel(PANE_ACTION, { text: strings.AccessibilityAllPass })]
            : [
                PropertyPaneLabel(PANE_ACTION, { text: strings.AccessibilityAdjusted }),
                ...theme.adjustments.map((a) =>
                  PropertyPaneLabel(PANE_ACTION, { text: `${a.label}: ${a.from} → ${a.to}` })
                )
              ]
      }
    ];

    // Collapsible groups only open in SharePoint when the page is an accordion.
    return {
      pages: [{ displayGroupsAsAccordion: true, header: { description: strings.SectionPaneDescription }, groups }]
    };
  }

  private _cardPane(index: number): IPropertyPaneConfiguration {
    const card = this.properties.cards[index];
    const total = this.properties.cards.length;
    const path = `cards[${index}].`;

    const colorFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup(`${path}accentMode`, {
        label: strings.CardAccentModeField,
        options: [
          { key: 'section', text: strings.CardAccentSection },
          { key: 'custom', text: strings.CardAccentCustom }
        ]
      })
    ];
    if (card.accentMode === 'custom') {
      colorFields.push(this._color(`${path}accentColor`, strings.CardAccentColorField, card.accentColor));
    }

    const arrange: IPropertyPaneField<unknown>[] = [];
    if (index > 0) arrange.push(this._button(strings.MoveCardEarlier, 'Back', () => this._moveCard(card.id, -1)));
    if (index < total - 1) arrange.push(this._button(strings.MoveCardLater, 'Forward', () => this._moveCard(card.id, 1)));
    arrange.push(
      this._button(strings.DuplicateCard, 'Copy', () => this._duplicateCard(card.id)),
      this._button(strings.DeleteCard, 'Delete', () => this._deleteCard(card.id))
    );

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.CardPaneDescription.replace('{0}', String(index + 1)).replace('{1}', String(total))
          },
          groups: [
            {
              groupFields: [this._button(strings.BackToSection, 'Back', () => this._selectCard(undefined))]
            },
            {
              groupName: strings.CardContentGroup,
              groupFields: [
                PropertyPaneTextField(`${path}title`, { label: strings.CardTitleField }),
                PropertyPaneTextField(`${path}description`, {
                  label: strings.CardDescriptionField,
                  multiline: true,
                  rows: 3
                })
              ]
            },
            {
              groupName: strings.CardIconGroup,
              groupFields: this._iconFields(path, card.iconSource, card.iconName, card.id)
            },
            {
              groupName: strings.CardLinkGroup,
              groupFields: [
                PropertyPaneTextField(`${path}linkUrl`, {
                  label: strings.LinkUrlField,
                  placeholder: 'https://',
                  onGetErrorMessage: validateLink,
                  deferredValidationTime: 400
                }),
                PropertyPaneTextField(`${path}linkText`, {
                  label: strings.LinkTextField,
                  placeholder: `${strings.LinkTextPlaceholder} (“${this.properties.defaultLinkText}”)`
                }),
                PropertyPaneToggle(`${path}openInNewTab`, { label: strings.OpenInNewTabField })
              ]
            },
            { groupName: strings.CardColorGroup, groupFields: colorFields },
            { groupName: strings.ArrangeGroup, groupFields: arrange }
          ]
        }
      ]
    };
  }

  private _ensureInterFont(): void {
    if (document.getElementById(INTER_STYLESHEET_ID)) return;
    const link = document.createElement('link');
    link.id = INTER_STYLESHEET_ID;
    link.rel = 'stylesheet';
    link.href = INTER_STYLESHEET;
    document.head.appendChild(link);
  }
}
