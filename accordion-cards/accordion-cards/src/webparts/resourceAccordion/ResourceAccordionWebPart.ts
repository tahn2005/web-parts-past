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

/** How long after a "configuration complete" event the pane is checked for being closed (see onPropertyPaneConfigurationComplete). */
const PANE_SETTLE_MS = 600;
/**
 * How long SharePoint may take to answer this web part's own open() with a
 * "configuration start". The first open also downloads the pane's code, which
 * can take a few seconds on a slow connection.
 */
const OWN_PANE_REQUEST_MS = 10000;
/** A "start" this soon after a "complete" is SharePoint re-drawing a pane that never left this web part. */
const PANE_RETARGET_MS = 250;
/** Raised on the document when a web part puts one of its cards in the pane, so any other one stops editing its card. */
const CARD_PANE_CLAIMED = 'card-settings-pane-claimed';

type PaneControls = typeof import('./propertyPane/paneControls');


export default class AccordionCardsWebPart extends BaseClientSideWebPart<IAccordionCardsWebPartProps> {
  /** The card whose settings the property pane is showing, if any. */
  private _selectedCardId: string | undefined;
  private _uploadStatus: string = '';
  private _refreshTimer: number | undefined;
  private _closeCheckTimer: number | undefined;
  private _openTimer: number | undefined;
  /** This web part asked for the pane itself (open()) and SharePoint hasn't answered with a "start" yet. */
  private _ownRequestPending: boolean = false;
  /** When this web part last asked for the pane itself, and when the pane last reported "complete". */
  private _paneRequestedAt: number | undefined;
  private _paneCompletedAt: number | undefined;
  /** The pane's last lifecycle event for this web part. */
  private _lastPaneEvent: 'start' | 'complete' | undefined;
  private _pane: PaneControls | undefined;

  protected onInit(): Promise<void> {
    normalizeProperties(this.properties);
    document.addEventListener(CARD_PANE_CLAIMED, this._onPaneClaimed);
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
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    document.removeEventListener(CARD_PANE_CLAIMED, this._onPaneClaimed);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    // Card settings belong to edit mode; the page was published or the edit discarded.
    if (this.displayMode !== DisplayMode.Edit) this._endCardEditing(false);
    super.onDisplayModeChanged(oldDisplayMode);
  }

  // Card operations — shared by the canvas toolbar and the property pane --

  private _selectCard = (id: string | undefined): void => {
    window.clearTimeout(this._closeCheckTimer);
    this._selectedCardId = id;
    this._uploadStatus = '';
    this.render();
    this.context.propertyPane.refresh();
    if (id) this._claimPane();
    else this._ownRequestPending = false;
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

  /**
   * Puts the pane on the selected card. open() (not just refresh()) also takes
   * the pane over when it shows something else. It runs on the next tick so it
   * is the last pane request for the click that selected the card, after
   * anything SharePoint's canvas does for the same click.
   */
  private _claimPane(): void {
    document.dispatchEvent(new CustomEvent<string>(CARD_PANE_CLAIMED, { detail: this.instanceId }));
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    this._ownRequestPending = true;
    this._paneRequestedAt = Date.now();
    this._openTimer = window.setTimeout(() => {
      // The author may have left the card, or edit mode, before this runs.
      if (this._selectedCardId !== undefined && this.displayMode === DisplayMode.Edit) {
        this.context.propertyPane.open();
      } else {
        this._ownRequestPending = false;
      }
    }, 0);
  }

  /** Another card web part on the page put one of its cards in the pane. */
  private _onPaneClaimed = (event: Event): void => {
    if ((event as CustomEvent<string>).detail !== this.instanceId) this._endCardEditing(false);
  };

  /** Back to "no card selected": the author asked for the settings, left edit mode, closed the pane or moved on. */
  private _endCardEditing(refreshPane: boolean): void {
    window.clearTimeout(this._closeCheckTimer);
    window.clearTimeout(this._openTimer);
    this._ownRequestPending = false;
    if (this._selectedCardId === undefined) return;
    this._selectedCardId = undefined;
    this._uploadStatus = '';
    this.render();
    if (refreshPane) this.context.propertyPane.refresh();
  }

  /**
   * The pane started showing this web part. A card's settings stay up when
   * this is the answer to this web part's own open(), or the pane never left
   * this web part: either no "complete" came since the last "start" (the
   * classic pane can send two "start"s for one opening when SharePoint's
   * canvas selects the web part during the opening animation), or the
   * "complete" came straight before it (SharePoint re-drawing the pane, for
   * example when the author clicks inside the web part). Otherwise the pane
   * has come back from somewhere else, such as another web part or the
   * section settings, or the author used the web part's own edit button:
   * they get the section settings, as with SharePoint's own Quick Links.
   */
  protected onPropertyPaneConfigurationStart(): void {
    window.clearTimeout(this._closeCheckTimer);
    const now = Date.now();
    const ownRequest =
      this._ownRequestPending && this._paneRequestedAt !== undefined && now - this._paneRequestedAt < OWN_PANE_REQUEST_MS;
    this._ownRequestPending = false;
    const neverLeft =
      this._lastPaneEvent === 'start' ||
      (this._paneCompletedAt !== undefined && now - this._paneCompletedAt < PANE_RETARGET_MS);
    this._lastPaneEvent = 'start';
    if (this._selectedCardId !== undefined && !ownRequest && !neverLeft) this._endCardEditing(true);
  }

  /**
   * SharePoint raises this when the pane closes or moves on, but also 5
   * seconds after the author's last change, when focus leaves the pane, and
   * just before it re-draws a pane that stays on this web part. None of those
   * may take the author off the card they are editing, so this only ends card
   * editing once the pane has really closed. (Moving to another web part or
   * to the section settings is handled when the pane comes back: see
   * onPropertyPaneConfigurationStart.) The pane is checked twice, a moment
   * apart, so a pane that is only re-drawing is never mistaken for a closed one.
   */
  protected onPropertyPaneConfigurationComplete(): void {
    this._paneCompletedAt = Date.now();
    this._lastPaneEvent = 'complete';
    window.clearTimeout(this._closeCheckTimer);
    if (this._selectedCardId === undefined) return;
    const paneClosed = (): boolean => this.displayMode !== DisplayMode.Edit || !this.context.propertyPane.isPropertyPaneOpen();
    this._closeCheckTimer = window.setTimeout(() => {
      if (!paneClosed()) return;
      this._closeCheckTimer = window.setTimeout(() => {
        this._closeCheckTimer = undefined;
        if (paneClosed()) this._endCardEditing(false);
      }, PANE_SETTLE_MS);
    }, PANE_SETTLE_MS);
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
      { groupName: strings.ColorsGroup, groupFields: colorFields },
      { groupName: strings.FontGroup, groupFields: [PropertyPaneToggle('useInterFont', { label: strings.UseInterFontField })] },
      {
        groupName: strings.AccessibilityGroup,
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

    // Every group is always open (no accordion, nothing collapsed), so no setting
    // can end up hidden behind a group the author can't expand.
    return { pages: [{ header: { description: strings.SectionPaneDescription }, groups }] };
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
