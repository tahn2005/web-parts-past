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
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import * as strings from 'LinkCardsWebPartStrings';

import { LinkCards, ILinkCardsProps } from './components/LinkCards';
import { createCard, ILinkCard, ILinkCardsWebPartProps, normalizeProperties, PALETTES } from './models';
import { resolveSectionPalette, resolveTheme } from './theme';
import { validateLink } from './urls';

const INTER_ID = 'link-cards-inter-font';
const INTER_HREF = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap';
const INTER_STACK = "'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";
const SHAREPOINT_STACK =
  "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";

/** Property-pane buttons and labels need a target property; this one is never persisted. */
const PANE_ACTION = 'paneAction';

type PaneControls = typeof import('./propertyPane/paneControls');

export default class LinkCardsWebPart extends BaseClientSideWebPart<ILinkCardsWebPartProps> {
  /** The card whose settings the property pane is showing, if any. */
  private _selectedCardId: string | undefined;
  private _pane: PaneControls | undefined;
  private _refreshTimer: number | undefined;

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

    const element = React.createElement<ILinkCardsProps>(LinkCards, {
      cards: p.cards,
      cardVars: resolveTheme(p).cards,
      columns: p.columns,
      fontFamily: p.useInterFont ? INTER_STACK : SHAREPOINT_STACK,
      isEditMode,
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
    const copy: ILinkCard = { ...cloneDeep(this.properties.cards[index]), id: Guid.newGuid().toString() };
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
    // eslint-disable-next-line no-alert
    if (!window.confirm(strings.ConfirmDeleteCard.replace('{0}', card.text.trim() || strings.UntitledCard))) return;
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
      /* webpackChunkName: 'link-cards-property-pane' */
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
    this.render();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === PANE_ACTION) {
      delete (this.properties as unknown as Record<string, unknown>)[PANE_ACTION];
      this.context.propertyPane.refresh();
      return;
    }

    // Switching to "Custom color" starts from the colour that was showing.
    if (propertyPath === 'palette' && newValue === 'custom' && typeof oldValue === 'string') {
      const previous = PALETTES[oldValue as keyof typeof PALETTES];
      if (previous) this.properties.iconColor = previous.accent;
    }
    const mode = /^cards\[(\d+)\]\.iconColorMode$/.exec(propertyPath);
    if (mode && newValue === 'custom') {
      const card = this.properties.cards[Number(mode[1])];
      if (card && !card.iconColor) card.iconColor = resolveSectionPalette(this.properties).accent;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();

    // Fields that change which other fields are shown need the pane rebuilt
    // now. Colour changes only affect the Accessibility notes, so they wait
    // until the author pauses.
    if (/(^|\.)(palette|iconColorMode)$/.test(propertyPath)) {
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

  private _sectionPane(): IPropertyPaneConfiguration {
    const p = this.properties;
    const theme = resolveTheme(p);

    const cardFields: IPropertyPaneField<unknown>[] = [PropertyPaneLabel(PANE_ACTION, { text: strings.CardsHelp })];
    p.cards.forEach((card, i) => {
      cardFields.push(
        this._button(
          strings.EditCardButton.replace('{0}', String(i + 1)).replace('{1}', card.text.trim() || strings.UntitledCard),
          'Edit',
          () => this._selectCard(card.id)
        )
      );
    });
    cardFields.push(this._button(strings.AddCard, 'Add', this._addCard, true));

    const appearance: IPropertyPaneField<unknown>[] = [
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
    if (p.palette === 'custom') appearance.push(this._color('iconColor', strings.IconColorField, p.iconColor));
    appearance.push(
      this._color('textColor', strings.TextColorField, p.textColor),
      this._color('cardBackground', strings.CardBackgroundField, p.cardBackground),
      PropertyPaneDropdown('columns', {
        label: strings.ColumnsField,
        options: [
          { key: 'auto', text: strings.ColumnsAuto },
          { key: '1', text: strings.ColumnsOne },
          { key: '2', text: strings.ColumnsTwo },
          { key: '3', text: strings.ColumnsThree }
        ]
      })
    );

    const groups: IPropertyPaneGroup[] = [
      { groupName: strings.CardsGroup, groupFields: cardFields },
      { groupName: strings.AppearanceGroup, groupFields: appearance },
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
                ...theme.adjustments.map((a) => PropertyPaneLabel(PANE_ACTION, { text: `${a.label}: ${a.from} → ${a.to}` }))
              ]
      }
    ];
    // SharePoint only lets authors open collapsed groups when displayGroupsAsAccordion is set.
    return {
      pages: [{ displayGroupsAsAccordion: true, header: { description: strings.SectionPaneDescription }, groups }]
    };
  }

  private _cardPane(index: number): IPropertyPaneConfiguration {
    const card = this.properties.cards[index];
    const total = this.properties.cards.length;
    const path = `cards[${index}].`;

    const iconFields: IPropertyPaneField<unknown>[] = [
      this.pane.PropertyPaneIconPicker(`${path}iconName`, { key: `icon-${card.id}`, label: strings.IconField, value: card.iconName }),
      PropertyPaneChoiceGroup(`${path}iconColorMode`, {
        label: strings.IconColorModeField,
        options: [
          { key: 'section', text: strings.IconColorSection },
          { key: 'custom', text: strings.IconColorCustom }
        ]
      })
    ];
    if (card.iconColorMode === 'custom') {
      iconFields.push(this._color(`${path}iconColor`, strings.CardIconColorField, card.iconColor));
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
          header: { description: strings.CardPaneDescription.replace('{0}', String(index + 1)).replace('{1}', String(total)) },
          groups: [
            { groupFields: [this._button(strings.BackToSection, 'Back', () => this._selectCard(undefined))] },
            { groupName: strings.CardTextGroup, groupFields: [PropertyPaneTextField(`${path}text`, { label: strings.CardTextField })] },
            {
              groupName: strings.CardLinkGroup,
              groupFields: [
                PropertyPaneTextField(`${path}url`, {
                  label: strings.LinkUrlField,
                  description: strings.LinkUrlHelp,
                  placeholder: 'https://',
                  onGetErrorMessage: validateLink,
                  deferredValidationTime: 400
                }),
                PropertyPaneToggle(`${path}openInNewTab`, { label: strings.OpenInNewTabField })
              ]
            },
            { groupName: strings.CardIconGroup, groupFields: iconFields },
            { groupName: strings.ArrangeGroup, groupFields: arrange }
          ]
        }
      ]
    };
  }

  private _ensureInterFont(): void {
    if (document.getElementById(INTER_ID)) return;
    const link = document.createElement('link');
    link.id = INTER_ID;
    link.rel = 'stylesheet';
    link.href = INTER_HREF;
    document.head.appendChild(link);
  }
}
