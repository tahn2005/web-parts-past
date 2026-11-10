import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { set } from '@microsoft/sp-lodash-subset';
import type { IFilePickerResult } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import * as strings from 'CustomWidthHeaderWebPartStrings';

import { CustomWidthHeader, ICustomWidthHeaderProps } from './components/CustomWidthHeader';
import { ICustomWidthHeaderWebPartProps, normalizeProperties, PALETTES } from './models';
import { resolveBaseColor, resolveHeaderTheme } from './theme';
import { validateImageUrl, validateLink } from './urls';
import { detectHeadingLevel } from './headingLevel';

const FONTS_ID = 'custom-width-header-fonts';
const FONTS_HREF =
  'https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Inter:wght@400;500;600;700&display=swap';
const INTER_STACK = "'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";
const MERRIWEATHER_STACK = "'Merriweather', Georgia, 'Times New Roman', serif";
const SHAREPOINT_STACK =
  "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif";

/** Labels in the pane need a target property; this one is never persisted. */
const PANE_NOTE = 'paneNote';
/** The PnP file picker writes its raw result here; it is removed straight away. */
const PICKER_SCRATCH = 'imagePickerResult';

type PaneControls = typeof import('./propertyPane/paneControls');

export default class CustomWidthHeaderWebPart extends BaseClientSideWebPart<ICustomWidthHeaderWebPartProps> {
  private _pane: PaneControls | undefined;
  private _uploadStatus: Record<string, string> = {};
  private _refreshTimer: number | undefined;
  private _headingTimers: number[] = [];

  protected onInit(): Promise<void> {
    normalizeProperties(this.properties);
    // SharePoint can finish drawing its title area after the web part's first
    // render, so look again shortly afterwards; render only if the answer changed.
    this._headingTimers = [400, 1500, 4000].map((delay) =>
      window.setTimeout(() => {
        if (this.properties.headingLevel === 'auto' && detectHeadingLevel(this.domElement) !== this._lastHeadingLevel) {
          this.render();
        }
      }, delay)
    );
    return super.onInit();
  }

  private _lastHeadingLevel: '1' | '2' | undefined;

  private _resolveHeadingLevel(): '1' | '2' {
    const level = this.properties.headingLevel;
    this._lastHeadingLevel = level === '1' || level === '2' ? level : detectHeadingLevel(this.domElement);
    return this._lastHeadingLevel;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  /**
   * Version 1.0 always used Heading 1 by default, which skipped or duplicated
   * a level on some pages. Headers saved with that default move to 'auto'.
   */
  protected onAfterDeserialize(deserializedObject: unknown, dataVersion: Version): ICustomWidthHeaderWebPartProps {
    const props = deserializedObject as ICustomWidthHeaderWebPartProps;
    if (dataVersion && dataVersion.major === 1 && dataVersion.minor === 0 && props && props.headingLevel === '1') {
      props.headingLevel = 'auto';
    }
    return props;
  }

  public render(): void {
    normalizeProperties(this.properties);
    const p = this.properties;
    if (p.useManagerHubFonts) this._ensureFonts();

    const element = React.createElement<ICustomWidthHeaderProps>(CustomWidthHeader, {
      ...p,
      isEditMode: this.displayMode === DisplayMode.Edit,
      siteHomeUrl: this.context.pageContext.web.absoluteUrl,
      themeVars: resolveHeaderTheme(p).vars,
      bodyFont: p.useManagerHubFonts ? INTER_STACK : SHAREPOINT_STACK,
      titleFont: p.useManagerHubFonts ? MERRIWEATHER_STACK : SHAREPOINT_STACK,
      resolvedHeadingLevel: this._resolveHeadingLevel()
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    window.clearTimeout(this._refreshTimer);
    this._headingTimers.forEach((timer) => window.clearTimeout(timer));
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // Property pane ----------------------------------------------------------

  protected async loadPropertyPaneResources(): Promise<void> {
    this._pane = await import(
      /* webpackChunkName: 'custom-width-header-property-pane' */
      './propertyPane/paneControls'
    );
  }

  /** Set once loadPropertyPaneResources() has run, which SPFx does before asking for the pane. */
  private get pane(): PaneControls {
    if (!this._pane) throw new Error('Property pane resources have not loaded.');
    return this._pane;
  }

  protected onPropertyPaneConfigurationComplete(): void {
    this._uploadStatus = {};
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    const bag = this.properties as unknown as Record<string, unknown>;
    if (propertyPath === PANE_NOTE || propertyPath === PICKER_SCRATCH) {
      delete bag[propertyPath];
      return;
    }

    // Switching to "Custom" starts from the colour that was showing.
    if (propertyPath === 'palette' && newValue === 'custom' && typeof oldValue === 'string') {
      const previous = PALETTES[oldValue as keyof typeof PALETTES];
      if (previous) this.properties.backgroundColor = previous;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();

    // Fields that change which other fields are shown need the pane rebuilt
    // now; colour and strength changes only affect the Accessibility notes, so
    // they wait until the author pauses rather than interrupting a picker.
    if (/^(palette|backgroundType|imageTextStyle|iconSource|showBackLink|headingLevel)$/.test(propertyPath)) {
      this.context.propertyPane.refresh();
    } else if (/(Color|Opacity|useGradient)$/.test(propertyPath)) {
      window.clearTimeout(this._refreshTimer);
      this._refreshTimer = window.setTimeout(() => this.context.propertyPane.refresh(), 1200);
    }
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

  /** File picker plus an editable address field, for any image property. */
  private _imageFields(path: 'iconImageUrl' | 'backgroundImageUrl', label: string): IPropertyPaneField<unknown>[] {
    const fields: IPropertyPaneField<unknown>[] = [
      this.pane.PropertyFieldFilePicker(PICKER_SCRATCH, {
        key: `picker-${path}`,
        context: this.context as never,
        label,
        buttonLabel: strings.ChooseImage,
        accepts: this.pane.IMAGE_EXTENSIONS,
        filePickerResult: undefined as unknown as IFilePickerResult,
        properties: this.properties,
        hideWebSearchTab: true,
        // Default is to remember the tab in the page URL's #hash.
        storeLastActiveTab: false,
        onSave: (result: IFilePickerResult) => this._applyPickedImage(result, path),
        onPropertyChange: () => {
          delete (this.properties as unknown as Record<string, unknown>)[PICKER_SCRATCH];
        }
      }),
      PropertyPaneTextField(path, {
        label: strings.ImageUrlField,
        onGetErrorMessage: validateImageUrl,
        deferredValidationTime: 400
      })
    ];
    if (this._uploadStatus[path]) fields.push(PropertyPaneLabel(PANE_NOTE, { text: this._uploadStatus[path] }));
    return fields;
  }

  private async _applyPickedImage(result: IFilePickerResult, path: string): Promise<void> {
    try {
      let url = result.fileAbsoluteUrl;
      if (!url) {
        this._uploadStatus[path] = strings.Uploading;
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
      delete this._uploadStatus[path];
    } catch (error) {
      this._uploadStatus[path] = strings.UploadFailed.replace('{0}', (error as Error).message);
    }
    this.render();
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    normalizeProperties(this.properties);
    const p = this.properties;
    const theme = resolveHeaderTheme(p);

    const textFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('showBackLink', { label: strings.ShowBackLinkField })
    ];
    if (p.showBackLink) {
      textFields.push(
        PropertyPaneTextField('backLinkText', { label: strings.BackLinkTextField }),
        PropertyPaneTextField('backLinkUrl', {
          label: strings.BackLinkUrlField,
          placeholder: strings.BackLinkUrlPlaceholder,
          onGetErrorMessage: validateLink,
          deferredValidationTime: 400
        })
      );
    }
    textFields.push(
      PropertyPaneTextField('eyebrow', { label: strings.EyebrowField, description: strings.EyebrowHelp }),
      PropertyPaneTextField('title', { label: strings.TitleField }),
      PropertyPaneTextField('description', { label: strings.DescriptionField, multiline: true, rows: 3 })
    );

    const iconFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('iconSource', {
        label: strings.IconSourceField,
        options: [
          { key: 'none', text: strings.IconSourceNone },
          { key: 'icon', text: strings.IconSourceIcon },
          { key: 'image', text: strings.IconSourceImage }
        ]
      })
    ];
    if (p.iconSource === 'icon') {
      iconFields.push(this.pane.PropertyPaneIconPicker('iconName', { key: 'icon', label: strings.IconField, value: p.iconName }));
    } else if (p.iconSource === 'image') {
      iconFields.push(...this._imageFields('iconImageUrl', strings.IconImageField));
    }

    const backgroundFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('backgroundType', {
        label: strings.BackgroundTypeField,
        options: [
          { key: 'color', text: strings.BackgroundColor },
          { key: 'image', text: strings.BackgroundImage }
        ]
      })
    ];
    if (p.backgroundType === 'color') {
      backgroundFields.push(
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
      );
      if (p.palette === 'custom') {
        backgroundFields.push(this._color('backgroundColor', strings.BackgroundColorField, resolveBaseColor(p)));
      }
      backgroundFields.push(
        PropertyPaneToggle('useGradient', { label: strings.GradientField }),
        PropertyPaneLabel(PANE_NOTE, { text: strings.GradientHelp })
      );
    } else {
      backgroundFields.push(
        ...this._imageFields('backgroundImageUrl', strings.BackgroundImageField),
        PropertyPaneDropdown('imagePosition', {
          label: strings.ImagePositionField,
          options: [
            { key: 'center', text: strings.PositionCenter },
            { key: 'top', text: strings.PositionTop },
            { key: 'bottom', text: strings.PositionBottom }
          ]
        }),
        PropertyPaneChoiceGroup('imageTextStyle', {
          label: strings.ImageTextStyleField,
          options: [
            { key: 'overlay', text: strings.ImageTextOverlay },
            { key: 'panel', text: strings.ImageTextPanel }
          ]
        }),
        PropertyPaneLabel(PANE_NOTE, { text: strings.ImageTextHelp })
      );
      if (p.imageTextStyle === 'panel') {
        backgroundFields.push(this._color('overlayColor', strings.PanelColorField, p.overlayColor));
      } else {
        backgroundFields.push(
          this._color('overlayColor', strings.OverlayColorField, p.overlayColor),
          PropertyPaneSlider('overlayOpacity', { label: strings.OverlayOpacityField, min: 0, max: 100, step: 5 })
        );
      }
    }
    backgroundFields.push(this._color('textColor', strings.TextColorField, p.textColor));

    const accessibilityFields: IPropertyPaneField<unknown>[] =
      theme.adjustments.length === 0
        ? [PropertyPaneLabel(PANE_NOTE, { text: strings.AccessibilityAllPass })]
        : [
            PropertyPaneLabel(PANE_NOTE, { text: strings.AccessibilityAdjusted }),
            ...theme.adjustments.map((a) => PropertyPaneLabel(PANE_NOTE, { text: `${a.label}: ${a.from} → ${a.to}` }))
          ];
    if (p.backgroundType === 'image' && p.imageTextStyle === 'overlay') {
      accessibilityFields.push(PropertyPaneLabel(PANE_NOTE, { text: strings.AccessibilityImageNote }));
    }
    if (p.headingLevel !== 'auto') {
      const fits = detectHeadingLevel(this.domElement);
      if (fits !== p.headingLevel) {
        accessibilityFields.push(
          PropertyPaneLabel(PANE_NOTE, { text: strings.HeadingLevelMismatch.replace('{0}', p.headingLevel).replace('{1}', fits) })
        );
      }
    }

    const groups: IPropertyPaneGroup[] = [
      { groupName: strings.TextGroup, groupFields: textFields },
      { groupName: strings.IconGroup, groupFields: iconFields },
      { groupName: strings.BackgroundGroup, groupFields: backgroundFields },
      {
        groupName: strings.LayoutGroup,
        isCollapsed: true,
        groupFields: [
          PropertyPaneDropdown('layout', {
            label: strings.LayoutField,
            options: [
              { key: 'page', text: strings.LayoutPage },
              { key: 'resource', text: strings.LayoutResource }
            ]
          }),
          PropertyPaneDropdown('height', {
            label: strings.HeightField,
            options: [
              { key: 'auto', text: strings.HeightAuto },
              { key: 'medium', text: strings.HeightMedium },
              { key: 'large', text: strings.HeightLarge }
            ]
          }),
          PropertyPaneDropdown('headingLevel', {
            label: strings.HeadingLevelField,
            options: [
              { key: 'auto', text: strings.HeadingLevelAuto },
              { key: '1', text: 'Heading 1' },
              { key: '2', text: 'Heading 2' }
            ]
          }),
          PropertyPaneLabel(PANE_NOTE, { text: strings.HeadingLevelHelp })
        ]
      },
      {
        groupName: strings.FontGroup,
        isCollapsed: true,
        groupFields: [PropertyPaneToggle('useManagerHubFonts', { label: strings.UseFontsField })]
      },
      {
        groupName: strings.AccessibilityGroup,
        isCollapsed: accessibilityFields.length === 1 && theme.adjustments.length === 0 && p.backgroundType === 'color',
        groupFields: accessibilityFields
      }
    ];

    return { pages: [{ header: { description: strings.PaneDescription }, groups }] };
  }

  private _ensureFonts(): void {
    if (document.getElementById(FONTS_ID)) return;
    const link = document.createElement('link');
    link.id = FONTS_ID;
    link.rel = 'stylesheet';
    link.href = FONTS_HREF;
    document.head.appendChild(link);
  }
}
