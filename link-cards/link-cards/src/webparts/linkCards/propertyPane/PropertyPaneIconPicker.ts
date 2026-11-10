import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IconPicker } from './IconPicker';

export interface IPropertyPaneIconPickerOptions {
  key: string;
  label: string;
  value: string;
}

/** Property pane field wrapping <IconPicker>. */
export function PropertyPaneIconPicker(
  targetProperty: string,
  options: IPropertyPaneIconPickerOptions
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: options.key,
      onRender: (
        element: HTMLElement,
        _context?: unknown,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void
      ): void => {
        ReactDom.render(
          React.createElement(IconPicker, {
            label: options.label,
            value: options.value,
            onChange: (name: string) => changeCallback?.(targetProperty, name)
          }),
          element
        );
      },
      onDispose: (element: HTMLElement): void => {
        ReactDom.unmountComponentAtNode(element);
      }
    }
  };
}
