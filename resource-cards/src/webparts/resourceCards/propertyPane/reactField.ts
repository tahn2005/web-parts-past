import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';

const roots: Map<HTMLElement, Root> = new Map();

/**
 * Wraps a React component as a custom property pane field.
 *
 * SharePoint calls `onRender` every time the pane refreshes, so the same
 * React root is reused for an element and simply re-rendered with new props.
 */
export function reactPropertyPaneField<P>(
  targetProperty: string,
  key: string,
  render: (onChange: (value: unknown) => void) => React.ReactElement<P>
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key,
      onRender: (element, _context, changeCallback) => {
        let root = roots.get(element);
        if (!root) {
          root = createRoot(element);
          roots.set(element, root);
        }
        root.render(
          render((value) => {
            if (changeCallback) {
              changeCallback(targetProperty, value);
            }
          })
        );
      },
      onDispose: (element) => {
        const root = roots.get(element);
        if (root) {
          root.unmount();
          roots.delete(element);
        }
      }
    }
  };
}
