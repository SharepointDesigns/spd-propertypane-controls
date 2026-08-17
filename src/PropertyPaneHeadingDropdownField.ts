import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";
import HeadingDropdown from "./HeadingDropdown";

export interface IPropertyPaneHeadingDropdownFieldOptions {
  key: string;
  selected: string;
  customSize?: number;
  onChange: (val: string) => void;
  onCustomSizeChange?: (val: number) => void;
}

/**
 * Property pane custom field with a heading-level dropdown (H1-H4/Normal) and,
 * when "Custom Font Size" is selected, a font-size slider.
 */
export function PropertyPaneHeadingDropdownField(
  opts: IPropertyPaneHeadingDropdownFieldOptions,
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: opts.key,
    properties: {
      key: opts.key,
      onRender: (
        elem: HTMLElement,
        _context?: unknown,
        changeCallback?: (targetProperty?: string, newValue?: unknown) => void,
      ) => {
        const element = React.createElement(HeadingDropdown, {
          selected: opts.selected,
          customSize: opts.customSize,
          onChange: (val: string) => {
            opts.onChange(val);
            changeCallback?.(opts.key, val);
          },
          onCustomSizeChange: (val: number) => {
            opts.onCustomSizeChange?.(val);
            changeCallback?.(opts.key, val.toString());
          },
        });
        ReactDom.render(element, elem);
      },
      onDispose: (elem: HTMLElement) => {
        ReactDom.unmountComponentAtNode(elem);
      },
    },
  };
}
