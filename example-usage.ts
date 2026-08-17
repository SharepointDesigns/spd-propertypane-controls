/**
 * Example: using @spdesigns/propertypane-controls in an SPFx web part.
 * Covers: single color picker, dual (theme swatch) color picker, image picker,
 * and heading dropdown.
 */

import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneGroup } from "@microsoft/sp-property-pane";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
  PropertyFieldFilePicker,
} from "@pnp/spfx-property-controls";
import {
  ColorPropertyControls,
  PropertyPaneImagePickerField,
  PropertyPaneHeadingDropdownField,
} from "@spdesigns/propertypane-controls";

export interface IExampleWebPartProps {
  textColor: string;
  selectedThemeIndex: number;
  selectedColors?: { backgroundColor: string; themePrimary: string };
  backgroundImageUrl?: string;
  headingLevel: string;
  customFontSize: number;
}

export default class ExampleWebPart extends BaseClientSideWebPart<IExampleWebPartProps> {
  // 1. Create one manager instance per web part.
  private _colorManager = new ColorPropertyControls();

  protected async onInit(): Promise<void> {
    // 2. Load the tenant's theme colors once, before the property pane can open.
    await this._colorManager.loadColors(this.context.serviceScope, -1);
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Example web part settings" },
          groups: [
            // ---- 3a. Single color picker ----
            {
              groupName: "Text color",
              groupFields: [
                this._colorManager.renderCompactColorPickerFields({
                  propertyName: "textColor",
                  label: "Text color",
                  getCurrentColor: () => this.properties.textColor,
                  onColorChange: (prop, color) => {
                    (this.properties as Record<string, string>)[prop] = color;
                    this.render();
                  },
                  onRefresh: () => this.context.propertyPane.refresh(),
                  onRender: () => this.render(),
                  // Optional: let users also type a free-form hex value.
                  additionalExpandedFields: [
                    PropertyFieldColorPicker("textColor", {
                      label: "",
                      selectedColor: this.properties.textColor || "#000",
                      onPropertyChange: (_prop, _old, newValue) => {
                        this.properties.textColor = newValue;
                        this.render();
                        this.context.propertyPane.refresh();
                      },
                      properties: this.properties,
                      style: PropertyFieldColorPickerStyle.Full,
                      key: "textColorCustomPicker",
                    }),
                  ],
                }),
              ],
            } as PropertyPaneGroup,

            // ---- 3b. Dual color (theme swatch) picker ----
            {
              groupName: "Button hover theme",
              groupFields: [
                this._colorManager.renderThemeSwatchPickerFields({
                  targetProperty: "selectedThemeIndex",
                  colorPairs: this._colorManager.colorPairs,
                  selectedIndex: this.properties.selectedThemeIndex ?? 0,
                  label: "Button hover theme",
                  onSelect: (index, pair) => {
                    this.properties.selectedThemeIndex = index;
                    this.properties.selectedColors = pair; // { backgroundColor, themePrimary }
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
              ],
            } as PropertyPaneGroup,

            // ---- 3c. Image picker ----
            {
              groupName: "Background image",
              groupFields: [
                // Preview field must come first, immediately above the file picker it drives.
                PropertyPaneImagePickerField({
                  key: "backgroundImagePreview",
                  currentImageUrl: this.properties.backgroundImageUrl,
                  onDelete: () => {
                    this.properties.backgroundImageUrl = undefined;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                PropertyFieldFilePicker("backgroundImageUrl", {
                  context: this.context as any,
                  filePickerResult: undefined,
                  onSave: (r) => {
                    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (r) => {
                    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
                  },
                  // Must match exactly — the image picker finds this button by its label text.
                  buttonLabel: "Select image",
                  properties: this.properties,
                  key: "backgroundImageUrlFilePicker",
                }),
              ],
            } as PropertyPaneGroup,

            // ---- 3d. Heading dropdown ----
            {
              groupName: "Heading",
              groupFields: [
                PropertyPaneHeadingDropdownField({
                  key: "headingLevel",
                  selected: this.properties.headingLevel || "20",
                  customSize: this.properties.customFontSize,
                  onChange: (val) => {
                    this.properties.headingLevel = val;
                    this.render();
                  },
                  onCustomSizeChange: (val) => {
                    this.properties.customFontSize = val;
                    this.render();
                  },
                }),
              ],
            } as PropertyPaneGroup,
          ],
        },
      ],
    };
  }
}
