# @spdesigns/propertypane-controls

Reusable SPFx PropertyPane controls — a single/dual color picker and an image
picker — shared across SPFx webparts instead of copy-pasted into each one.

![Property pane controls hero](assets/hero%20image.png)

> A complete, runnable example wiring up all three controls in one web part
> lives in [`example-usage.ts`](example-usage.ts).

## What's included

- **`ColorPropertyControls`** — theme-aware color picker manager.
  - `renderCompactColorPickerFields()` — single color: swatch + hex + toggle.
  - `renderThemeSwatchPickerFields()` — dual color (background + accent card grid).
  - `loadColors()` — reads the tenant theme colors so swatches always match the site.
- **`PropertyPaneImagePickerField`** — image preview/placeholder field that
  drives a `PropertyFieldFilePicker` (from `@pnp/spfx-property-controls`).
- **`PropertyPaneHeadingDropdownField`** — heading-level dropdown (H1–H4/Normal)
  with a custom font-size slider.

## Install

```
npm install @spdesigns/propertypane-controls
```

For local development against an unpublished change:
```json
// package.json
"dependencies": {
  "@spdesigns/propertypane-controls": "file:../../../shared-packages/spd-propertypane-controls"
}
```
Then run `npm install` and `npm run build` inside the package once so `lib/` exists.

## Setup (once per webpart)

```ts
import { ColorPropertyControls } from "@spdesigns/propertypane-controls";

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  private _colorManager = new ColorPropertyControls();

  protected async onInit(): Promise<void> {
    await this._colorManager.loadColors(this.context.serviceScope, -1);
    return super.onInit();
  }
```

That's it — `_colorManager` now has the tenant's swatches and color pairs
loaded, ready to use in `getPropertyPaneConfiguration()`.

## 1. Single color picker

One color property (e.g. text color, background color). Add one field to your
property pane group per color you want configurable:

![Single color picker](assets/single%20color%20picker.png)

See it wired up in [`example-usage.ts`](example-usage.ts#L40-L70) (the "Text color" group).

```ts
this._colorManager.renderCompactColorPickerFields({
  propertyName: "textColor",           // matches this.properties.textColor
  label: "Text color",
  getCurrentColor: () => this.properties.textColor,
  onColorChange: (prop, color) => {
    (this.properties as Record<string, string>)[prop] = color;
    this.render();
  },
  onRefresh: () => this.context.propertyPane.refresh(),
  onRender: () => this.render(),
}),
```

Add as many of these as you have color properties — each call is independent.
Keep the property itself a plain `string` (not an object) so nothing else
needs to change to use it.

To also let users enter a free-form hex value, pass a
`PropertyFieldColorPicker` (from `@pnp/spfx-property-controls`) via
`additionalExpandedFields` — it renders below the swatch grid whenever the
picker is open:

```ts
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls";

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
```

> **Optional — object-shaped color properties.** If your color property is an
> object (e.g. `selectedColor: { themePrimary: string }`) instead of a flat
> string, `PropertyFieldColorPicker` will still write a raw string to
> `targetProperty` and overwrite the whole object. In that case, add an
> `onPropertyPaneFieldChanged` override to repair it — this is not required
> for flat string properties:
> ```ts
> protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
>   if (propertyPath === "selectedColor" && typeof newValue === "string") {
>     this.properties.selectedColor = { themePrimary: newValue };
>   }
> }
> ```

## 2. Dual color (theme swatch) picker

For a background + accent color pair (e.g. button hover theme, gradient
card), backed by a single index property:

![Dual color picker](assets/dual%20color%20picker.png)

See it wired up in [`example-usage.ts`](example-usage.ts#L73-L90) (the "Button hover theme" group).

```ts
this._colorManager.renderThemeSwatchPickerFields({
  targetProperty: "selectedThemeIndex",
  colorPairs: this._colorManager.colorPairs,
  selectedIndex: this.properties.selectedThemeIndex ?? 0,
  label: "Button hover theme",
  onSelect: (index, pair) => {
    this.properties.selectedThemeIndex = index;
    this.properties.selectedColors = pair;   // { backgroundColor, themePrimary }
    this.context.propertyPane.refresh();
    this.render();
  },
}),
```

## 3. Image picker

Place `PropertyPaneImagePickerField` immediately before the
`PropertyFieldFilePicker` it drives, and keep `buttonLabel` set to
`"Select image"` — the image picker finds the file picker's button by that
label text, so it must match exactly.

![Image picker](assets/image%20picker.png)

See it wired up in [`example-usage.ts`](example-usage.ts#L92-L123) (the "Background image" group).

```ts
import { PropertyPaneImagePickerField } from "@spdesigns/propertypane-controls";
import { PropertyFieldFilePicker } from "@pnp/spfx-property-controls";

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
  context: this.context,
  filePickerResult: undefined,
  onSave: (r) => {
    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
    this.context.propertyPane.refresh();
    this.render();
  },
  onChanged: (r) => {
    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
  },
  buttonLabel: "Select image",
  properties: this.properties,
  key: "backgroundImageUrlFilePicker",
}),
```

## 4. Heading dropdown

A heading-level dropdown (H1–H4/Normal) that reveals a font-size slider when
"Custom Font Size" is selected.

```ts
import { PropertyPaneHeadingDropdownField } from "@spdesigns/propertypane-controls";

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
```

`selected` is the option key — `"32"`/`"28"`/`"24"`/`"20"` (Heading 1–4),
`"18"` (Normal), or `"custom"`. Apply the resulting `headingLevel` /
`customFontSize` to your rendered heading however your webpart already maps
font sizes (e.g. a CSS custom property).

## Migrating an existing webpart off a copy-pasted version

1. Add the dependency (see Install above).
2. Delete the webpart's local `ColorPropertyControls.ts` / `PropertyPaneImagePickerField.ts` /
   `HeadingDropdown.tsx` / `HeadingDropdownPropertyPane.ts`.
3. Replace the import with `from "@spdesigns/propertypane-controls"`.
4. No API changes needed — this package is a byte-for-byte extraction.

## Tips

- Keep color properties as flat strings (`color: string`), not nested objects
  — it's less to get wrong when pairing with a free-form hex picker.
- If you restore a saved color/index in `onInit()`, only assign it when the
  property is still unset (`if (!this.properties.textColor) ...`) so you
  don't overwrite a color the user already picked every time the pane opens.
- `onDelete` for the image picker should just clear the property if the file
  belongs to the user's library; only delete the underlying file too if your
  webpart owns/uploaded it.

## Changelog

### 1.1.0
- Added `PropertyPaneHeadingDropdownField` (heading-level dropdown with a
  custom font-size slider) and the underlying `HeadingDropdown` React component.

### 1.0.2
- `renderThemeSwatchPickerFields()` dual picker now cross-joins every unique
  `themePrimary` with every unique `backgroundColor` found across the site's
  secondary color palettes, instead of only the as-authored pairs.

### 1.0.1
- `renderCompactColorPickerFields()` single picker now also includes colors
  from the site's secondary/branding palettes (`secondaryColors > light`),
  not just the primary theme palette — matching what the `ThemeProvider`
  fallback path already did.

### 1.0.0
- Initial release: `ColorPropertyControls` (single + dual color pickers) and
  `PropertyPaneImagePickerField`.
