# @spdesigns/propertypane-controls

Reusable SPFx PropertyPane controls extracted from `spd-product-design6`, so
they can be shared across SPFx solutions/repos instead of being copy-pasted
into each webpart's `src/shared/components/` folder.

## What's included

- **`ColorPropertyControls`** — theme-aware color picker manager.
  - `renderCompactColorPickerFields()` — **single** color: swatch + hex + toggle.
  - `renderThemeSwatchPickerFields()` — **dual** color (background + themePrimary card grid).
  - `renderPropertyPaneField()` — legacy single-color palette picker (back-compat).
  - `loadColors()` — reads `theme.spcolor`, falls back to `ThemeProvider`.
- **`PropertyPaneImagePickerField`** — image preview/placeholder field that drives
  a `PropertyFieldFilePicker` (from `@pnp/spfx-property-controls`) placed right after it,
  and hides that picker's native row.

## Install into an SPFx solution

```
npm install @spdesigns/propertypane-controls
```

Or, for local development against an unpublished change (monorepo / same checkout):
```json
// package.json
"dependencies": {
  "@spdesigns/propertypane-controls": "file:../../../shared-packages/spd-propertypane-controls"
}
```
Then `npm install` and `npm run build` inside `shared-packages/spd-propertypane-controls` once so `lib/` exists.

## Usage

### Single color picker

```ts
import { ColorPropertyControls } from "@spdesigns/propertypane-controls";

private _colorManager = new ColorPropertyControls();

protected onInit(): Promise<void> {
  return this._colorManager
    .loadColors(this.context.serviceScope, this.properties.selectedThemeIndextxt ?? -1)
    .then(() => super.onInit());
}

// in getPropertyPaneConfiguration():
...this._colorManager.renderCompactColorPickerFields({
  propertyName: "welcomeTextColor",
  label: "Text color",
  getCurrentColor: () => this.properties.welcomeTextColor,
  onColorChange: (prop, color) => {
    (this.properties as Record<string, string>)[prop] = color;
    this.render();
  },
  onRefresh: () => this.context.propertyPane.refresh(),
  onRender: () => this.render(),
}),
```

### Dual color (theme swatch) picker

```ts
...this._colorManager.renderThemeSwatchPickerFields({
  targetProperty: "selectedThemeIndex",
  colorPairs: this._colorManager.colorPairs,
  selectedIndex: this.properties.selectedThemeIndex ?? 0,
  label: "Button hover theme",
  onSelect: (index, pair) => {
    this.properties.selectedThemeIndex = index;
    this.properties.selectedColors = pair;
    this.context.propertyPane.refresh();
    this.render();
  },
}),
```

### Image picker

Place it immediately before the corresponding `PropertyFieldFilePicker`:

```ts
import { PropertyPaneImagePickerField } from "@spdesigns/propertypane-controls";
import { PropertyFieldFilePicker, PropertyFieldFilePickerOrientation } from "@pnp/spfx-property-controls";

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
  context: this.context as unknown as IPropertyFieldFilePickerHostProps["context"],
  filePickerResult: undefined,
  onSave: (r: IFilePickerResult) => {
    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
    this.context.propertyPane.refresh();
    this.render();
  },
  onChanged: (r: IFilePickerResult) => {
    this.properties.backgroundImageUrl = r.fileAbsoluteUrl;
  },
  buttonLabel: "Select image",
  properties: this.properties,
  key: "backgroundImageUrlFilePicker",
}),
```

## Migrating an existing webpart off the duplicated copy

1. Add the dependency (Option A or B above).
2. Delete the webpart's local `src/shared/components/ColorPropertyControls.ts` and/or
   `PropertyPaneImagePickerField.ts`.
3. Replace the import with `from "@spdesigns/propertypane-controls"`.
4. No API changes are needed — this package is a byte-for-byte extraction.
