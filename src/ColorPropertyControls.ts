import { IReadonlyTheme, ThemeProvider } from "@microsoft/sp-component-base";
import { IPropertyPaneField } from "@microsoft/sp-property-pane";
import { ServiceKey } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";

/**
 * Structural subset of ServiceScope. Consumers may have their own separately
 * installed copy of @microsoft/sp-core-library; typing this as the concrete
 * ServiceScope class would fail cross-package type identity checks (private
 * members like _registrations differ by physical install, not just version).
 */
export interface IServiceScopeLike {
  consume<T>(serviceKey: ServiceKey<T>): T;
}

export interface IColorPropertySwatch {
  color: string;
  label: string;
}

export interface IColorPropertyChangeEvent {
  index: number;
  color: string;
}

export interface IColorPropertyState {
  selectedColorstxt?: { themePrimary: string };
}

/**
 * Options for the compact inline color picker fields.
 * Pass one of these to renderCompactColorPickerFields() for each single-color property.
 */
export interface ICompactColorPickerFieldOptions {
  /** Property name used as key and CSS class suffix (e.g. "welcomeTextColor") */
  propertyName: string;
  /** Human-readable label shown above the compact row */
  label: string;
  /** Called at field render time to get the live color value */
  getCurrentColor: () => string | undefined;
  /** Called when the user picks a swatch; update your property and call render() yourself */
  onColorChange: (propertyName: string, color: string) => void;
  /** Called to refresh the property pane (this.context.propertyPane.refresh()) */
  onRefresh: () => void;
  /** Called to re-render the web part (this.render()) */
  onRender: () => void;
  /**
   * Optional extra fields appended after the swatch grid when the picker is open.
   * Use this to inject a PnP PropertyFieldColorPicker for free-form hex input.
   */
  additionalExpandedFields?: IPropertyPaneField<unknown>[];
}

/**
 * Options for the dual-color theme swatch picker field.
 * Pass this to renderThemeSwatchPickerFields() for background + themePrimary pair selection.
 */
export interface IThemeSwatchPickerFieldOptions {
  /** Property path used for targetProperty (e.g. "selectedThemeIndex") */
  targetProperty: string;
  /** Color pairs from theme's secondaryColors.light */
  colorPairs: Array<{ backgroundColor: string; themePrimary: string }>;
  selectedIndex: number;
  label?: string;
  onSelect: (
    index: number,
    pair: { backgroundColor: string; themePrimary: string },
  ) => void;
}

const PALETTE_KEYS = [
  "themeDarker",
  "themeDark",
  "themeDarkAlt",
  "themePrimary",
  "themeSecondary",
  "themeTertiary",
  "themeLight",
  "themeLighter",
  "themeLighterAlt",
  "black",
  "neutralDark",
  "neutralPrimary",
  "neutralPrimaryAlt",
  "neutralSecondary",
  "neutralTertiary",
  "neutralTertiaryAlt",
  "neutralLight",
  "neutralLighter",
  "neutralLighterAlt",
  "neutralQuaternaryAlt",
  "neutralQuaternary",
  "accent",
];

const FALLBACK_COLORS: IColorPropertySwatch[] = [
  { color: "#ffb900", label: "Yellow" },
  { color: "#fff100", label: "Light Yellow" },
  { color: "#d83b01", label: "Orange" },
  { color: "#e81123", label: "Red" },
  { color: "#a80000", label: "Dark Red" },
  { color: "#5c005c", label: "Dark Magenta" },
  { color: "#e3008c", label: "Light Magenta" },
  { color: "#5c2d91", label: "Purple" },
  { color: "#0078d4", label: "Blue" },
  { color: "#00bcf2", label: "Light Blue" },
  { color: "#008272", label: "Teal" },
  { color: "#107c10", label: "Green" },
  { color: "#bad80a", label: "Light Green" },
  { color: "#eaeaea", label: "Gray" },
  { color: "#333333", label: "Neutral" },
  { color: "#000000", label: "Black" },
  { color: "rgba(102, 102, 102, 0.5)", label: "Half Gray" },
];

function _contrastText(hex: string): string {
  const c = hex.replace("#", "");
  if (c.length < 6) return "#000000";
  const r = parseInt(c.slice(0, 2), 16);
  const g = parseInt(c.slice(2, 4), 16);
  const b = parseInt(c.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.55
    ? "#000000"
    : "#ffffff";
}

/**
 * Manages color swatches and color picker property pane UI.
 * Instantiate once per WebPart and call loadColors() in onInit().
 *
 * - renderCompactColorPickerFields() — single-color swatch + hex input + toggle
 * - renderThemeSwatchPickerFields()  — dual-color (backgroundColor + themePrimary) grid cards
 * - renderPropertyPaneField()        — legacy single-color palette picker (kept for compatibility)
 */
export class ColorPropertyControls {
  private _themeVariant: IReadonlyTheme | undefined;
  private _swatches: IColorPropertySwatch[] = [];
  private _colorPairs: Array<{
    backgroundColor: string;
    themePrimary: string;
  }> = [];
  /** State for the legacy renderPropertyPaneField() toggle */
  private _isPickerOpen: boolean = false;
  /** Tracks which compact picker field is currently expanded (null = all closed) */
  private _openPickerField: string | null = null;

  get swatches(): IColorPropertySwatch[] {
    return this._swatches;
  }

  get colorPairs(): Array<{ backgroundColor: string; themePrimary: string }> {
    return this._colorPairs;
  }

  /**
   * Builds the color swatch list and dual-color pairs from the site's theme.spcolor XML.
   * Falls back to ThemeProvider if the XML fetch fails.
   * Dual-color pairs are stored in this.colorPairs for use with renderThemeSwatchPickerFields().
   * Returns the resolved color for the saved index so the caller can update its property.
   */
  async loadColors(
    serviceScope: IServiceScopeLike,
    savedIndex: number,
  ): Promise<IColorPropertyState> {
    const pageContext = serviceScope.consume(PageContext.serviceKey);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const themedCssFolderUrl = (pageContext as any).legacyPageContext
      ?.themedCssFolderUrl as string | undefined;
    const spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);

    if (themedCssFolderUrl && spHttpClient) {
      const response = await spHttpClient.get(
        `${themedCssFolderUrl}/theme.spcolor`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "text/xml, application/xml" } },
      );

      if (response.ok) {
        const xmlText = await response.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, "application/xml");

        const parseHex = (raw: string): string =>
          raw.length >= 8 ? `#${raw.slice(2)}` : `#${raw}`;

        const getColor = (name: string): string | undefined => {
          const nodes = xmlDoc.querySelectorAll("colorPalette > color");
          for (const node of Array.from(nodes)) {
            if (node.getAttribute("name") === name) {
              const raw = node.getAttribute("value") ?? "";
              return parseHex(raw);
            }
          }
          return undefined;
        };

        const seen = new Set<string>();
        this._swatches = PALETTE_KEYS.map((k) => ({
          color: getColor(k),
          label: k,
        })).filter((s): s is IColorPropertySwatch => {
          if (!s.color || seen.has(s.color)) return false;
          seen.add(s.color);
          return true;
        });
        if (this._swatches.length === 0) {
          this._swatches = [...FALLBACK_COLORS];
        }

        const secondaryPaletteNodes = xmlDoc.querySelectorAll(
          "secondaryColors > light > colorPalette",
        );
        const palettesFromXml = Array.from(secondaryPaletteNodes).map(
          (palette) => {
            const entry: Record<string, string> = {};
            palette.querySelectorAll("color").forEach((c) => {
              const n = c.getAttribute("name") ?? "";
              const v = c.getAttribute("value") ?? "";
              entry[n] = parseHex(v);
            });
            return entry;
          },
        );

        const white = "#ffffff";
        const fallbackPairs = (
          [
            [white, getColor("themePrimary")],
            [white, getColor("accent")],
            [getColor("themePrimary"), white],
            [white, getColor("themeSecondary")],
            [getColor("themePrimary"), getColor("neutralLighter") ?? "#f3f2f1"],
            [white, getColor("themeTertiary")],
            [white, getColor("themeLight")],
            [white, getColor("HeaderBackground")],
          ] as Array<[string, string | undefined]>
        )
          .filter((pair): pair is [string, string] => !!(pair[0] && pair[1]))
          .map(([themePrimary, backgroundColor]) => ({
            themePrimary,
            backgroundColor,
          }));

        this._colorPairs = (
          palettesFromXml.length > 0 ? palettesFromXml : fallbackPairs
        ).map((e) => ({
          themePrimary: e.themePrimary ?? white,
          backgroundColor: e.backgroundColor ?? white,
        }));

        if (savedIndex >= 0 && this._swatches[savedIndex]) {
          return {
            selectedColorstxt: {
              themePrimary: this._swatches[savedIndex].color,
            },
          };
        }
        return {};
      }
    }

    // Fallback: ThemeProvider
    const themeProvider = serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = themeProvider?.tryGetTheme();

    if (this._themeVariant) {
      const themeJson = JSON.parse(JSON.stringify(this._themeVariant));
      const secondaryLight: Array<{
        themePrimary: string;
        backgroundColor?: string;
      }> = themeJson.secondaryColors?.light || [];
      const secondarySwatches: IColorPropertySwatch[] = secondaryLight.map(
        (c, i) => ({
          color: c.themePrimary,
          label: `brand-${i + 1}`,
        }),
      );
      const palette: Record<string, string> = themeJson.palette || {};
      const paletteSwatches = PALETTE_KEYS.filter((k) => !!palette[k]).map(
        (k) => ({ color: palette[k], label: k }),
      );
      const seen = new Set<string>();
      this._swatches = [...secondarySwatches, ...paletteSwatches].filter(
        (s) => {
          if (seen.has(s.color)) return false;
          seen.add(s.color);
          return true;
        },
      );

      const white = palette.white || "#ffffff";
      const pairsFromTheme = secondaryLight
        .filter((c) => !!c.themePrimary && !!c.backgroundColor)
        .map((c) => ({
          themePrimary: c.themePrimary,
          backgroundColor: c.backgroundColor as string,
        }));

      const fallbackPairs = (
        [
          [white, palette.themePrimary],
          [white, palette.accent],
          [palette.themePrimary, white],
          [white, palette.themeSecondary],
          [palette.themePrimary, palette.neutralLighter ?? "#f3f2f1"],
          [white, palette.themeTertiary],
          [white, palette.themeLight],
        ] as Array<[string, string | undefined]>
      )
        .filter((pair): pair is [string, string] => !!(pair[0] && pair[1]))
        .map(([themePrimary, backgroundColor]) => ({
          themePrimary,
          backgroundColor,
        }));

      this._colorPairs =
        pairsFromTheme.length > 0 ? pairsFromTheme : fallbackPairs;
    } else {
      this._swatches = [...FALLBACK_COLORS];
    }

    if (savedIndex >= 0 && this._swatches[savedIndex]) {
      return {
        selectedColorstxt: { themePrimary: this._swatches[savedIndex].color },
      };
    }
    return {};
  }

  /**
   * Resolves the color for a given swatch index.
   * Call inside onPropertyPaneFieldChanged when propertyPath === "selectedThemeIndextxt".
   */
  resolveColor(index: number): IColorPropertyChangeEvent | undefined {
    if (index >= 0 && this._swatches[index]) {
      return { index, color: this._swatches[index].color };
    }
    return undefined;
  }

  /**
   * Returns IPropertyPaneField[] for a compact inline single-color picker.
   *
   * Always renders: a labeled preview swatch + hex display + toggle button.
   * When open also renders: theme swatch grid + any additionalExpandedFields.
   *
   * Usage in getPropertyPaneConfiguration():
   * ```ts
   * ...this._colorManager.renderCompactColorPickerFields({
   *   propertyName: "welcomeTextColor",
   *   label: "Text color",
   *   getCurrentColor: () => this.properties.welcomeTextColor,
   *   onColorChange: (prop, color) => {
   *     (this.properties as Record<string, string>)[prop] = color;
   *     this.render();
   *   },
   *   onRefresh: () => this.context.propertyPane.refresh(),
   *   onRender: () => this.render(),
   * }),
   * ```
   */
  renderCompactColorPickerFields(
    options: ICompactColorPickerFieldOptions,
  ): IPropertyPaneField<unknown>[] {
    const {
      propertyName,
      label,
      getCurrentColor,
      onColorChange,
      onRefresh,
      onRender,
      additionalExpandedFields = [],
    } = options;
    const isOpen = this._openPickerField === propertyName;
    const p = propertyName;

    const compactField: IPropertyPaneField<unknown> = {
      type: 1,
      targetProperty: propertyName,
      properties: {
        key: `compactColor_${p}`,
        onRender: (elem: HTMLElement) => {
          const color = getCurrentColor() || "#ffffff";
          elem.innerHTML = `
            <style>
              .spd-cr-${p}{display:flex;flex-direction:column;margin-bottom:8px;}
              .spd-cl-${p}{font-size:14px;font-weight:600;color:#323130;font-family:inherit;margin-bottom:6px;display:block;}
              .spd-cc-${p}{display:flex;align-items:center;gap:6px;}
              .spd-csw-${p}{
                width:32px;height:32px;border-radius:4px;border:1px solid #c8c6c4;
                cursor:pointer;flex-shrink:0;
                background-image:linear-gradient(45deg,#ccc 25%,transparent 25%),
                  linear-gradient(-45deg,#ccc 25%,transparent 25%),
                  linear-gradient(45deg,transparent 75%,#ccc 75%),
                  linear-gradient(-45deg,transparent 75%,#ccc 75%);
                background-size:8px 8px;
                background-position:0 0,0 4px,4px -4px,-4px 0;
                overflow:hidden;
              }
              .spd-csi-${p}{width:100%;height:100%;background-color:${color};}
              .spd-ch-${p}{
                flex:1;height:32px;border:1px solid #c8c6c4;border-radius:4px;
                padding:0 8px;font-size:13px;font-family:monospace;
                color:#323130;background:#faf9f8;outline:none;box-sizing:border-box;
              }
              .spd-ch-${p}:focus{border-color:#0078d4;}
              .spd-cb-${p}{
                width:32px;height:32px;border:1px solid #c8c6c4;border-radius:4px;
                background:#faf9f8;cursor:pointer;display:flex;align-items:center;
                justify-content:center;color:#605e5c;flex-shrink:0;padding:0;
              }
              .spd-cb-${p}:hover{background:#edebe9;color:#323130;}
            </style>
            <div class="spd-cr-${p}">
              <span class="spd-cl-${p}">${label}</span>
              <div class="spd-cc-${p}">
                <div class="spd-csw-${p} js-toggle-${p}">
                  <div class="spd-csi-${p}"></div>
                </div>
                <input class="spd-ch-${p}" type="text" value="${color}" maxlength="9" spellcheck="false" readonly />
                <button class="spd-cb-${p} js-toggle-${p}" title="${isOpen ? "Close picker" : "Pick color"}">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <circle cx="8" cy="2.5" r="1.5" fill="currentColor"/>
                    <circle cx="8" cy="8" r="1.5" fill="currentColor"/>
                    <circle cx="8" cy="13.5" r="1.5" fill="currentColor"/>
                  </svg>
                </button>
              </div>
            </div>
          `;

          elem
            .querySelectorAll<HTMLElement>(`.js-toggle-${p}`)
            .forEach((el) => {
              el.addEventListener("click", () => {
                this._openPickerField = this._openPickerField === p ? null : p;
                onRefresh();
              });
            });
        },
        onDispose: (): void => undefined,
      },
    } as IPropertyPaneField<unknown>;

    if (!isOpen) return [compactField];

    const labelStyle = `font-size:14px;font-weight:600;color:#323130;font-family:inherit;display:block;margin-bottom:6px;`;

    const swatchGridField: IPropertyPaneField<unknown> = {
      type: 1,
      targetProperty: `${p}_swatches`,
      properties: {
        key: `swatches_${p}`,
        onRender: (elem: HTMLElement) => {
          const currentColor = getCurrentColor();
          elem.innerHTML = `
            <style>
              .spd-sg-wrap-${p}{margin-bottom:8px;}
              .spd-sg-grid-${p}{display:flex;flex-wrap:wrap;gap:4px;}
              .spd-sg-btn-${p}{
                width:24px;height:24px;border-radius:50%;
                cursor:pointer;padding:0;flex-shrink:0;
                box-sizing:border-box;border:2px solid #c8c6c4;
                transition:transform 0.1s;
              }
              .spd-sg-btn-${p}:hover{transform:scale(1.15);}
              .spd-sg-btn-${p}.active{border:3px solid #000;}
            </style>
            <div class="spd-sg-wrap-${p}">
              <span style="${labelStyle}">Theme colors</span>
              <div class="spd-sg-grid-${p}">
                ${this._swatches
                  .map(
                    (s, i) => `
                  <div
                    class="spd-sg-btn-${p}${currentColor?.toLowerCase() === s.color?.toLowerCase() ? " active" : ""}"
                    data-color="${s.color}"
                    data-index="${i}"
                    title="${s.label} (${s.color})"
                    style="background-color:${s.color};"
                  ></div>
                `,
                  )
                  .join("")}
              </div>
            </div>
          `;

          elem
            .querySelectorAll<HTMLDivElement>(`.spd-sg-btn-${p}`)
            .forEach((btn) => {
              btn.addEventListener("click", () => {
                const newColor = btn.getAttribute("data-color") || "#ffffff";
                onColorChange(propertyName, newColor);
                onRefresh();
                onRender();
              });
            });
        },
        onDispose: (): void => undefined,
      },
    } as IPropertyPaneField<unknown>;

    const customLabelField: IPropertyPaneField<unknown> = {
      type: 1,
      targetProperty: `${p}_customLabel`,
      properties: {
        key: `customLabel_${p}`,
        onRender: (elem: HTMLElement) => {
          elem.innerHTML = `
            <span style="${labelStyle}margin-top:4px;">Custom color</span>
            <style>
              .ms-ColorPicker-table td:nth-child(2),
              .ms-ColorPicker-table td:nth-child(3),
              .ms-ColorPicker-table td:nth-child(4),
              .ms-ColorPicker-table th:nth-child(2),
              .ms-ColorPicker-table th:nth-child(3),
              .ms-ColorPicker-table th:nth-child(4) {
                display: none !important;
              }
            </style>
          `;
        },
        onDispose: (): void => undefined,
      },
    } as IPropertyPaneField<unknown>;

    return [
      compactField,
      swatchGridField,
      customLabelField,
      ...additionalExpandedFields,
    ];
  }

  /**
   * Returns IPropertyPaneField[] for a dual-color theme swatch picker.
   * Renders a grid of cards each showing a backgroundColor block and a themePrimary block,
   * matching the SharePoint theme picker visual style.
   *
   * Use this for gradient / background+primary color selection (e.g. button hover theme,
   * icon box theme, dashboard gradient).
   *
   * Usage in getPropertyPaneConfiguration():
   * ```ts
   * ...this._colorManager.renderThemeSwatchPickerFields({
   *   targetProperty: "selectedThemeIndex",
   *   colorPairs: colorPairs,
   *   selectedIndex: this.properties.selectedThemeIndex ?? 0,
   *   label: "Button hover theme",
   *   onSelect: (index, pair) => {
   *     this.properties.selectedThemeIndex = index;
   *     this.properties.selectedColors = pair;
   *     this.context.propertyPane.refresh();
   *     this.render();
   *   },
   * }),
   * ```
   */
  renderThemeSwatchPickerFields(
    options: IThemeSwatchPickerFieldOptions,
  ): IPropertyPaneField<unknown>[] {
    const { targetProperty, colorPairs, selectedIndex, label, onSelect } =
      options;
    if (!colorPairs || colorPairs.length === 0) return [];

    const uid = `tsp_${targetProperty}`;

    return [
      {
        type: 1,
        targetProperty,
        properties: {
          key: `themeSwatchPicker_${targetProperty}`,
          onRender: (elem: HTMLElement) => {
            const swatchesHTML = colorPairs
              .map((pair, i) => {
                const isSelected = i === selectedIndex;
                const bgText = _contrastText(pair.backgroundColor);
                const primaryText = _contrastText(pair.themePrimary);
                return `
              <button
                class="${uid}-card${isSelected ? ` ${uid}-card--selected` : ""}"
                data-index="${i}"
                title="Background: ${pair.backgroundColor} / Primary: ${pair.themePrimary}"
                aria-pressed="${isSelected}"
              >
                <span class="${uid}-block ${uid}-block--large" style="background:${pair.backgroundColor};color:${bgText};">Aa</span>
                <span class="${uid}-block ${uid}-block--small" style="background:${pair.themePrimary};color:${primaryText};">Aa</span>
              </button>`;
              })
              .join("");

            elem.innerHTML = `
            <style>
              .${uid}-wrap{margin-bottom:8px;}
              .${uid}-label{font-size:14px;font-weight:600;color:#323130;font-family:inherit;margin-bottom:8px;display:block;}
              .${uid}-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;}
              .${uid}-card{
                display:flex;flex-direction:column;gap:4px;padding:5px;
                background:#ffffff;border:2px solid #c8c6c4;border-radius:8px;
                cursor:pointer;box-sizing:border-box;transition:border-color 0.12s,box-shadow 0.12s;
              }
              .${uid}-card:hover{border-color:#0078d4;}
              .${uid}-card--selected{border:2.5px solid #0078d4 !important;box-shadow:0 0 0 2px #c7e0f4;}
              .${uid}-block{
                display:flex;align-items:center;justify-content:center;
                border-radius:4px;font-family:inherit;font-weight:600;line-height:1;
                width:100%;box-sizing:border-box;
              }
              .${uid}-block--large{height:34px;font-size:11px;}
              .${uid}-block--small{height:20px;font-size:9px;}
            </style>
            <div class="${uid}-wrap">
              ${label ? `<span class="${uid}-label">${label}</span>` : ""}
              <div class="${uid}-grid">${swatchesHTML}</div>
            </div>`;

            elem
              .querySelectorAll<HTMLButtonElement>(`.${uid}-card`)
              .forEach((btn) => {
                btn.addEventListener("click", () => {
                  const idx = parseInt(
                    btn.getAttribute("data-index") || "0",
                    10,
                  );
                  onSelect(idx, colorPairs[idx]);
                });
              });
          },
          onDispose: (): void => undefined,
        },
      } as IPropertyPaneField<unknown>,
    ];
  }

  /**
   * Returns the IPropertyPaneField array for the legacy palette swatch picker.
   * Existing webparts that use selectedThemeIndextxt/selectedColorstxt can continue using this.
   * New webparts should prefer renderCompactColorPickerFields().
   */
  renderPropertyPaneField(options: {
    label: string;
    selectedIndex: number;
    selectedColor?: { themePrimary: string };
    onColorSelect: (event: IColorPropertyChangeEvent) => void;
    onRefresh: () => void;
    onRender: () => void;
  }): IPropertyPaneField<unknown>[] {
    if (this._swatches.length === 0) return [];

    return [
      {
        type: 1,
        targetProperty: "titleColorSwatches",
        properties: {
          key: "titleColorSwatches",
          onRender: (elem: HTMLElement) => {
            elem.innerHTML = "";

            const labelEl = document.createElement("label");
            labelEl.textContent = options.label;
            labelEl.style.cssText =
              "display:block;font-weight:600;font-size:14px;margin-top:4px;margin-bottom:6px;";
            elem.appendChild(labelEl);

            const header = document.createElement("div");
            header.style.cssText = "display:flex;align-items:center;gap:6px;";

            const preview = document.createElement("div");
            const themePalette =
              (
                this._themeVariant as unknown as {
                  palette: Record<string, string>;
                }
              )?.palette || {};
            const defaultColor = themePalette.black || "#000000";
            const selectedColor =
              options.selectedColor?.themePrimary || defaultColor;
            preview.style.cssText =
              "flex:1;height:28px;border:1px solid #c8c6c4;border-radius:2px;background-color:" +
              selectedColor +
              ";box-sizing:border-box;";
            header.appendChild(preview);

            const paletteBtn = document.createElement("span");
            paletteBtn.innerHTML = "&#x1F3A8;";
            paletteBtn.title = this._isPickerOpen
              ? "Close picker"
              : "Pick color";
            paletteBtn.style.cssText =
              "cursor:pointer;font-size:18px;line-height:1;flex-shrink:0;";
            paletteBtn.onclick = () => {
              this._isPickerOpen = !this._isPickerOpen;
              options.onRefresh();
            };
            header.appendChild(paletteBtn);
            elem.appendChild(header);

            if (!this._isPickerOpen) return;

            const grid = document.createElement("div");
            grid.style.cssText =
              "display:flex;flex-wrap:wrap;gap:4px;margin-top:8px;";
            this._swatches.forEach((swatch, index) => {
              const el = document.createElement("div");
              const isSelected = index === options.selectedIndex;
              el.style.cssText =
                "width:24px;height:24px;border-radius:50%;background-color:" +
                swatch.color +
                ";cursor:pointer;border:" +
                (isSelected ? "3px solid #000" : "2px solid #c8c6c4") +
                ";box-sizing:border-box;transition:transform 0.1s;";
              el.title = swatch.label + " (" + swatch.color + ")";
              el.onmouseenter = () => (el.style.transform = "scale(1.15)");
              el.onmouseleave = () => (el.style.transform = "scale(1)");
              el.onclick = () => {
                options.onColorSelect({ index, color: swatch.color });
                options.onRefresh();
                options.onRender();
              };
              grid.appendChild(el);
            });
            elem.appendChild(grid);
          },
          onDispose: (): void => undefined,
        },
      } as IPropertyPaneField<unknown>,
    ];
  }
}
