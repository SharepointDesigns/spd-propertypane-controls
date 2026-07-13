import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";

export interface IPropertyPaneImagePickerFieldOptions {
  key: string;
  /** Current image absolute URL — show placeholder when undefined */
  currentImageUrl: string | undefined;
  /** Called when the user clicks the Remove button below the image */
  onDelete?: () => void;
}

/**
 * Property pane custom field that shows a clickable image preview (or placeholder).
 * - Clicking the preview opens the immediately-following PropertyFieldFilePicker.
 * - When an image is set a "Remove" button below the image clears it via `onDelete`.
 * - The PropertyFieldFilePicker field is automatically hidden.
 *
 * Place this field immediately before a `PropertyFieldFilePicker` (from
 * `@pnp/spfx-property-controls`) in the same group — this control drives that
 * picker's button by finding it in the DOM and hiding its native field row.
 */
export function PropertyPaneImagePickerField(
  opts: IPropertyPaneImagePickerFieldOptions,
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: "",
    shouldFocus: false,
    properties: {
      key: opts.key,
      onRender: (elem: HTMLElement) => _render(elem, opts),
      onDispose: (elem: HTMLElement) => {
        elem.innerHTML = "";
      },
    },
  };
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Search the property pane for the PropertyFieldFilePicker button by its label text.
 * Text-based lookup is reliable regardless of DOM depth / SPFx version.
 */
function _findPickerButton(buttonLabel = "Select image"): HTMLElement | undefined {
  // Scope to the property pane if possible
  const root: Element =
    document.querySelector(
      '[class*="PropertyPanePage"], [class*="ms-PropertyPane"]',
    ) ?? document.body;

  return Array.from(root.querySelectorAll<HTMLElement>("button")).find((btn) =>
    (btn.textContent ?? "").trim().includes(buttonLabel),
  );
}

/**
 * Given the picker button, find the SPFx field wrapper that contains it
 * by walking up until we find an element whose parent also contains `customElem`.
 */
function _findPickerFieldWrapper(
  pickerBtn: HTMLElement,
  customElem: HTMLElement,
): HTMLElement | undefined {
  // Find the "fields container" — nearest ancestor of customElem that has ≥2 children
  let customNode: HTMLElement | undefined = customElem;
  let groupContainer: HTMLElement | undefined;
  for (let i = 0; i < 6; i++) {
    const p: HTMLElement | undefined = customNode?.parentElement ?? undefined;
    if (!p) break;
    if (p.children.length >= 2) { groupContainer = p; break; }
    customNode = p;
  }
  if (!groupContainer) return undefined;

  // Walk up from pickerBtn to find a direct child of groupContainer
  let node: HTMLElement | undefined = pickerBtn;
  while (node) {
    if (node.parentElement === groupContainer) return node;
    node = node.parentElement ?? undefined;
  }
  return undefined;
}

/**
 * Poll until the PnP button appears (it renders async), then hide its field wrapper.
 * Stores the button reference in `cache` so the click handler can reuse it.
 */
function _hideWhenReady(
  customElem: HTMLElement,
  cache: { btn: HTMLElement | undefined },
): void {
  let tries = 0;
  const attempt = (): void => {
    const btn = _findPickerButton();
    if (btn) {
      cache.btn = btn;
      const wrapper = _findPickerFieldWrapper(btn, customElem);
      if (wrapper) wrapper.style.display = "none";
      return;
    }
    if (++tries < 20) setTimeout(attempt, 100);
  };
  setTimeout(attempt, 0);
}

// ── Render ────────────────────────────────────────────────────────────────────

function _render(
  elem: HTMLElement,
  opts: IPropertyPaneImagePickerFieldOptions,
): void {
  const { currentImageUrl, onDelete } = opts;
  elem.innerHTML = "";

  // Cache is local to this render call; reset on each property-pane refresh
  const cache: { btn: HTMLElement | undefined } = { btn: undefined };

  // ── Wrapper ──────────────────────────────────────────────────────────────
  const wrapper = document.createElement("div");
  wrapper.style.cssText = "display:flex;flex-direction:column;gap:6px;";

  // ── Clickable image / placeholder area ────────────────────────────────────
  const container = document.createElement("div");
  container.style.cssText = [
    "cursor:pointer",
    "border:1.5px dashed #c8c8c8",
    "border-radius:4px",
    "overflow:hidden",
    "background:#f3f2f1",
    "min-height:90px",
    "display:flex",
    "align-items:center",
    "justify-content:center",
    "position:relative",
    "transition:border-color 0.15s",
  ].join(";");

  if (currentImageUrl) {
    const img = document.createElement("img");
    img.src = currentImageUrl;
    img.style.cssText = "width:100%;max-height:120px;object-fit:cover;display:block;";
    img.alt = "Background image preview";
    container.appendChild(img);
  } else {
    const placeholder = document.createElement("div");
    placeholder.style.cssText =
      "text-align:center;padding:16px;color:#605e5c;pointer-events:none;";
    placeholder.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24"
           fill="none" stroke="#605e5c" stroke-width="1.5"
           stroke-linecap="round" stroke-linejoin="round">
        <rect x="3" y="3" width="18" height="18" rx="2"/>
        <circle cx="8.5" cy="8.5" r="1.5"/>
        <polyline points="21 15 16 10 5 21"/>
      </svg>
      <p style="margin:6px 0 0;font-size:12px;">Click to select background image</p>
    `;
    container.appendChild(placeholder);
  }

  // ── Hover overlay ─────────────────────────────────────────────────────────
  const overlay = document.createElement("div");
  overlay.style.cssText = [
    "position:absolute",
    "top:0;left:0;right:0;bottom:0",
    "background:rgba(0,0,0,0.40)",
    "display:flex",
    "align-items:center",
    "justify-content:center",
    "opacity:0",
    "transition:opacity 0.15s",
    "pointer-events:none",
  ].join(";");
  overlay.innerHTML = `<span style="color:#fff;font-size:12px;font-weight:500;">${currentImageUrl ? "Change image" : "Select image"}</span>`;
  container.appendChild(overlay);

  container.addEventListener("mouseenter", () => {
    overlay.style.opacity = "1";
    container.style.borderColor = "#0078d4";
  });
  container.addEventListener("mouseleave", () => {
    overlay.style.opacity = "0";
    container.style.borderColor = "#c8c8c8";
  });

  // Click → open file picker
  container.addEventListener("click", () => {
    // Use cached button; if stale/null try a fresh lookup
    const btn = cache.btn ?? _findPickerButton();
    if (btn) {
      cache.btn = btn;
      // Ensure the field wrapper stays hidden (may have been reset by pane refresh)
      const fw = _findPickerFieldWrapper(btn, elem);
      if (fw) fw.style.display = "none";
      btn.click();
    }
  });

  wrapper.appendChild(container);

  // ── Remove button (only when an image is set) ────────────────────────────
  if (currentImageUrl && onDelete) {
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.style.cssText = [
      "display:inline-flex",
      "align-items:center",
      "gap:5px",
      "align-self:flex-start",
      "background:none",
      "border:none",
      "color:#a4262c",
      "font-size:12px",
      "cursor:pointer",
      "padding:2px 0",
      "line-height:1.4",
    ].join(";");
    removeBtn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
           fill="none" stroke="currentColor" stroke-width="2"
           stroke-linecap="round" stroke-linejoin="round">
        <polyline points="3 6 5 6 21 6"/>
        <path d="M19 6l-1 14H6L5 6"/>
        <path d="M10 11v6M14 11v6"/>
        <path d="M9 6V4h6v2"/>
      </svg>
      Remove
    `;
    removeBtn.addEventListener("click", () => onDelete());
    wrapper.appendChild(removeBtn);
  }

  elem.appendChild(wrapper);

  // Hide the PropertyFieldFilePicker field with retry (PnP renders async)
  _hideWhenReady(elem, cache);
}
