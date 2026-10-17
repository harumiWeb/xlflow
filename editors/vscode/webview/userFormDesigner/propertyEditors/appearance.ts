import type { JSX } from "preact";
import type {
  DesignerDocument,
  PropertyGridData,
  PropertyTarget,
} from "../../../src/userFormEditor/protocol";
import type { EditorPreview } from "./types";
import { objectValue } from "./types";

// Display approximations only: neither CSS resolution nor fallback fonts are
// ever written back to an authored color/font.
const systemColors: Record<string, string> = {
  scrollBar: "Scrollbar",
  desktop: "Canvas",
  activeCaption: "ActiveCaption",
  inactiveCaption: "InactiveCaption",
  menu: "Menu",
  window: "Canvas",
  windowFrame: "WindowFrame",
  menuText: "MenuText",
  windowText: "CanvasText",
  captionText: "CaptionText",
  activeBorder: "ActiveBorder",
  inactiveBorder: "InactiveBorder",
  appWorkspace: "AppWorkspace",
  highlight: "Highlight",
  highlightText: "HighlightText",
  buttonFace: "ButtonFace",
  buttonShadow: "ButtonShadow",
  grayText: "GrayText",
  buttonText: "ButtonText",
  inactiveCaptionText: "InactiveCaptionText",
  buttonHighlight: "ButtonHighlight",
  buttonDarkShadow: "ThreeDDarkShadow",
  buttonLight: "ThreeDLightShadow",
  infoText: "InfoText",
  infoBackground: "InfoBackground",
  hotLight: "LinkText",
  gradientActiveCaption: "ActiveCaption",
  gradientInactiveCaption: "InactiveCaption",
  menuHighlight: "Highlight",
  menuBar: "Menu",
};
export function colorStyle(value: unknown): string | undefined {
  const color = objectValue(value);
  if (color.kind === "rgb" && typeof color.value === "string" && /^#[\da-f]{6}$/i.test(color.value))
    return color.value;
  if (color.kind === "system" && typeof color.value === "string") return systemColors[color.value];
  if (
    color.kind === "ole" &&
    typeof color.value === "string" &&
    /^0x[\da-f]{8}$/i.test(color.value)
  ) {
    const raw = Number.parseInt(color.value.slice(2), 16);
    if (raw <= 0xffffff)
      return `#${[raw & 255, (raw >>> 8) & 255, (raw >>> 16) & 255].map((n) => n.toString(16).padStart(2, "0")).join("")}`;
    // Unknown OLE values retain their representation; there is no RGB guess.
  }
  return undefined;
}
export function fontStyle(value: unknown): JSX.CSSProperties {
  const font = objectValue(value);
  const decoration = [
    font.underline === true ? "underline" : "",
    font.strikeThrough === true ? "line-through" : "",
  ]
    .filter(Boolean)
    .join(" ");
  return {
    fontFamily:
      typeof font.name === "string" ? `${JSON.stringify(font.name)}, sans-serif` : undefined,
    fontSize:
      typeof font.size === "number" && Number.isFinite(font.size) ? `${font.size}pt` : undefined,
    fontWeight: typeof font.bold === "boolean" ? (font.bold ? "bold" : "normal") : undefined,
    fontStyle: typeof font.italic === "boolean" ? (font.italic ? "italic" : "normal") : undefined,
    textDecoration:
      decoration || (font.underline === false && font.strikeThrough === false ? "none" : undefined),
  };
}
export function appearanceStyle(value: unknown): JSX.CSSProperties {
  const appearance = objectValue(value);
  return {
    ...fontStyle(appearance.font),
    color: colorStyle(appearance.foreColor),
    backgroundColor: colorStyle(appearance.backColor),
    textAlign:
      appearance.textAlign === 1
        ? "left"
        : appearance.textAlign === 2
          ? "center"
          : appearance.textAlign === 3
            ? "right"
            : undefined,
    borderStyle:
      appearance.borderStyle === 0 ? "none" : appearance.borderStyle === 1 ? "solid" : undefined,
  };
}
export function previewDocument(
  document: DesignerDocument,
  preview?: EditorPreview,
): DesignerDocument {
  if (!preview) return document;
  if (preview.controlId === undefined) return { ...document, ...preview.values };
  return {
    ...document,
    controls: document.controls.map((control) =>
      control.id === preview.controlId ? { ...control, ...preview.values } : control,
    ),
  };
}

export function metadataAppearanceDocument(
  document: DesignerDocument,
  grid?: PropertyGridData,
): DesignerDocument {
  if (!grid) return document;
  const appearance = (target?: PropertyTarget) =>
    Object.fromEntries(
      (target?.descriptors ?? [])
        .filter(
          (descriptor) =>
            [
              "font",
              "foreColor",
              "backColor",
              "textAlign",
              "borderStyle",
              "pictureAlignment",
              "pictureSizeMode",
            ].includes(descriptor.field) && target?.values[descriptor.field]?.present,
        )
        .map((descriptor) => [descriptor.field, target!.values[descriptor.field].value]),
    );
  // The server supplies canonical values for legacy properties. Reuse those
  // values for display instead of decoding a second legacy schema in the UI.
  return {
    ...document,
    ...appearance(grid.form),
    controls: document.controls.map((control) => ({
      ...control,
      ...appearance(grid.controls[control.id]),
    })),
  };
}
