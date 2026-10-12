import type { DesignerDocument, SpecControl } from "./protocol";
import { gridSize, stableDelta } from "./geometry";

// Shared by rendering fallback and Toolbox creation. Dimensions are points,
// matching the pure-Go compiler's built-in generation defaults.
export const controlMetadata = [
  { type: "Label", width: 72, height: 18, caption: true },
  { type: "TextBox", width: 120, height: 18 },
  { type: "ComboBox", width: 120, height: 18 },
  { type: "ListBox", width: 120, height: 72 },
  { type: "CommandButton", width: 72, height: 24, caption: true },
  { type: "CheckBox", width: 72, height: 18, caption: true },
  { type: "OptionButton", width: 72, height: 18, caption: true },
  { type: "ToggleButton", width: 72, height: 18, caption: true },
  { type: "SpinButton", width: 18, height: 36 },
  { type: "ScrollBar", width: 120, height: 18 },
  { type: "Image", width: 72, height: 72 },
  { type: "Frame", width: 144, height: 108, caption: true },
  { type: "MultiPage", width: 240, height: 180 },
  { type: "TabStrip", width: 240, height: 48 },
] as const;

export type ToolboxType = (typeof controlMetadata)[number]["type"];
export type NewControl = Pick<SpecControl, "id" | "name" | "caption" | "selectedIndex" | "tabs"> & {
  type: ToolboxType;
  left: number;
  top: number;
  width: number;
  height: number;
};

export function controlDefaults(type: string): readonly [number, number] {
  const metadata = controlMetadata.find((m) => m.type.toLowerCase() === type.toLowerCase());
  return metadata
    ? [metadata.width, metadata.height]
    : type.toLowerCase() === "page"
      ? [0, 0]
      : [72, 24];
}

export function createControl(
  document: DesignerDocument,
  type: ToolboxType,
  left: number,
  top: number,
  snapping: boolean,
  uuid: () => string = () => crypto.randomUUID(),
): NewControl | undefined {
  const metadata = controlMetadata.find((m) => m.type === type);
  if (
    !metadata ||
    !Number.isFinite(left) ||
    !Number.isFinite(top) ||
    metadata.width > document.width ||
    metadata.height > document.height
  )
    return;
  const names = new Set(
    [document.name, ...document.controls.map((c) => c.name)].map((n) => n.toLowerCase()),
  );
  let suffix = 1;
  while (names.has(`${type}${suffix}`.toLowerCase())) suffix++;
  const ids = new Set(document.controls.map((c) => c.id));
  let id: string;
  // Fail safely if the random source is broken rather than spin forever.
  for (let attempt = 0; attempt < 100; attempt++) {
    id = `control-${uuid()}`;
    if (ids.has(id)) continue;
    const position = (value: number, maximum: number) =>
      Math.max(
        0,
        Math.min(maximum, snapping ? Math.round(value / gridSize) * gridSize : stableDelta(value)),
      );
    const control: NewControl = {
      id,
      name: `${type}${suffix}`,
      type,
      left: position(left, document.width - metadata.width),
      top: position(top, document.height - metadata.height),
      width: metadata.width,
      height: metadata.height,
    };
    if ("caption" in metadata && metadata.caption) control.caption = control.name;
    if (type === "MultiPage" || type === "TabStrip") control.selectedIndex = -1;
    if (type === "TabStrip") control.tabs = [];
    return control;
  }
}

// Preview controls are canonical and flattened, including legacy nested input.
export function subtreeIds(
  controls: readonly Pick<SpecControl, "id" | "parentId">[],
  id: string,
): Set<string> {
  const result = new Set([id]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const control of controls) {
      if (control.parentId && result.has(control.parentId) && !result.has(control.id)) {
        result.add(control.id);
        changed = true;
      }
    }
  }
  return result;
}
