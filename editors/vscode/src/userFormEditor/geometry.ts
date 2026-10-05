import type { DesignerControl, DesignerDocument, GeometryOperation } from "./protocol";

export interface Geometry {
  left: number;
  top: number;
  width: number;
  height: number;
}
export interface Bounds {
  width: number;
  height: number;
}
export type ResizeHandle = "nw" | "n" | "ne" | "w" | "e" | "sw" | "s" | "se";
export const resizeHandles: ResizeHandle[] = ["nw", "n", "ne", "w", "e", "sw", "s", "se"];
export const zoomLevels = [0.5, 0.75, 1, 1.25, 1.5, 2];
export const gridSize = 8;
export const minimumSize = 1;

// Container decoration has one point-space definition for rendering and editing.
export function contentInsets(type: string) {
  if (type.toLowerCase() === "multipage") return { left: 1.5, top: 18, right: 1.5, bottom: 1.5 };
  if (type.toLowerCase() === "frame") return { left: 0.75, top: 0.75, right: 0.75, bottom: 0.75 };
  return { left: 0, top: 0, right: 0, bottom: 0 };
}
export function contentBounds(control: DesignerControl): Bounds {
  const inset = contentInsets(control.type);
  return {
    width: Math.max(0, control.width - inset.left - inset.right),
    height: Math.max(0, control.height - inset.top - inset.bottom),
  };
}
export function parentBounds(document: DesignerDocument, control: DesignerControl): Bounds {
  const parent = document.controls.find((candidate) => candidate.id === control.parentId);
  if (!parent) return { width: document.width, height: document.height };
  if (parent.type.toLowerCase() === "page") {
    const multi = document.controls.find((candidate) => candidate.id === parent.parentId);
    return multi ? contentBounds(multi) : { width: 0, height: 0 };
  }
  return contentBounds(parent);
}
export function requiredSize(document: DesignerDocument, control: DesignerControl): Bounds {
  const type = control.type.toLowerCase();
  if (type !== "frame" && type !== "multipage") return { width: minimumSize, height: minimumSize };
  const owners = new Set([control.id]);
  if (type === "multipage")
    document.controls
      .filter((c) => c.parentId === control.id && c.type.toLowerCase() === "page")
      .forEach((page) => owners.add(page.id));
  const children = document.controls.filter(
    (c) => owners.has(c.parentId ?? "") && c.type.toLowerCase() !== "page",
  );
  const inset = contentInsets(control.type);
  return {
    width:
      Math.max(minimumSize, ...children.map((c) => c.left + c.width)) + inset.left + inset.right,
    height:
      Math.max(minimumSize, ...children.map((c) => c.top + c.height)) + inset.top + inset.bottom,
  };
}
const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));
const snap = (value: number, enabled: boolean) =>
  enabled ? Math.round(value / gridSize) * gridSize : value;
export const stableDelta = (value: number) => Math.round(value * 1e6) / 1e6;

export function moveGeometry(
  start: Geometry,
  dx: number,
  dy: number,
  parent: Bounds,
  snapping = false,
): Geometry | undefined {
  if (start.width > parent.width || start.height > parent.height) return undefined;
  const moveAxis = (position: number, delta: number, limit: number) => {
    const displacement = stableDelta(delta);
    if (displacement === 0) return position >= 0 && position <= limit ? position : undefined;
    return clamp(snap(position + displacement, snapping), 0, limit);
  };
  const left = moveAxis(start.left, dx, parent.width - start.width);
  const top = moveAxis(start.top, dy, parent.height - start.height);
  if (left === undefined || top === undefined) return undefined;
  return {
    ...start,
    left,
    top,
  };
}
export function resizeGeometry(
  start: Geometry,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  parent: Bounds,
  minimum: Bounds,
  snapping = false,
): Geometry | undefined {
  const axis = (
    position: number,
    size: number,
    near: boolean,
    far: boolean,
    delta: number,
    limit: number,
    minSize: number,
  ) => {
    if ((!near && !far) || stableDelta(delta) === 0)
      return position >= 0 && size >= minimumSize && position + size <= limit
        ? { position, size }
        : undefined;
    if (minSize > limit) return undefined;
    if (near) {
      const end = position + size;
      if (end > limit || end < minSize) return undefined;
      const next = clamp(snap(position + stableDelta(delta), snapping), 0, end - minSize);
      return { position: next, size: next === position ? size : end - next };
    }
    if (position < 0 || position + minSize > limit) return undefined;
    const end = clamp(
      snap(position + size + stableDelta(delta), snapping),
      position + minSize,
      limit,
    );
    return { position, size: end === position + size ? size : end - position };
  };
  const horizontal = axis(
    start.left,
    start.width,
    handle.includes("w"),
    handle.includes("e"),
    dx,
    parent.width,
    minimum.width,
  );
  const vertical = axis(
    start.top,
    start.height,
    handle.includes("n"),
    handle.includes("s"),
    dy,
    parent.height,
    minimum.height,
  );
  if (!horizontal || !vertical) return undefined;
  return {
    left: horizontal.position,
    top: vertical.position,
    width: horizontal.size,
    height: vertical.size,
  };
}
export function geometryOperations(
  id: string,
  start: Geometry,
  end: Geometry,
): GeometryOperation[] {
  const ops: GeometryOperation[] = [];
  if (start.left !== end.left || start.top !== end.top)
    ops.push({ type: "moveControl", controlId: id, left: end.left, top: end.top });
  if (start.width !== end.width || start.height !== end.height)
    ops.push({ type: "resizeControl", controlId: id, width: end.width, height: end.height });
  return ops;
}
export function canEditGeometry(control: DesignerControl): boolean {
  return (
    control.type.toLowerCase() !== "page" &&
    [
      "label",
      "textbox",
      "commandbutton",
      "checkbox",
      "optionbutton",
      "combobox",
      "listbox",
      "togglebutton",
      "spinbutton",
      "scrollbar",
      "image",
      "frame",
      "multipage",
      "tabstrip",
    ].includes(control.type.toLowerCase())
  );
}

export function isControlVisible(document: DesignerDocument, id: string): boolean {
  let current = document.controls.find((control) => control.id === id);
  const visited = new Set<string>();
  while (current) {
    if (current.visible === false || visited.has(current.id)) return false;
    visited.add(current.id);
    const parent = document.controls.find((control) => control.id === current!.parentId);
    if (current.type.toLowerCase() === "page" && parent?.type.toLowerCase() === "multipage") {
      const pages = document.controls
        .filter((control) => control.parentId === parent.id)
        .map((control, index) => ({ control, index }))
        .sort((a, b) => (a.control.zIndex ?? a.index) - (b.control.zIndex ?? b.index))
        .map(({ control }) => control);
      if (pages[parent.selectedIndex ?? (pages.length ? 0 : -1)]?.id !== current.id) return false;
    }
    current = parent;
  }
  return visited.size > 0;
}
