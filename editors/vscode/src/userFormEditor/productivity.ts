import type {
  DesignerControl,
  DesignerDocument,
  ProductivityOperation,
  PropertyGridData,
} from "./protocol";
import {
  gridSize,
  isControlVisible,
  orderedChildren,
  parentBounds,
  stableDelta,
  worldGeometry,
} from "./geometry";
import { controlMetadata } from "./toolbox";

export interface SelectionRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const isPage = (control: DesignerControl) => control.type.toLowerCase() === "page";

export function isProductivityControl(control: DesignerControl): boolean {
  const metadata = controlMetadata.find(
    (candidate) => candidate.type.toLowerCase() === control.type.toLowerCase(),
  );
  return (
    metadata !== undefined &&
    (!control.progId || control.progId.toLowerCase() === `forms.${metadata.type.toLowerCase()}.1`)
  );
}

export function siblingSelection(
  document: DesignerDocument,
  ids: readonly string[],
): DesignerControl[] {
  const unique = [...new Set(ids)];
  const controls = unique.map((id) => document.controls.find((control) => control.id === id));
  if (
    controls.length === 0 ||
    controls.some((control) => !control || isPage(control) || !isProductivityControl(control)) ||
    new Set(controls.map((control) => control!.parentId ?? "")).size !== 1
  )
    return [];
  return controls as DesignerControl[];
}

export function updateSelection(
  document: DesignerDocument,
  selectedIds: readonly string[],
  clickedId: string | undefined,
  toggle = false,
): string[] {
  if (!clickedId) return [];
  const clicked = document.controls.find((control) => control.id === clickedId);
  if (!clicked) return [];
  if (isPage(clicked) || !toggle) return [clickedId];
  if (!isProductivityControl(clicked)) return [clickedId];

  const current = siblingSelection(document, selectedIds);
  if (current.length && (current[0].parentId ?? "") !== (clicked.parentId ?? ""))
    return [clickedId];
  const next = current.map((control) => control.id);
  const index = next.indexOf(clickedId);
  if (index >= 0) next.splice(index, 1);
  else next.push(clickedId);
  return next;
}

export function promoteSelection(
  document: DesignerDocument,
  selectedIds: readonly string[],
  clickedId: string,
): string[] {
  const current = siblingSelection(document, selectedIds);
  if (current.length < 2 || !current.some((control) => control.id === clickedId))
    return [clickedId];
  return [...current.map((control) => control.id).filter((id) => id !== clickedId), clickedId];
}

export function moveSiblingGroup(
  document: DesignerDocument,
  ids: readonly string[],
  dx: number,
  dy: number,
  snapping = false,
): ProductivityOperation | undefined {
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return;
  const controls = siblingSelection(document, ids);
  if (controls.length < 2) return;
  const primary = controls[controls.length - 1];
  const parent = parentBounds(document, primary);
  const left = Math.min(...controls.map((control) => control.left));
  const top = Math.min(...controls.map((control) => control.top));
  const right = Math.max(...controls.map((control) => control.left + control.width));
  const bottom = Math.max(...controls.map((control) => control.top + control.height));
  if (
    left < 0 ||
    top < 0 ||
    right > parent.width ||
    bottom > parent.height ||
    controls.some(
      (control) =>
        ![control.left, control.top, control.width, control.height].every(Number.isFinite) ||
        control.width < 1 ||
        control.height < 1 ||
        control.width > parent.width ||
        control.height > parent.height,
    )
  )
    return;

  let moveX = stableDelta(dx);
  let moveY = stableDelta(dy);
  if (snapping) {
    if (moveX !== 0)
      moveX = Math.round((primary.left + moveX) / gridSize) * gridSize - primary.left;
    if (moveY !== 0) moveY = Math.round((primary.top + moveY) / gridSize) * gridSize - primary.top;
  }
  if (moveX !== 0) moveX = Math.max(-left, Math.min(parent.width - right, moveX));
  if (moveY !== 0) moveY = Math.max(-top, Math.min(parent.height - bottom, moveY));
  moveX = stableDelta(moveX);
  moveY = stableDelta(moveY);
  if (moveX === 0 && moveY === 0) return;
  return {
    type: "moveControls",
    moves: controls.map((control) => ({
      controlId: control.id,
      left: moveX === 0 ? control.left : control.left + moveX,
      top: moveY === 0 ? control.top : control.top + moveY,
    })),
  };
}

export function selectedWorldBounds(
  document: DesignerDocument,
  ids: readonly string[],
): SelectionRect | undefined {
  const controls = siblingSelection(document, ids);
  if (!controls.length) return;
  const bounds = controls.map((control) => worldGeometry(document, control));
  return {
    left: Math.min(...bounds.map((geometry) => geometry.left)),
    top: Math.min(...bounds.map((geometry) => geometry.top)),
    right: Math.max(...bounds.map((geometry) => geometry.left + geometry.width)),
    bottom: Math.max(...bounds.map((geometry) => geometry.top + geometry.height)),
  };
}

export function fullyEnclosedChildren(
  document: DesignerDocument,
  parentId: string,
  selection: SelectionRect,
  rectFor: (id: string) => SelectionRect | undefined,
): string[] {
  return orderedChildren(document, parentId || undefined)
    .filter(
      (control) =>
        !isPage(control) &&
        isProductivityControl(control) &&
        isControlVisible(document, control.id),
    )
    .filter((control) => {
      const rect = rectFor(control.id);
      return (
        rect !== undefined &&
        rect.left >= selection.left &&
        rect.top >= selection.top &&
        rect.right <= selection.right &&
        rect.bottom <= selection.bottom
      );
    })
    .map((control) => control.id);
}

export function defaultPasteParent(
  document: DesignerDocument,
  selectedIds: readonly string[],
): string {
  const uniqueIds = [...new Set(selectedIds)];
  if (uniqueIds.length === 1) {
    const selected = document.controls.find((control) => control.id === uniqueIds[0]);
    if (selected?.type.toLowerCase() === "page" || selected?.type.toLowerCase() === "frame")
      return selected.id;
    return selected?.parentId ?? "";
  }
  const selected = siblingSelection(document, uniqueIds);
  const primary = selected.at(-1);
  if (!primary) return "";
  return primary.parentId ?? "";
}

export function defaultDuplicateParent(
  document: DesignerDocument,
  selectedIds: readonly string[],
): string {
  const controls = [...new Set(selectedIds)].map((id) =>
    document.controls.find((control) => control.id === id),
  );
  if (controls.length === 0 || controls.some((control) => !control)) return "";
  const parents = new Set(controls.map((control) => control!.parentId ?? ""));
  return parents.size === 1 ? (controls[0]!.parentId ?? "") : "";
}

export function eligibleTabOrder(
  document: DesignerDocument,
  propertyGrid: PropertyGridData | undefined,
  parentId: string,
): string[] {
  return orderedChildren(document, parentId || undefined)
    .map((control, index) => ({ control, index, target: propertyGrid?.controls[control.id] }))
    .filter(
      ({ control, target }) =>
        !isPage(control) &&
        target?.descriptors.some((descriptor) => descriptor.field === "tabIndex"),
    )
    .sort((a, b) => {
      const aValue = a.target?.values.tabIndex;
      const bValue = b.target?.values.tabIndex;
      const aIndex = aValue?.present && typeof aValue.value === "number" ? aValue.value : Infinity;
      const bIndex = bValue?.present && typeof bValue.value === "number" ? bValue.value : Infinity;
      return aIndex - bIndex || a.index - b.index;
    })
    .map(({ control }) => control.id);
}

export function tabOrderOperation(
  parentId: string,
  controlIds: readonly string[],
): ProductivityOperation {
  return { type: "setTabOrder", parentId, controlIds: [...controlIds] };
}

export function alignOperation(
  controlIds: readonly string[],
  anchorId: string,
  alignment: "left" | "center" | "right" | "top" | "middle" | "bottom",
): ProductivityOperation {
  return { type: "alignControls", controlIds: [...controlIds], anchorId, alignment };
}

export function distributeOperation(
  controlIds: readonly string[],
  axis: "horizontal" | "vertical",
): ProductivityOperation {
  return { type: "distributeControls", controlIds: [...controlIds], axis };
}

export function sameSizeOperation(
  controlIds: readonly string[],
  anchorId: string,
  dimension: "width" | "height" | "size",
): ProductivityOperation {
  return { type: "sameSizeControls", controlIds: [...controlIds], anchorId, dimension };
}

export function reorderOperation(
  controlIds: readonly string[],
  command: "front" | "forward" | "backward" | "back",
): ProductivityOperation {
  return { type: "reorderControls", controlIds: [...controlIds], command };
}
