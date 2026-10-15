import type { DesignerControl, DesignerDocument } from "../../src/userFormEditor/protocol";
import { pointsToPixels } from "../../src/userFormEditor/model";
import { pixelsToPoints } from "../../src/userFormEditor/model";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { displayPages } from "../../src/userFormEditor/collections";
import type { ActivePages } from "../../src/userFormEditor/collections";
import type { JSX } from "preact";
import type {
  GeometryOperation,
  ParentOperation,
  PropertyOperation,
  ProductivityOperation,
} from "../../src/userFormEditor/protocol";
import {
  fullyEnclosedChildren,
  moveSiblingGroup,
  promoteSelection,
  selectedWorldBounds,
  siblingSelection,
  updateSelection,
} from "../../src/userFormEditor/productivity";
import type { SelectionRect } from "../../src/userFormEditor/productivity";
import { subtreeIds } from "../../src/userFormEditor/toolbox";
import {
  canEditGeometry,
  isControlVisible,
  isDesignerFrame,
  isDesignerPage,
  contentInsets,
  geometryOperations,
  gridSize,
  moveGeometry,
  parentBounds,
  requiredSize,
  resizeGeometry,
  resizeHandles,
  dropTarget,
  placeGeometry,
  worldGeometry,
  orderedChildren,
  selectionPath,
  stableDelta,
} from "../../src/userFormEditor/geometry";
import type { Geometry, ResizeHandle, DropTarget } from "../../src/userFormEditor/geometry";
import { appearanceStyle } from "./propertyEditors/appearance";
const px = (points: number) => `${pointsToPixels(points)}px`;

export function ControlGlyph({ control }: { control: DesignerControl }) {
  const type = control.type.toLowerCase();
  const caption = control.caption ?? control.name;
  switch (type) {
    case "label":
    case "commandbutton":
    case "togglebutton":
      return <span>{caption}</span>;
    case "textbox":
      return <span>{control.text ?? String(control.value ?? "")}</span>;
    case "checkbox":
      return <span>☐ {caption}</span>;
    case "optionbutton":
      return <span>◯ {caption}</span>;
    case "combobox":
      return (
        <>
          <span class="combo-text">
            {control.text ?? String(control.value ?? control.list?.[0] ?? "")}
          </span>
          <span class="combo-arrow" aria-hidden="true">
            <span class="combo-chevron" />
          </span>
        </>
      );
    case "listbox":
      return (
        <span>
          {(control.list ?? []).map((item, index) => (
            <div key={index}>{item}</div>
          ))}
        </span>
      );
    case "spinbutton":
      return (
        <span class="spin">
          ▴<br />▾
        </span>
      );
    case "scrollbar":
      return <span class="scrollbar">◂ ━ ▸</span>;
    case "image":
      return control.picturePreview && control.picture?.remove !== true ? (
        <img
          class="control-picture"
          src={control.picturePreview}
          alt=""
          draggable={false}
          style={{
            objectFit:
              control.pictureSizeMode === 1
                ? "fill"
                : control.pictureSizeMode === 3
                  ? "contain"
                  : "none",
            objectPosition:
              ["left top", "right top", "center", "left bottom", "right bottom"][
                control.pictureAlignment ?? 2
              ] ?? "center",
          }}
        />
      ) : (
        <span class="placeholder">▧ Image</span>
      );
    case "frame":
      return <span class="frame-caption">{caption}</span>;
    case "multipage":
      return null;
    case "page":
      return null;
    case "tabstrip":
      return (
        <div class="tabs">
          {(control.tabs ?? []).map(
            (tab, index) =>
              tab.visible !== false && (
                <span key={tab.name} class={index === (control.selectedIndex ?? 0) ? "active" : ""}>
                  {tab.caption ?? tab.name}
                </span>
              ),
          )}
        </div>
      );
    default:
      return (
        <span class="placeholder">
          {control.type}: {control.name}
        </span>
      );
  }
}

export function DesignerCanvas({
  document: sourceDocument,
  approximateBounds = "approximate bounds",
  editable = false,
  parentEditable = false,
  pageEditable = false,
  formResizable = false,
  activePages,
  onActivePageChange,
  busy = false,
  zoom = 1,
  showGrid = true,
  snapping = false,
  productivityEditable = false,
  cannotMove = "This control is larger than its parent. Resize it to fit before moving.",
  selectedId: externalSelected,
  selectedIds: externalSelectedIds,
  onSelectionChange,
  onSelectionChangeMany,
  onCommit = () => {},
  onFormResize,
  onContextMenu,
  onShortcut,
  placing = false,
  onPlace,
  onCancelPlacement,
  onDelete,
  hierarchy = "Selection path",
  invalidDrop = "This control cannot be placed in this container.",
  dropTargetLabel = "Drop target",
  placementSize,
  resizeWidthLabel = "Resize form width",
  resizeHeightLabel = "Resize form height",
  resizeBothLabel = "Resize form width and height",
}: {
  document: DesignerDocument;
  approximateBounds?: string;
  editable?: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  formResizable?: boolean;
  activePages?: ActivePages;
  onActivePageChange?: (ownerId: string, pageId: string) => void;
  busy?: boolean;
  zoom?: number;
  showGrid?: boolean;
  snapping?: boolean;
  productivityEditable?: boolean;
  cannotMove?: string;
  selectedId?: string;
  selectedIds?: string[];
  onSelectionChange?: (id: string | undefined) => void;
  onSelectionChangeMany?: (ids: string[]) => void;
  onCommit?: (operations: (GeometryOperation | ParentOperation | ProductivityOperation)[]) => void;
  onFormResize?: (operations: PropertyOperation[]) => void;
  onContextMenu?: (controlId: string, x: number, y: number) => void;
  onShortcut?: (action: "copy" | "cut" | "paste" | "duplicate" | "delete") => boolean;
  placing?: boolean;
  onPlace?: (left: number, top: number, parentId?: string) => void;
  onCancelPlacement?: () => void;
  onDelete?: () => void;
  hierarchy?: string;
  invalidDrop?: string;
  dropTargetLabel?: string;
  placementSize?: { width: number; height: number };
  resizeWidthLabel?: string;
  resizeHeightLabel?: string;
  resizeBothLabel?: string;
}) {
  const [localActivePages, setLocalActivePages] = useState<ActivePages>({});
  const document = useMemo(
    () => displayPages(sourceDocument, activePages ?? localActivePages),
    [sourceDocument, activePages, localActivePages],
  );
  const [localSelection, setLocalSelection] = useState<string[]>([]);
  const controlledSelection = !!onSelectionChangeMany || !!onSelectionChange;
  const selectionIds = controlledSelection
    ? (externalSelectedIds ?? (externalSelected ? [externalSelected] : []))
    : localSelection;
  const selected = selectionIds.at(-1);
  const setSelection = (ids: string[]) => {
    if (onSelectionChangeMany) onSelectionChangeMany(ids);
    else if (onSelectionChange) onSelectionChange(ids.at(-1));
    else setLocalSelection(ids);
  };
  const [preview, setPreview] = useState<{ id: string; geometry: Geometry; floating?: boolean }>();
  const [formResizePreview, setFormResizePreview] = useState<{ width: number; height: number }>();
  const [groupPreview, setGroupPreview] = useState<{ id: string; left: number; top: number }[]>();
  const [marqueeBox, setMarqueeBox] = useState<SelectionRect>();
  const [destination, setDestination] = useState<{ target?: DropTarget; valid: boolean }>();
  const [notice, setNotice] = useState<string>();
  const gesture = useRef<{
    control: DesignerControl;
    start: Geometry;
    end: Geometry;
    x: number;
    y: number;
    handle?: ResizeHandle;
    pointerId?: number;
    keys?: Set<string>;
    dx: number;
    dy: number;
    target?: DropTarget;
    valid?: boolean;
    parenting?: boolean;
    group?: DesignerControl[];
    groupMoves?: Extract<ProductivityOperation, { type: "moveControls" }>;
  }>();
  const formResizeGesture = useRef<{
    handle: "e" | "s" | "se";
    pointerId: number;
    x: number;
    y: number;
    start: { width: number; height: number };
    end: { width: number; height: number };
  }>();
  const marquee = useRef<{
    pointerId: number;
    parentId: string;
    startX: number;
    startY: number;
    endX: number;
    endY: number;
  }>();
  const canvas = useRef<HTMLDivElement>(null);
  const cancel = () => {
    gesture.current = undefined;
    formResizeGesture.current = undefined;
    marquee.current = undefined;
    setPreview(undefined);
    setFormResizePreview(undefined);
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    setDestination(undefined);
  };
  const cancelPointer = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (
      gesture.current?.pointerId === event.pointerId ||
      formResizeGesture.current?.pointerId === event.pointerId ||
      marquee.current?.pointerId === event.pointerId
    )
      cancel();
  };
  const finish = () => {
    const current = gesture.current;
    gesture.current = undefined;
    if (!current) return;
    setDestination(undefined);
    if (current.group) {
      const operation = current.groupMoves;
      const changed = operation?.moves.some((move) => {
        const start = current.group!.find((control) => control.id === move.controlId)!;
        return move.left !== start.left || move.top !== start.top;
      });
      if (changed && editable) onCommit([operation!]);
      else setGroupPreview(undefined);
      return;
    }
    if (current.parenting && !current.valid) {
      setPreview(undefined);
      return;
    }
    const operations: (GeometryOperation | ParentOperation)[] =
      current.parenting &&
      current.target &&
      current.target.parentId !== (current.control.parentId ?? "")
        ? [
            {
              type: "setParent",
              controlId: current.control.id,
              parentId: current.target.parentId,
              left: current.end.left,
              top: current.end.top,
            },
          ]
        : geometryOperations(current.control.id, current.start, current.end);
    if (operations.length && editable) onCommit(operations);
    else setPreview(undefined);
  };
  useEffect(() => {
    cancel();
    if (selectionIds.length > 1) {
      const visible = siblingSelection(document, selectionIds)
        .filter((control) => isControlVisible(document, control.id))
        .map((control) => control.id);
      if (visible.length !== selectionIds.length) setSelection(visible);
    }
    if (
      selected &&
      !isControlVisible(document, selected) &&
      !document.controls.some(
        (c) =>
          c.id === selected &&
          c.type.toLowerCase() === "page" &&
          !!c.parentId &&
          document.controls.some((owner) => owner.id === c.parentId),
      )
    )
      setSelection([]);
  }, [document]);
  useEffect(() => {
    if (!formResizable) {
      formResizeGesture.current = undefined;
      setFormResizePreview(undefined);
    } else if (busy || placing) {
      if (formResizeGesture.current) {
        formResizeGesture.current = undefined;
        setFormResizePreview(undefined);
      }
    }
  }, [formResizable, busy, placing]);
  useEffect(() => {
    if (!editable) {
      gesture.current = undefined;
      if (!busy) setPreview(undefined);
    }
  }, [editable, busy]);
  useEffect(() => {
    cancel();
  }, [parentEditable, pageEditable, placing]);
  useEffect(() => {
    if (!busy && !gesture.current) cancel();
  }, [busy]);
  useEffect(() => {
    cancel();
  }, [zoom, snapping]);
  useEffect(() => {
    const blur = () => {
      if (gesture.current?.keys) finish();
      else cancel();
    };
    window.addEventListener("blur", blur);
    return () => window.removeEventListener("blur", blur);
  }, [editable, onCommit]);
  const startGesture = (
    control: DesignerControl,
    x: number,
    y: number,
    handle?: ResizeHandle,
    pointerId?: number,
  ) => {
    if (!editable || busy || !canEditGeometry(control)) return;
    const start = {
      left: control.left,
      top: control.top,
      width: control.width,
      height: control.height,
    };
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    gesture.current = { control, start, end: start, x, y, handle, pointerId, dx: 0, dy: 0 };
    setNotice(undefined);
    return gesture.current;
  };
  const startFormResize = (handle: "e" | "s" | "se", x: number, y: number, pointerId: number) => {
    if (!formResizable || busy || placing || selectionIds.length !== 0) return;
    const start = { width: document.width, height: document.height };
    formResizeGesture.current = { handle, x, y, pointerId, start, end: start };
    setFormResizePreview(start);
  };
  const updateFormResize = (dx: number, dy: number) => {
    const current = formResizeGesture.current;
    if (!current) return;
    const rootChildren = document.controls.filter((control) => !control.parentId);
    const minimumWidth = Math.max(
      1,
      ...rootChildren.map((control) => control.left + control.width),
    );
    const minimumHeight = Math.max(
      1,
      ...rootChildren.map((control) => control.top + control.height),
    );
    const resizeAxis = (start: number, delta: number, minimum: number) => {
      if (stableDelta(delta) === 0) return start;
      const target = start + stableDelta(delta);
      return stableDelta(
        Math.max(minimum, snapping ? Math.round(target / gridSize) * gridSize : target),
      );
    };
    const end = {
      width:
        current.handle === "e" || current.handle === "se"
          ? resizeAxis(current.start.width, dx, minimumWidth)
          : current.start.width,
      height:
        current.handle === "s" || current.handle === "se"
          ? resizeAxis(current.start.height, dy, minimumHeight)
          : current.start.height,
    };
    current.end = end;
    setFormResizePreview(end);
  };
  const finishFormResize = () => {
    const current = formResizeGesture.current;
    formResizeGesture.current = undefined;
    if (!current) return;
    const operations: PropertyOperation[] = [];
    if (current.end.width !== current.start.width)
      operations.push({ type: "setFormProperty", field: "width", value: current.end.width });
    if (current.end.height !== current.start.height)
      operations.push({ type: "setFormProperty", field: "height", value: current.end.height });
    if (!operations.length || !onFormResize || !formResizable || busy || placing) {
      setFormResizePreview(undefined);
      return;
    }
    onFormResize?.(operations);
  };
  const startGroupGesture = (
    controls: DesignerControl[],
    x: number,
    y: number,
    pointerId?: number,
  ) => {
    if (
      !productivityEditable ||
      !editable ||
      busy ||
      controls.length < 2 ||
      controls.some((control) => !canEditGeometry(control))
    )
      return;
    const primary = controls[controls.length - 1];
    const start = {
      left: primary.left,
      top: primary.top,
      width: primary.width,
      height: primary.height,
    };
    setPreview(undefined);
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    gesture.current = {
      control: primary,
      start,
      end: start,
      x,
      y,
      pointerId,
      dx: 0,
      dy: 0,
      group: controls,
    };
    setNotice(undefined);
    return gesture.current;
  };
  const updateGroupGesture = (dx: number, dy: number) => {
    const current = gesture.current;
    if (!current?.group) return;
    const operation = moveSiblingGroup(
      document,
      current.group.map((control) => control.id),
      dx,
      dy,
      snapping,
    );
    if (!operation || operation.type !== "moveControls") {
      current.groupMoves = undefined;
      setGroupPreview(undefined);
      setNotice(cannotMove);
      return;
    }
    current.groupMoves = operation;
    const primaryMove = operation.moves.find((move) => move.controlId === current.control.id);
    if (primaryMove) current.end = { ...current.end, left: primaryMove.left, top: primaryMove.top };
    setGroupPreview(
      operation.moves.map(({ controlId, left, top }) => ({ id: controlId, left, top })),
    );
    setNotice(undefined);
  };
  const updateGesture = (dx: number, dy: number) => {
    const current = gesture.current;
    if (!current) return;
    if (current.group) {
      updateGroupGesture(dx, dy);
      return;
    }
    const bounds = parentBounds(document, current.control);
    const oldParent = document.controls.find((c) => c.id === current.control.parentId);
    if (
      parentEditable &&
      !current.handle &&
      !current.keys &&
      (!oldParent ||
        isDesignerFrame(oldParent) ||
        (pageEditable && isDesignerPage(document, oldParent))) &&
      (dx !== 0 || dy !== 0)
    ) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (!form) return;
      const rect = form.getBoundingClientRect();
      const point = {
        left: pixelsToPoints(current.x - rect.left, zoom) + dx,
        top: pixelsToPoints(current.y - rect.top, zoom) + dy,
      };
      const target = dropTarget(
        document,
        point,
        subtreeIds(document.controls, current.control.id),
        pageEditable,
      );
      const startWorld = worldGeometry(document, current.control);
      const world = { ...startWorld, left: startWorld.left + dx, top: startWorld.top + dy };
      const end =
        target &&
        (target.parentId === (current.control.parentId ?? "")
          ? moveGeometry(current.start, dx, dy, target.bounds, snapping)
          : placeGeometry(
              {
                ...world,
                left: world.left - target.origin.left,
                top: world.top - target.origin.top,
              },
              target.bounds,
              snapping,
            ));
      current.parenting = true;
      current.target = target;
      current.valid = !!end;
      if (end && target) current.end = end;
      setDestination({ target, valid: !!end });
      setNotice(end ? undefined : invalidDrop);
      setPreview({
        id: current.control.id,
        geometry:
          end && target
            ? { ...end, left: end.left + target.origin.left, top: end.top + target.origin.top }
            : world,
        floating: true,
      });
      return;
    }
    current.parenting = false;
    current.target = undefined;
    current.valid = undefined;
    setDestination(undefined);
    const end =
      dx === 0 && dy === 0
        ? current.start
        : current.handle
          ? resizeGeometry(
              current.start,
              current.handle,
              dx,
              dy,
              bounds,
              requiredSize(document, current.control),
              snapping,
            )
          : moveGeometry(current.start, dx, dy, bounds, current.keys ? false : snapping);
    if (!end) {
      setNotice(cannotMove);
      return;
    }
    current.end = end;
    setPreview({ id: current.control.id, geometry: end });
  };
  const pointerDown = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (
      event.button !== 0 ||
      gesture.current?.pointerId !== undefined ||
      formResizeGesture.current?.pointerId !== undefined ||
      marquee.current?.pointerId !== undefined
    )
      return;
    if (gesture.current?.keys) finish();
    const target = event.target as Element;
    if (placing) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (form && form.contains(target) && !busy) {
        event.preventDefault();
        cancel();
        canvas.current?.focus({ preventScroll: true });
        const bounds = form.getBoundingClientRect();
        const point = {
          left: pixelsToPoints(event.clientX - bounds.left, zoom),
          top: pixelsToPoints(event.clientY - bounds.top, zoom),
        };
        const target = parentEditable
          ? dropTarget(document, point, new Set(), pageEditable)
          : { parentId: "", origin: { left: 0, top: 0 }, bounds: document };
        if (!target) {
          setNotice(invalidDrop);
          return;
        }
        onPlace?.(
          point.left - target.origin.left,
          point.top - target.origin.top,
          target.parentId || undefined,
        );
      }
      return;
    }
    const formHandle = target.closest<HTMLElement>("[data-form-handle]")?.dataset.formHandle as
      | "e"
      | "s"
      | "se"
      | undefined;
    if (formHandle) {
      event.preventDefault();
      canvas.current?.focus({ preventScroll: true });
      startFormResize(formHandle, event.clientX, event.clientY, event.pointerId);
      if (formResizeGesture.current) {
        try {
          event.currentTarget.setPointerCapture(event.pointerId);
        } catch {
          /* no active pointer */
        }
      }
      return;
    }
    if (target.closest("[data-page-tab]")) return;
    const blankArea = target.closest<HTMLElement>(".container-content, .page-content, .form");
    const candidate = target.closest<HTMLElement>(
      ".control[data-control-id], .selection-outline[data-control-id]",
    );
    const id =
      candidate && (!blankArea || blankArea.contains(candidate))
        ? candidate.dataset.controlId
        : undefined;
    const control = document.controls.find((c) => c.id === id);
    setNotice(undefined);
    canvas.current?.focus({ preventScroll: true });
    if (!control) {
      setSelection([]);
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (productivityEditable && form?.contains(target) && !busy) {
        const container = target.closest<HTMLElement>("[data-container-id]");
        const current = {
          pointerId: event.pointerId,
          parentId: container?.dataset.containerId ?? "",
          startX: event.clientX,
          startY: event.clientY,
          endX: event.clientX,
          endY: event.clientY,
        };
        marquee.current = current;
        setMarqueeBox({
          left: current.startX,
          top: current.startY,
          right: current.startX,
          bottom: current.startY,
        });
        try {
          event.currentTarget.setPointerCapture(event.pointerId);
        } catch {
          /* no active pointer */
        }
      }
      return;
    }
    event.preventDefault();
    const modified = event.ctrlKey || event.metaKey || event.shiftKey;
    if (productivityEditable && modified) {
      setSelection(updateSelection(document, selectionIds, control.id, true));
      setMarqueeBox(undefined);
      return;
    }
    const remainsSelected = selectionIds.includes(control.id);
    const nextSelection =
      productivityEditable && remainsSelected && siblingSelection(document, selectionIds).length > 1
        ? promoteSelection(document, selectionIds, control.id)
        : [control.id];
    if (nextSelection.some((id, index) => id !== selectionIds[index])) setSelection(nextSelection);
    const selectedControls =
      productivityEditable && nextSelection.length > 1
        ? siblingSelection(document, nextSelection)
        : [];
    const handle = target.closest<HTMLElement>("[data-handle]")?.dataset.handle as
      | ResizeHandle
      | undefined;
    if (selectedControls.length > 1 && !handle)
      startGroupGesture(selectedControls, event.clientX, event.clientY, event.pointerId);
    else startGesture(control, event.clientX, event.clientY, handle, event.pointerId);
    if (gesture.current) {
      // Synthetic Webview tests have no active native pointer to capture.
      try {
        event.currentTarget.setPointerCapture(event.pointerId);
      } catch {
        /* no active pointer */
      }
    }
  };
  const contextMenu = (event: JSX.TargetedMouseEvent<HTMLDivElement>) => {
    if (!productivityEditable) return;
    const target = event.target as Element;
    if (target.closest("[data-page-tab]")) return;
    const blankArea = target.closest<HTMLElement>(".container-content, .page-content, .form");
    const candidate = target.closest<HTMLElement>(
      ".control[data-control-id], .selection-outline[data-control-id]",
    );
    const id =
      candidate && (!blankArea || blankArea.contains(candidate))
        ? candidate.dataset.controlId
        : undefined;
    if (!id) return;
    event.preventDefault();
    if (document.controls.find((control) => control.id === id)?.type.toLowerCase() === "page")
      setSelection([id]);
    else if (!selectionIds.includes(id)) setSelection([id]);
    else setSelection(promoteSelection(document, selectionIds, id));
    canvas.current?.focus({ preventScroll: true });
    onContextMenu?.(id, event.clientX, event.clientY);
  };
  const keyDown = (event: JSX.TargetedKeyboardEvent<HTMLDivElement>) => {
    if ((event.target as Element).closest("input, textarea, select, [contenteditable='true']"))
      return;
    if (
      productivityEditable &&
      (event.ctrlKey || event.metaKey) &&
      !event.altKey &&
      !event.repeat &&
      ["c", "x", "v", "d"].includes(event.key.toLowerCase())
    ) {
      const action =
        event.key.toLowerCase() === "c"
          ? "copy"
          : event.key.toLowerCase() === "x"
            ? "cut"
            : event.key.toLowerCase() === "v"
              ? "paste"
              : "duplicate";
      if (onShortcut?.(action)) event.preventDefault();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      cancel();
      onCancelPlacement?.();
      return;
    }
    if (
      event.key === "Delete" &&
      event.target === event.currentTarget &&
      !busy &&
      !placing &&
      !event.repeat &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey
    ) {
      if (productivityEditable && onShortcut?.("delete")) {
        event.preventDefault();
        cancel();
        return;
      }
      if (onDelete) {
        event.preventDefault();
        cancel();
        onDelete();
        return;
      }
    }
    if (
      placing ||
      !editable ||
      busy ||
      event.ctrlKey ||
      event.metaKey ||
      event.altKey ||
      !["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key)
    )
      return;
    const control = document.controls.find((c) => c.id === selected);
    if (!control || !canEditGeometry(control) || gesture.current?.pointerId !== undefined) return;
    event.preventDefault();
    if (!gesture.current) {
      const selectedControls =
        productivityEditable && selectionIds.length > 1
          ? siblingSelection(document, selectionIds)
          : [];
      const created =
        selectedControls.length > 1
          ? startGroupGesture(selectedControls, 0, 0)
          : startGesture(control, 0, 0);
      if (created) created.keys = new Set();
    }
    const current = gesture.current;
    if (!current) return;
    current.keys!.add(event.key);
    const step = event.shiftKey ? 10 : 1;
    current.dx =
      current.end.left -
      current.start.left +
      (event.key === "ArrowLeft" ? -step : event.key === "ArrowRight" ? step : 0);
    current.dy =
      current.end.top -
      current.start.top +
      (event.key === "ArrowUp" ? -step : event.key === "ArrowDown" ? step : 0);
    updateGesture(current.dx, current.dy);
  };
  const children = (parentId?: string) => orderedChildren(document, parentId);
  const draw = (control: DesignerControl | undefined, floating = false): JSX.Element | null => {
    if (!control || control.visible === false) return null;
    if (preview?.floating && preview.id === control.id && !floating) return null;
    const type = control.type.toLowerCase();
    const pages = type === "multipage" ? children(control.id) : [];
    const selected = control.selectedIndex ?? (pages.length ? 0 : -1);
    const page = pages[selected];
    const groupPosition = groupPreview?.find((item) => item.id === control.id);
    const shown =
      preview?.id === control.id
        ? { ...control, ...preview.geometry }
        : groupPosition
          ? { ...control, left: groupPosition.left, top: groupPosition.top }
          : control;
    const insets = contentInsets(type);
    const selectedControl = selectionIds.includes(control.id);
    return (
      <>
        <div
          key={control.id}
          class={`control ${type} ${control.enabled === false ? "disabled" : ""}`}
          data-control-id={control.id}
          title={`${control.name} (${control.type})${control.approximate ? ` — ${approximateBounds}` : ""}`}
          style={{
            ...appearanceStyle(shown),
            left: px(shown.left),
            top: px(shown.top),
            width: px(shown.width),
            height: px(shown.height),
          }}
        >
          <ControlGlyph control={shown} />
          {type === "multipage" ? (
            <>
              <div class="tabs">
                {pages.map(
                  (p, index) =>
                    p.visible !== false && (
                      <button
                        type="button"
                        key={p.id}
                        data-page-tab={p.id}
                        class={`${index === selected ? "active" : ""} ${selectedId === p.id ? "selected-page" : ""}`}
                        aria-pressed={selectedId === p.id}
                        onClick={(event) => {
                          event.stopPropagation();
                          cancel();
                          if (onActivePageChange) onActivePageChange(control.id, p.id);
                          else
                            setLocalActivePages((current) => ({ ...current, [control.id]: p.id }));
                          setSelection([p.id]);
                        }}
                      >
                        {p.caption ?? p.name}
                      </button>
                    ),
                )}
              </div>
              {page && page.visible !== false && (
                <div
                  class="control page page-content"
                  data-container-id={page.id}
                  title={page.name}
                  style={{
                    ...appearanceStyle(page),
                    left: px(insets.left),
                    top: px(insets.top),
                    right: px(insets.right),
                    bottom: px(insets.bottom),
                  }}
                >
                  {children(page.id).map((c) => draw(c))}
                </div>
              )}
            </>
          ) : type === "frame" ? (
            <div
              class="container-content"
              data-container-id={control.id}
              style={{
                left: px(insets.left),
                top: px(insets.top),
                right: px(insets.right),
                bottom: px(insets.bottom),
              }}
            >
              {children(control.id).map((c) => draw(c))}
            </div>
          ) : null}
        </div>
        {selectedControl && (
          <div
            class="selection-outline"
            data-control-id={control.id}
            style={{
              left: px(shown.left),
              top: px(shown.top),
              width: px(shown.width),
              height: px(shown.height),
            }}
          >
            {editable &&
              selectionIds.length === 1 &&
              canEditGeometry(control) &&
              resizeHandles.map((handle) => (
                <span
                  class={`resize-handle handle-${handle}`}
                  data-handle={handle}
                  style={{ transform: `translate(-50%, -50%) scale(${1 / zoom})` }}
                />
              ))}
          </div>
        )}
      </>
    );
  };
  const selectedId = selected;
  const groupBounds =
    productivityEditable && selectionIds.length > 1
      ? selectedWorldBounds(
          groupPreview
            ? {
                ...document,
                controls: document.controls.map((control) => {
                  const position = groupPreview.find((item) => item.id === control.id);
                  return position
                    ? { ...control, left: position.left, top: position.top }
                    : control;
                }),
              }
            : document,
          selectionIds,
        )
      : undefined;
  return (
    <div class="designer-canvas">
      <nav class="selection-path" aria-label={hierarchy}>
        <button
          type="button"
          disabled={busy}
          onClick={() => {
            cancel();
            setSelection([]);
            canvas.current?.focus({ preventScroll: true });
          }}
        >
          {document.name}
        </button>
        {selectionPath(document, selected).map((control) => (
          <span key={control.id}>
            {" "}
            ›{" "}
            <button
              type="button"
              disabled={busy}
              aria-current={control.id === selected ? "true" : undefined}
              onClick={() => {
                cancel();
                setSelection([control.id]);
                canvas.current?.focus({ preventScroll: true });
              }}
            >
              {control.name}
            </button>
          </span>
        ))}
      </nav>
      {notice && (
        <p class="notice" role="status">
          {notice}
        </p>
      )}
      <div
        ref={canvas}
        class={`designer-viewport ${placing ? "placing-control" : ""}`}
        tabIndex={0}
        aria-label={document.name}
        onPointerDown={pointerDown}
        onContextMenu={contextMenu}
        onPointerMove={(event) => {
          const current = gesture.current;
          const formResize = formResizeGesture.current;
          if (placing && parentEditable && !busy) {
            const form = canvas.current?.querySelector<HTMLElement>(".form");
            if (form) {
              const rect = form.getBoundingClientRect();
              const target = dropTarget(
                document,
                {
                  left: pixelsToPoints(event.clientX - rect.left, zoom),
                  top: pixelsToPoints(event.clientY - rect.top, zoom),
                },
                new Set(),
                pageEditable,
              );
              const valid =
                !!target &&
                (!placementSize ||
                  (placementSize.width <= target.bounds.width &&
                    placementSize.height <= target.bounds.height));
              setDestination({ target, valid });
              setNotice(valid ? undefined : invalidDrop);
            }
          }
          const activeMarquee = marquee.current;
          if (activeMarquee?.pointerId === event.pointerId) {
            activeMarquee.endX = event.clientX;
            activeMarquee.endY = event.clientY;
            setMarqueeBox({
              left: Math.min(activeMarquee.startX, activeMarquee.endX),
              top: Math.min(activeMarquee.startY, activeMarquee.endY),
              right: Math.max(activeMarquee.startX, activeMarquee.endX),
              bottom: Math.max(activeMarquee.startY, activeMarquee.endY),
            });
          }
          if (current?.pointerId === event.pointerId)
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
          if (formResize?.pointerId === event.pointerId)
            updateFormResize(
              pixelsToPoints(event.clientX - formResize.x, zoom),
              pixelsToPoints(event.clientY - formResize.y, zoom),
            );
        }}
        onPointerUp={(event) => {
          const formResize = formResizeGesture.current;
          if (formResize?.pointerId === event.pointerId) {
            updateFormResize(
              pixelsToPoints(event.clientX - formResize.x, zoom),
              pixelsToPoints(event.clientY - formResize.y, zoom),
            );
            finishFormResize();
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
            return;
          }
          const current = gesture.current;
          if (current?.pointerId === event.pointerId) {
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
            finish();
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
            return;
          }
          const activeMarquee = marquee.current;
          if (activeMarquee?.pointerId === event.pointerId) {
            activeMarquee.endX = event.clientX;
            activeMarquee.endY = event.clientY;
            const selection = {
              left: Math.min(activeMarquee.startX, activeMarquee.endX),
              top: Math.min(activeMarquee.startY, activeMarquee.endY),
              right: Math.max(activeMarquee.startX, activeMarquee.endX),
              bottom: Math.max(activeMarquee.startY, activeMarquee.endY),
            };
            const form = canvas.current?.querySelector<HTMLElement>(".form");
            const rectFor = (id: string) => {
              const element = Array.from(
                form?.querySelectorAll<HTMLElement>(".control[data-control-id]") ?? [],
              ).find((candidate) => candidate.dataset.controlId === id);
              if (!element) return undefined;
              const rect = element.getBoundingClientRect();
              return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
            };
            setSelection(
              fullyEnclosedChildren(document, activeMarquee.parentId, selection, rectFor),
            );
            marquee.current = undefined;
            setMarqueeBox(undefined);
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
          }
        }}
        onPointerCancel={cancelPointer}
        onLostPointerCapture={cancelPointer}
        onKeyDown={keyDown}
        onKeyUp={(event) => {
          const current = gesture.current;
          current?.keys?.delete(event.key);
          if (current?.keys && !current.keys.size) finish();
        }}
        onBlur={() => {
          if (gesture.current?.keys) finish();
          else cancel();
        }}
        style={{
          width: `${pointsToPixels((formResizePreview ?? document).width + 2, zoom)}px`,
          height: `${pointsToPixels((formResizePreview ?? document).height + 24, zoom)}px`,
        }}
      >
        <div class="form-shell" style={{ transform: `scale(${zoom})` }}>
          <div class="form-title">{document.caption}</div>
          <div class="form-client-area">
            <div
              class={`form ${showGrid ? "show-grid" : ""} ${selectedId === undefined ? "form-selected" : ""}`}
              data-container-id=""
              style={{
                ...appearanceStyle(document),
                width: px((formResizePreview ?? document).width),
                height: px((formResizePreview ?? document).height),
                backgroundSize: `${px(gridSize)} ${px(gridSize)}`,
              }}
            >
              {children().map((c) => draw(c))}
              {groupBounds && (
                <div
                  class="group-selection-outline"
                  aria-hidden="true"
                  style={{
                    left: px(groupBounds.left),
                    top: px(groupBounds.top),
                    width: px(groupBounds.right - groupBounds.left),
                    height: px(groupBounds.bottom - groupBounds.top),
                  }}
                />
              )}
              {marqueeBox &&
                (() => {
                  const form = canvas.current?.querySelector<HTMLElement>(".form");
                  const rect = form?.getBoundingClientRect();
                  if (!rect) return null;
                  return (
                    <div
                      class="marquee-selection"
                      aria-hidden="true"
                      style={{
                        left: px(pixelsToPoints(marqueeBox.left - rect.left, zoom)),
                        top: px(pixelsToPoints(marqueeBox.top - rect.top, zoom)),
                        width: px(pixelsToPoints(marqueeBox.right - marqueeBox.left, zoom)),
                        height: px(pixelsToPoints(marqueeBox.bottom - marqueeBox.top, zoom)),
                      }}
                    />
                  );
                })()}
              {destination?.target && (
                <div
                  class={`drop-target ${destination.valid ? "valid" : "invalid"}`}
                  data-drop-parent-id={destination.target.parentId}
                  aria-label={dropTargetLabel}
                  style={{
                    left: px(destination.target.origin.left),
                    top: px(destination.target.origin.top),
                    width: px(destination.target.bounds.width),
                    height: px(destination.target.bounds.height),
                  }}
                />
              )}
              {preview?.floating && (
                <div class="drag-preview" aria-hidden="true">
                  {draw(
                    document.controls.find((c) => c.id === preview.id),
                    true,
                  )}
                </div>
              )}
            </div>
            {formResizable && !busy && !placing && selectionIds.length === 0 && (
              <>
                {(
                  [
                    ["e", resizeWidthLabel],
                    ["s", resizeHeightLabel],
                    ["se", resizeBothLabel],
                  ] as const
                ).map(([handle, label]) => (
                  <span
                    key={handle}
                    class={`form-resize-handle form-resize-${handle}`}
                    data-form-handle={handle}
                    role="button"
                    tabIndex={-1}
                    aria-label={label}
                    title={label}
                    style={{ transform: `translate(-50%, -50%) scale(${1 / zoom})` }}
                  />
                ))}
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
