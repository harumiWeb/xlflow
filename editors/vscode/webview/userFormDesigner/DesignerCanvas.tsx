import type { DesignerControl, DesignerDocument } from "../../src/userFormEditor/protocol";
import { pointsToPixels } from "../../src/userFormEditor/model";
import { pixelsToPoints } from "../../src/userFormEditor/model";
import { useEffect, useRef, useState } from "preact/hooks";
import type { JSX } from "preact";
import type { GeometryOperation, ParentOperation } from "../../src/userFormEditor/protocol";
import { subtreeIds } from "../../src/userFormEditor/toolbox";
import {
  canEditGeometry,
  isControlVisible,
  isDesignerFrame,
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
} from "../../src/userFormEditor/geometry";
import type { Geometry, ResizeHandle, DropTarget } from "../../src/userFormEditor/geometry";
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
      return <span class="placeholder">▧ Image</span>;
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
  document,
  approximateBounds = "approximate bounds",
  editable = false,
  parentEditable = false,
  busy = false,
  zoom = 1,
  showGrid = true,
  snapping = false,
  cannotMove = "This control is larger than its parent. Resize it to fit before moving.",
  selectedId: externalSelected,
  onSelectionChange,
  onCommit = () => {},
  placing = false,
  onPlace,
  onCancelPlacement,
  onDelete,
  hierarchy = "Selection path",
  invalidDrop = "This control cannot be placed in this container.",
  dropTargetLabel = "Drop target",
  placementSize,
}: {
  document: DesignerDocument;
  approximateBounds?: string;
  editable?: boolean;
  parentEditable?: boolean;
  busy?: boolean;
  zoom?: number;
  showGrid?: boolean;
  snapping?: boolean;
  cannotMove?: string;
  selectedId?: string;
  onSelectionChange?: (id: string | undefined) => void;
  onCommit?: (operations: (GeometryOperation | ParentOperation)[]) => void;
  placing?: boolean;
  onPlace?: (left: number, top: number, parentId?: string) => void;
  onCancelPlacement?: () => void;
  onDelete?: () => void;
  hierarchy?: string;
  invalidDrop?: string;
  dropTargetLabel?: string;
  placementSize?: { width: number; height: number };
}) {
  const [localSelected, setLocalSelected] = useState<string>();
  const selected = onSelectionChange ? externalSelected : localSelected;
  const setSelected = onSelectionChange ?? setLocalSelected;
  const [preview, setPreview] = useState<{ id: string; geometry: Geometry; floating?: boolean }>();
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
  }>();
  const canvas = useRef<HTMLDivElement>(null);
  const cancel = () => {
    gesture.current = undefined;
    setPreview(undefined);
    setDestination(undefined);
  };
  const cancelPointer = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (gesture.current?.pointerId === event.pointerId) cancel();
  };
  const finish = () => {
    const current = gesture.current;
    gesture.current = undefined;
    if (!current) return;
    setDestination(undefined);
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
    if (
      selected &&
      !isControlVisible(document, selected) &&
      !document.controls.some(
        (c) =>
          c.id === selected &&
          c.type.toLowerCase() === "page" &&
          c.visible !== false &&
          !!c.parentId &&
          isControlVisible(document, c.parentId),
      )
    )
      setSelected(undefined);
  }, [document]);
  useEffect(() => {
    if (!editable) {
      gesture.current = undefined;
      if (!busy) setPreview(undefined);
    }
  }, [editable, busy]);
  useEffect(() => {
    cancel();
  }, [parentEditable, placing]);
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
    gesture.current = { control, start, end: start, x, y, handle, pointerId, dx: 0, dy: 0 };
    setNotice(undefined);
    return gesture.current;
  };
  const updateGesture = (dx: number, dy: number) => {
    const current = gesture.current;
    if (!current) return;
    const bounds = parentBounds(document, current.control);
    const oldParent = document.controls.find((c) => c.id === current.control.parentId);
    if (
      parentEditable &&
      !current.handle &&
      !current.keys &&
      (!oldParent || isDesignerFrame(oldParent)) &&
      (dx !== 0 || dy !== 0)
    ) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (!form) return;
      const rect = form.getBoundingClientRect();
      const point = {
        left: pixelsToPoints(current.x - rect.left, zoom) + dx,
        top: pixelsToPoints(current.y - rect.top, zoom) + dy,
      };
      const target = dropTarget(document, point, subtreeIds(document.controls, current.control.id));
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
    if (event.button !== 0 || gesture.current?.pointerId !== undefined) return;
    if (gesture.current?.keys) finish();
    const target = event.target as Element;
    if (placing) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (form && form.contains(target) && !busy) {
        event.preventDefault();
        cancel();
        canvas.current?.focus();
        const bounds = form.getBoundingClientRect();
        const point = {
          left: pixelsToPoints(event.clientX - bounds.left, zoom),
          top: pixelsToPoints(event.clientY - bounds.top, zoom),
        };
        const target = parentEditable
          ? dropTarget(document, point)
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
    if (target.closest("[data-page-tab]")) return;
    const id = target.closest<HTMLElement>("[data-control-id]")?.dataset.controlId;
    const control = document.controls.find((c) => c.id === id);
    setSelected(control?.id);
    setNotice(undefined);
    canvas.current?.focus();
    if (!control) return;
    event.preventDefault();
    startGesture(
      control,
      event.clientX,
      event.clientY,
      target.closest<HTMLElement>("[data-handle]")?.dataset.handle as ResizeHandle | undefined,
      event.pointerId,
    );
    if (gesture.current) {
      // Synthetic Webview tests have no active native pointer to capture.
      try {
        event.currentTarget.setPointerCapture(event.pointerId);
      } catch {
        /* no active pointer */
      }
    }
  };
  const keyDown = (event: JSX.TargetedKeyboardEvent<HTMLDivElement>) => {
    if ((event.target as Element).closest("input, textarea, select, [contenteditable='true']"))
      return;
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
      !event.altKey &&
      onDelete
    ) {
      event.preventDefault();
      cancel();
      onDelete();
      return;
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
      const created = startGesture(control, 0, 0);
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
  const draw = (control: DesignerControl, floating = false): JSX.Element | null => {
    if (control.visible === false) return null;
    if (preview?.floating && preview.id === control.id && !floating) return null;
    const type = control.type.toLowerCase();
    const pages = type === "multipage" ? children(control.id) : [];
    const selected = control.selectedIndex ?? (pages.length ? 0 : -1);
    const page = pages[selected];
    const shown = preview?.id === control.id ? { ...control, ...preview.geometry } : control;
    const insets = contentInsets(type);
    const selectedControl = selectedId === control.id;
    return (
      <>
        <div
          key={control.id}
          class={`control ${type} ${control.enabled === false ? "disabled" : ""}`}
          data-control-id={control.id}
          title={`${control.name} (${control.type})${control.approximate ? ` — ${approximateBounds}` : ""}`}
          style={{
            left: px(shown.left),
            top: px(shown.top),
            width: px(shown.width),
            height: px(shown.height),
          }}
        >
          <ControlGlyph control={control} />
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
                          setSelected(p.id);
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
                  title={page.name}
                  style={{
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
          ) : (
            <div
              class="container-content"
              style={{
                left: px(insets.left),
                top: px(insets.top),
                right: px(insets.right),
                bottom: px(insets.bottom),
              }}
            >
              {children(control.id).map((c) => draw(c))}
            </div>
          )}
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
  return (
    <div class="designer-canvas">
      <nav class="selection-path" aria-label={hierarchy}>
        <button
          type="button"
          disabled={busy}
          onClick={() => {
            cancel();
            setSelected(undefined);
            canvas.current?.focus();
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
                setSelected(control.id);
                canvas.current?.focus();
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
        onPointerMove={(event) => {
          const current = gesture.current;
          if (placing && parentEditable && !busy) {
            const form = canvas.current?.querySelector<HTMLElement>(".form");
            if (form) {
              const rect = form.getBoundingClientRect();
              const target = dropTarget(document, {
                left: pixelsToPoints(event.clientX - rect.left, zoom),
                top: pixelsToPoints(event.clientY - rect.top, zoom),
              });
              const valid =
                !!target &&
                (!placementSize ||
                  (placementSize.width <= target.bounds.width &&
                    placementSize.height <= target.bounds.height));
              setDestination({ target, valid });
              setNotice(valid ? undefined : invalidDrop);
            }
          }
          if (current?.pointerId === event.pointerId)
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
        }}
        onPointerUp={(event) => {
          const current = gesture.current;
          if (current?.pointerId === event.pointerId) {
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
            finish();
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
          width: `${pointsToPixels(document.width + 2, zoom)}px`,
          height: `${pointsToPixels(document.height + 24, zoom)}px`,
        }}
      >
        <div class="form-shell" style={{ transform: `scale(${zoom})` }}>
          <div class="form-title">{document.caption}</div>
          <div
            class={`form ${showGrid ? "show-grid" : ""} ${selectedId === undefined ? "form-selected" : ""}`}
            style={{
              width: px(document.width),
              height: px(document.height),
              backgroundSize: `${px(gridSize)} ${px(gridSize)}`,
            }}
          >
            {children().map((c) => draw(c))}
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
                {draw(document.controls.find((c) => c.id === preview.id)!, true)}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
