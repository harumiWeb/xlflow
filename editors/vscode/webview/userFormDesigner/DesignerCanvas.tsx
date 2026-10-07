import type { DesignerControl, DesignerDocument } from "../../src/userFormEditor/protocol";
import { pointsToPixels } from "../../src/userFormEditor/model";
import { pixelsToPoints } from "../../src/userFormEditor/model";
import { useEffect, useRef, useState } from "preact/hooks";
import type { JSX } from "preact";
import type { GeometryOperation } from "../../src/userFormEditor/protocol";
import {
  canEditGeometry,
  isControlVisible,
  contentInsets,
  geometryOperations,
  gridSize,
  moveGeometry,
  parentBounds,
  requiredSize,
  resizeGeometry,
  resizeHandles,
} from "../../src/userFormEditor/geometry";
import type { Geometry, ResizeHandle } from "../../src/userFormEditor/geometry";
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
}: {
  document: DesignerDocument;
  approximateBounds?: string;
  editable?: boolean;
  busy?: boolean;
  zoom?: number;
  showGrid?: boolean;
  snapping?: boolean;
  cannotMove?: string;
  selectedId?: string;
  onSelectionChange?: (id: string | undefined) => void;
  onCommit?: (operations: GeometryOperation[]) => void;
  placing?: boolean;
  onPlace?: (left: number, top: number) => void;
  onCancelPlacement?: () => void;
  onDelete?: () => void;
}) {
  const [localSelected, setLocalSelected] = useState<string>();
  const selected = onSelectionChange ? externalSelected : localSelected;
  const setSelected = onSelectionChange ?? setLocalSelected;
  const [preview, setPreview] = useState<{ id: string; geometry: Geometry }>();
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
  }>();
  const canvas = useRef<HTMLDivElement>(null);
  const cancel = () => {
    gesture.current = undefined;
    setPreview(undefined);
  };
  const cancelPointer = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (gesture.current?.pointerId === event.pointerId) cancel();
  };
  const finish = () => {
    const current = gesture.current;
    gesture.current = undefined;
    if (!current) return;
    const operations = geometryOperations(current.control.id, current.start, current.end);
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
        onPlace?.(
          pixelsToPoints(event.clientX - bounds.left, zoom),
          pixelsToPoints(event.clientY - bounds.top, zoom),
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
  const children = (parentId?: string): DesignerControl[] =>
    document.controls
      .filter((control) => (control.parentId || undefined) === parentId)
      .map((control, index) => ({ control, index }))
      .sort((a, b) => (a.control.zIndex ?? a.index) - (b.control.zIndex ?? b.index))
      .map(({ control }) => control);
  const draw = (control: DesignerControl) => {
    if (control.visible === false) return null;
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
                  {children(page.id).map(draw)}
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
              {children(control.id).map(draw)}
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
    <>
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
            if (event.currentTarget.hasPointerCapture(event.pointerId))
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
            {children().map(draw)}
          </div>
        </div>
      </div>
    </>
  );
}
