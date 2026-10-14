import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { isDesignerFrame, isDesignerPage } from "../../src/userFormEditor/geometry";
import type {
  DesignerDocument,
  DesignerStrings,
  SemanticOperation,
  PropertyGridData,
} from "../../src/userFormEditor/protocol";
import {
  alignOperation,
  defaultDuplicateParent,
  defaultPasteParent,
  distributeOperation,
  eligibleTabOrder,
  reorderOperation,
  sameSizeOperation,
  siblingSelection,
  tabOrderOperation,
} from "../../src/userFormEditor/productivity";

export type ProductivityClipboardAction = "copy" | "cut" | "paste" | "duplicate";

export function ProductivityToolbar({
  document,
  selectionIds,
  propertyGrid,
  strings,
  enabled,
  structuralEnabled,
  duplicateEditable = true,
  tabEnabled,
  clipboardAvailable,
  canPaste,
  busy,
  contextMenu,
  onOperation,
  onClipboard,
  onProperties,
  onCloseMenu,
}: {
  document: DesignerDocument;
  selectionIds: string[];
  propertyGrid?: PropertyGridData;
  strings: DesignerStrings;
  enabled: boolean;
  structuralEnabled: boolean;
  duplicateEditable?: boolean;
  tabEnabled: boolean;
  clipboardAvailable: boolean;
  canPaste: boolean;
  busy: boolean;
  contextMenu?: { x: number; y: number };
  onOperation: (operation: SemanticOperation) => void;
  onClipboard: (
    action: ProductivityClipboardAction,
    controlIds: string[],
    parentId: string,
  ) => void;
  onProperties: () => void;
  onCloseMenu: () => void;
}) {
  const controls = siblingSelection(document, selectionIds);
  const ids = controls.map((control) => control.id);
  const primaryId = ids.at(-1);
  const anchor = primaryId ?? "";
  const parentForPaste = defaultPasteParent(document, selectionIds);
  const originalParent = defaultDuplicateParent(document, selectionIds);
  const canArrange = enabled && ids.length >= 2;
  const canReorder = enabled && ids.length >= 1;
  const canSelect = ids.length >= 1;
  const hasSelectedControl = selectionIds.some((id) =>
    document.controls.some((control) => control.id === id),
  );
  // The host clipboard bridge is independently capability-gated; it does not
  // require the UserForm geometry edit capability.
  const canCopy = canSelect && !busy;
  const pasteAllowed = clipboardAvailable && canPaste && !busy;
  const cutAllowed = canCopy && structuralEnabled;
  const duplicateAllowed = canCopy && structuralEnabled && duplicateEditable;
  const canRemove =
    selectionIds.some((id) =>
      document.controls.some(
        (control) => control.id === id && control.type.toLowerCase() !== "page",
      ),
    ) &&
    structuralEnabled &&
    !busy;
  const menuRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!contextMenu) return;
    const items = () =>
      Array.from(
        menuRef.current?.querySelectorAll<HTMLButtonElement>(
          'button[role="menuitem"]:not(:disabled)',
        ) ?? [],
      );
    items()[0]?.focus();
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCloseMenu();
        window.document.querySelector<HTMLElement>(".designer-viewport")?.focus();
      } else if (
        menuRef.current?.contains(event.target as Node) &&
        ["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)
      ) {
        event.preventDefault();
        const enabledItems = items();
        const current = enabledItems.indexOf(window.document.activeElement as HTMLButtonElement);
        const next =
          event.key === "Home"
            ? 0
            : event.key === "End"
              ? enabledItems.length - 1
              : (current + (event.key === "ArrowDown" ? 1 : -1) + enabledItems.length) %
                enabledItems.length;
        enabledItems[next]?.focus();
      }
    };
    const outside = (event: PointerEvent) => {
      if (!menuRef.current?.contains(event.target as Node)) onCloseMenu();
    };
    window.addEventListener("keydown", keydown);
    window.addEventListener("pointerdown", outside, true);
    return () => {
      window.removeEventListener("keydown", keydown);
      window.removeEventListener("pointerdown", outside, true);
    };
  }, [contextMenu, onCloseMenu]);

  const parentChoices = useMemo(() => {
    const seen = new Set<string>();
    const choices = eligibleParents(document, propertyGrid);
    return choices.filter((id) => {
      if (seen.has(id)) return false;
      seen.add(id);
      return true;
    });
  }, [document, propertyGrid]);
  const preferredParent = controls[0]?.parentId ?? "";
  const [tabParent, setTabParent] = useState(preferredParent);
  const [tabOrder, setTabOrder] = useState(() =>
    eligibleTabOrder(document, propertyGrid, tabParent),
  );
  useEffect(() => {
    const nextParent = parentChoices.includes(preferredParent) ? preferredParent : "";
    setTabParent(nextParent);
  }, [document, propertyGrid, preferredParent]);
  useEffect(() => {
    setTabOrder(eligibleTabOrder(document, propertyGrid, tabParent));
  }, [document, propertyGrid, tabParent]);
  const canonicalTabOrder = eligibleTabOrder(document, propertyGrid, tabParent);
  const tabOrderChanged =
    canonicalTabOrder.length === tabOrder.length &&
    canonicalTabOrder.some((id, index) => id !== tabOrder[index]);
  const tabIndexesAreSequential = tabOrder.every((id, index) => {
    const state = propertyGrid?.controls[id]?.values.tabIndex;
    return state?.present === true && typeof state.value === "number" && state.value === index;
  });
  const tabOrderNeedsApply = tabOrderChanged || !tabIndexesAreSequential;
  const setTabPosition = (index: number, delta: -1 | 1) => {
    const destination = index + delta;
    if (destination < 0 || destination >= tabOrder.length) return;
    setTabOrder((current) => {
      const next = [...current];
      [next[index], next[destination]] = [next[destination], next[index]];
      return next;
    });
  };
  const clipboard = (action: ProductivityClipboardAction) =>
    onClipboard(
      action,
      action === "paste" ? [] : ids,
      action === "paste"
        ? parentForPaste
        : action === "duplicate"
          ? originalParent
          : parentForPaste,
    );
  const remove = () => {
    if (ids.length) onOperation({ type: "removeControls", controlIds: ids, cascade: true });
    else {
      const controlId = selectionIds.at(-1);
      const control = document.controls.find((candidate) => candidate.id === controlId);
      if (control && control.type.toLowerCase() !== "page")
        onOperation({ type: "removeControl", controlId: control.id, cascade: true });
    }
  };
  const showProperties = () => {
    onProperties();
    onCloseMenu();
  };
  const runClipboard = (action: ProductivityClipboardAction) => {
    clipboard(action);
    onCloseMenu();
  };
  const runRemove = () => {
    remove();
    onCloseMenu();
  };
  const zOrder = (command: "front" | "forward" | "backward" | "back") => {
    onOperation(reorderOperation(ids, command));
    onCloseMenu();
  };

  return (
    <>
      <section class="productivity-toolbar" aria-label={strings.productivity}>
        <div class="productivity-command-row">
          <button type="button" disabled={!cutAllowed} onClick={() => clipboard("cut")}>
            {strings.cut}
          </button>
          <button type="button" disabled={!canCopy} onClick={() => clipboard("copy")}>
            {strings.copy}
          </button>
          <button type="button" disabled={!pasteAllowed} onClick={() => clipboard("paste")}>
            {strings.paste}
          </button>
          <button type="button" disabled={!canRemove} onClick={remove}>
            {strings.deleteItem}
          </button>
          <button type="button" disabled={!duplicateAllowed} onClick={() => clipboard("duplicate")}>
            {strings.duplicate}
          </button>
          <button
            type="button"
            disabled={!canReorder}
            onClick={() => onOperation(reorderOperation(ids, "front"))}
          >
            {strings.bringToFront}
          </button>
          <button
            type="button"
            disabled={!canReorder}
            onClick={() => onOperation(reorderOperation(ids, "back"))}
          >
            {strings.sendToBack}
          </button>
          <button type="button" disabled={!hasSelectedControl} onClick={showProperties}>
            {strings.propertiesAction}
          </button>
          <span class="productivity-selection-count" aria-live="polite">
            {strings.selectionCount}: {selectionIds.length}
          </span>
        </div>
        <details>
          <summary>
            {strings.align} / {strings.distribute} / {strings.sameSize}
          </summary>
          <div class="productivity-command-row">
            {(
              [
                ["left", strings.alignLeft],
                ["center", strings.alignCenter],
                ["right", strings.alignRight],
                ["top", strings.alignTop],
                ["middle", strings.alignMiddle],
                ["bottom", strings.alignBottom],
              ] as const
            ).map(([alignment, label]) => (
              <button
                type="button"
                disabled={!canArrange}
                onClick={() => onOperation(alignOperation(ids, anchor, alignment))}
              >
                {label}
              </button>
            ))}
            <button
              type="button"
              disabled={!enabled || ids.length < 3}
              onClick={() => onOperation(distributeOperation(ids, "horizontal"))}
            >
              {strings.distributeHorizontal}
            </button>
            <button
              type="button"
              disabled={!enabled || ids.length < 3}
              onClick={() => onOperation(distributeOperation(ids, "vertical"))}
            >
              {strings.distributeVertical}
            </button>
            {(
              [
                ["width", strings.sameWidth],
                ["height", strings.sameHeight],
                ["size", strings.sameDimensions],
              ] as const
            ).map(([dimension, label]) => (
              <button
                type="button"
                disabled={!canArrange}
                onClick={() => onOperation(sameSizeOperation(ids, anchor, dimension))}
              >
                {label}
              </button>
            ))}
            <button
              type="button"
              disabled={!canReorder}
              onClick={() => onOperation(reorderOperation(ids, "forward"))}
            >
              {strings.bringForward}
            </button>
            <button
              type="button"
              disabled={!canReorder}
              onClick={() => onOperation(reorderOperation(ids, "backward"))}
            >
              {strings.sendBackward}
            </button>
          </div>
        </details>
        <details>
          <summary>{strings.tabOrder}</summary>
          <div class="tab-order-editor">
            <label>
              {strings.tabOrderParent}
              <select
                aria-label={strings.tabOrderParent}
                value={tabParent}
                disabled={!tabEnabled || busy}
                onChange={(event) => setTabParent(event.currentTarget.value)}
              >
                {parentChoices.map((parentId) => (
                  <option value={parentId}>
                    {parentId
                      ? (document.controls.find((control) => control.id === parentId)?.name ??
                        parentId)
                      : document.name}
                  </option>
                ))}
              </select>
            </label>
            {tabOrder.length === 0 ? (
              <p class="notice">{strings.noEditableTabOrder}</p>
            ) : (
              <ol>
                {tabOrder.map((id, index) => {
                  const control = document.controls.find((candidate) => candidate.id === id);
                  if (!control) return null;
                  return (
                    <li key={id}>
                      <span>
                        {index}. {control.name}
                      </span>
                      <button
                        type="button"
                        aria-label={`${strings.moveUp}: ${control.name}`}
                        disabled={!tabEnabled || busy || index === 0}
                        onClick={() => setTabPosition(index, -1)}
                      >
                        {strings.moveUp}
                      </button>
                      <button
                        type="button"
                        aria-label={`${strings.moveDown}: ${control.name}`}
                        disabled={!tabEnabled || busy || index === tabOrder.length - 1}
                        onClick={() => setTabPosition(index, 1)}
                      >
                        {strings.moveDown}
                      </button>
                    </li>
                  );
                })}
              </ol>
            )}
            <button
              type="button"
              disabled={!tabEnabled || busy || !tabOrderNeedsApply}
              onClick={() => onOperation(tabOrderOperation(tabParent, tabOrder))}
            >
              {strings.applyTabOrder}
            </button>
          </div>
        </details>
      </section>
      {contextMenu && (
        <div
          ref={menuRef}
          class="designer-context-menu"
          role="menu"
          style={{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }}
          onContextMenu={(event) => event.preventDefault()}
        >
          <button
            role="menuitem"
            type="button"
            disabled={!cutAllowed}
            onClick={() => runClipboard("cut")}
          >
            {strings.cut}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!canCopy}
            onClick={() => runClipboard("copy")}
          >
            {strings.copy}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!pasteAllowed}
            onClick={() => runClipboard("paste")}
          >
            {strings.paste}
          </button>
          <button role="menuitem" type="button" disabled={!canRemove} onClick={runRemove}>
            {strings.deleteItem}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!duplicateAllowed}
            onClick={() => runClipboard("duplicate")}
          >
            {strings.duplicate}
          </button>
          <hr />
          <button
            role="menuitem"
            type="button"
            disabled={!canReorder}
            onClick={() => zOrder("front")}
          >
            {strings.bringToFront}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!canReorder}
            onClick={() => zOrder("forward")}
          >
            {strings.bringForward}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!canReorder}
            onClick={() => zOrder("backward")}
          >
            {strings.sendBackward}
          </button>
          <button
            role="menuitem"
            type="button"
            disabled={!canReorder}
            onClick={() => zOrder("back")}
          >
            {strings.sendToBack}
          </button>
          <hr />
          <button
            role="menuitem"
            type="button"
            disabled={!hasSelectedControl}
            onClick={showProperties}
          >
            {strings.propertiesAction}
          </button>
        </div>
      )}
    </>
  );
}

function eligibleParents(document: DesignerDocument, propertyGrid?: PropertyGridData): string[] {
  const eligibleIds = new Set(
    document.controls
      .filter(
        (control) =>
          control.type.toLowerCase() !== "page" &&
          propertyGrid?.controls[control.id]?.descriptors.some(
            (descriptor) => descriptor.field === "tabIndex",
          ),
      )
      .map((control) => control.parentId ?? ""),
  );
  return [
    "",
    ...[...eligibleIds].filter((id) => {
      const parent = id && document.controls.find((control) => control.id === id);
      return !!parent && (isDesignerFrame(parent) || isDesignerPage(document, parent));
    }),
  ];
}
