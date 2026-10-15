import { useEffect, useRef } from "preact/hooks";
import { CommandMenu } from "./CommandMenu";
import { Icon } from "./Icon";
import type {
  DesignerDocument,
  DesignerStrings,
  SemanticOperation,
} from "../../src/userFormEditor/protocol";
import {
  alignOperation,
  defaultDuplicateParent,
  defaultPasteParent,
  distributeOperation,
  reorderOperation,
  sameSizeOperation,
  siblingSelection,
} from "../../src/userFormEditor/productivity";

export type ProductivityClipboardAction = "copy" | "cut" | "paste" | "duplicate";

export function ProductivityToolbar({
  document,
  selectionIds,
  strings,
  enabled,
  structuralEnabled,
  duplicateEditable = true,
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
  strings: DesignerStrings;
  enabled: boolean;
  structuralEnabled: boolean;
  duplicateEditable?: boolean;
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
  const canArrange = enabled && !busy && ids.length >= 2;
  const canReorder = enabled && !busy && ids.length >= 1;
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
    const popup = menuRef.current;
    if (popup) {
      const bounds = popup.getBoundingClientRect();
      popup.style.left = `${Math.max(4, Math.min(contextMenu.x, window.innerWidth - bounds.width - 4))}px`;
      popup.style.top = `${Math.max(4, Math.min(contextMenu.y, window.innerHeight - bounds.height - 4))}px`;
    }
    items()[0]?.focus();
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        onCloseMenu();
        window.document
          .querySelector<HTMLElement>(".designer-viewport")
          ?.focus({ preventScroll: true });
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
    window.addEventListener("keydown", keydown, true);
    window.addEventListener("pointerdown", outside, true);
    return () => {
      window.removeEventListener("keydown", keydown, true);
      window.removeEventListener("pointerdown", outside, true);
    };
  }, [contextMenu, onCloseMenu]);

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
          {(
            [
              ["cut", strings.cut, "screen-cut", cutAllowed],
              ["copy", strings.copy, "copy", canCopy],
              ["paste", strings.paste, "clippy", pasteAllowed],
            ] as const
          ).map(([action, label, icon, allowed]) => (
            <button
              type="button"
              class="icon-button"
              aria-label={label}
              title={label}
              disabled={!allowed}
              onClick={() => clipboard(action)}
            >
              <Icon name={icon} />
              <span class="sr-only">{label}</span>
            </button>
          ))}
          <button
            type="button"
            class="icon-button"
            aria-label={strings.deleteItem}
            title={strings.deleteItem}
            disabled={!canRemove}
            onClick={remove}
          >
            <Icon name="trash" />
            <span class="sr-only">{strings.deleteItem}</span>
          </button>
          <span class="toolbar-separator" aria-hidden="true" />
          {ids.length >= 2 && (
            <CommandMenu label={strings.align}>
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
                  role="menuitem"
                  type="button"
                  disabled={!canArrange}
                  onClick={() => onOperation(alignOperation(ids, anchor, alignment))}
                >
                  {label}
                </button>
              ))}
            </CommandMenu>
          )}
          {ids.length >= 3 && (
            <CommandMenu label={strings.distribute}>
              <button
                role="menuitem"
                type="button"
                disabled={!canArrange}
                onClick={() => onOperation(distributeOperation(ids, "horizontal"))}
              >
                {strings.distributeHorizontal}
              </button>
              <button
                role="menuitem"
                type="button"
                disabled={!canArrange}
                onClick={() => onOperation(distributeOperation(ids, "vertical"))}
              >
                {strings.distributeVertical}
              </button>
            </CommandMenu>
          )}
          {ids.length >= 2 && (
            <CommandMenu label={strings.sameSize}>
              {(
                [
                  ["width", strings.sameWidth],
                  ["height", strings.sameHeight],
                  ["size", strings.sameDimensions],
                ] as const
              ).map(([dimension, label]) => (
                <button
                  role="menuitem"
                  type="button"
                  disabled={!canArrange}
                  onClick={() => onOperation(sameSizeOperation(ids, anchor, dimension))}
                >
                  {label}
                </button>
              ))}
            </CommandMenu>
          )}
          {ids.length >= 1 && (
            <CommandMenu label={strings.order}>
              {(
                [
                  ["front", strings.bringToFront],
                  ["forward", strings.bringForward],
                  ["backward", strings.sendBackward],
                  ["back", strings.sendToBack],
                ] as const
              ).map(([command, label]) => (
                <button
                  role="menuitem"
                  type="button"
                  disabled={!canReorder}
                  onClick={() => zOrder(command)}
                >
                  {label}
                </button>
              ))}
            </CommandMenu>
          )}
          <CommandMenu label={strings.moreActions}>
            <button
              role="menuitem"
              type="button"
              disabled={!duplicateAllowed}
              onClick={() => clipboard("duplicate")}
            >
              {strings.duplicate}
            </button>
            <button
              role="menuitem"
              type="button"
              disabled={!hasSelectedControl}
              onClick={showProperties}
            >
              {strings.propertiesAction}
            </button>
          </CommandMenu>
        </div>
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
