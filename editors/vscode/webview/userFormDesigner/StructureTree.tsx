import { useEffect, useRef, useState } from "preact/hooks";
import { orderedChildren } from "../../src/userFormEditor/geometry";
import type {
  DesignerControl,
  DesignerDocument,
  DesignerStrings,
} from "../../src/userFormEditor/protocol";
import { Icon } from "./Icon";

interface TreeEntry {
  id: string;
  parentId?: string;
  control?: DesignerControl;
  children: TreeEntry[];
}

function createTree(document: DesignerDocument) {
  const ids = new Set(document.controls.map((control) => control.id));
  const controls: DesignerControl[] = document.controls.map((control) => {
    if (!control.parentId || ids.has(control.parentId)) return control;
    const rootControl = { ...control };
    delete rootControl.parentId;
    return rootControl;
  });
  const normalized = { ...document, controls };
  const entries = new Map<string, TreeEntry>();
  const visited = new Set<string>();

  const build = (
    control: DesignerControl,
    parentId: string,
    ancestors: Set<string>,
  ): TreeEntry | undefined => {
    if (visited.has(control.id) || ancestors.has(control.id)) return undefined;
    visited.add(control.id);
    const nextAncestors = new Set(ancestors).add(control.id);
    const entry: TreeEntry = { id: control.id, parentId, control, children: [] };
    entries.set(control.id, entry);
    entry.children = orderedChildren(normalized, control.id)
      .map((child) => build(child, control.id, nextAncestors))
      .filter((child): child is TreeEntry => child !== undefined);
    return entry;
  };

  const roots = orderedChildren(normalized)
    .map((control) => build(control, "", new Set()))
    .filter((entry): entry is TreeEntry => entry !== undefined);

  // Keep malformed cyclic parent chains visible as stable root entries.
  for (const control of controls) {
    if (visited.has(control.id)) continue;
    const entry = build(control, "", new Set());
    if (entry) roots.push(entry);
  }

  const root: TreeEntry = { id: "", children: roots };
  entries.set("", root);
  return { root, entries };
}

function iconName(entry: TreeEntry): string {
  if (!entry.control) return "symbol-namespace";
  if (entry.control.type.toLowerCase() === "frame") return "symbol-class";
  if (entry.control.type.toLowerCase() === "page") return "file-directory";
  return "symbol-field";
}

export function StructureTree({
  document,
  selectionIds,
  strings,
  busy,
  eventEditable = false,
  eventPending = false,
  onSelect,
  onDelete,
  onOpenEvent,
}: {
  document: DesignerDocument;
  selectionIds: string[];
  strings: DesignerStrings;
  busy: boolean;
  eventEditable?: boolean;
  eventPending?: boolean;
  onSelect: (id?: string) => void;
  onDelete?: () => boolean;
  onOpenEvent?: (targetId?: string, choose?: boolean) => void;
}) {
  const { root, entries } = createTree(document);
  const selectionKey = JSON.stringify(selectionIds);
  const [expandedIds, setExpandedIds] = useState<Set<string>>(() => new Set([""]));
  const [focusedId, setFocusedId] = useState(() => selectionIds.at(-1) ?? "");
  const [focusMoved, setFocusMoved] = useState(false);
  const [contextMenu, setContextMenu] = useState<
    { x: number; y: number; targetId?: string } | undefined
  >();
  const itemRefs = useRef(new Map<string, HTMLDivElement>());
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const focusAfterRender = useRef<string | undefined>(undefined);
  const previousSelectionKey = useRef(selectionKey);

  useEffect(() => {
    if (!contextMenu) return;
    const popup = contextMenuRef.current;
    if (popup) {
      const bounds = popup.getBoundingClientRect();
      popup.style.left = `${Math.max(4, Math.min(contextMenu.x, window.innerWidth - bounds.width - 4))}px`;
      popup.style.top = `${Math.max(4, Math.min(contextMenu.y, window.innerHeight - bounds.height - 4))}px`;
      popup.querySelector<HTMLButtonElement>('button[role="menuitem"]')?.focus();
    }
    const keydown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      event.stopPropagation();
      setContextMenu(undefined);
    };
    const outside = (event: PointerEvent) => {
      if (!contextMenuRef.current?.contains(event.target as Node)) setContextMenu(undefined);
    };
    window.addEventListener("keydown", keydown, true);
    window.addEventListener("pointerdown", outside, true);
    return () => {
      window.removeEventListener("keydown", keydown, true);
      window.removeEventListener("pointerdown", outside, true);
    };
  }, [contextMenu]);
  const previousParents = useRef(new Map<string, string>());
  const activeStructureId =
    typeof window === "undefined"
      ? undefined
      : window.document.activeElement?.closest<HTMLElement>("[role='treeitem']")?.dataset
          .structureId;
  const removedActiveId =
    activeStructureId !== undefined && !entries.has(activeStructureId)
      ? activeStructureId
      : undefined;

  const visibleEntries: { entry: TreeEntry; level: number }[] = [];
  const collectVisible = (entry: TreeEntry, level: number) => {
    visibleEntries.push({ entry, level });
    if (!expandedIds.has(entry.id)) return;
    for (const child of entry.children) collectVisible(child, level + 1);
  };
  collectVisible(root, 1);
  const visibleIds = new Set(visibleEntries.map(({ entry }) => entry.id));
  const visibleKey = visibleEntries.map(({ entry }) => entry.id).join("\0");

  useEffect(() => {
    const next = new Set(expandedIds);
    next.add("");
    for (const id of selectionIds) {
      const visited = new Set<string>();
      let current = entries.get(id);
      while (current?.control && !visited.has(current.id)) {
        visited.add(current.id);
        next.add(current.parentId ?? "");
        current = current.parentId ? entries.get(current.parentId) : undefined;
      }
    }
    if (next.size !== expandedIds.size || [...next].some((id) => !expandedIds.has(id)))
      setExpandedIds(next);
  }, [selectionKey, document.controls]);

  useEffect(() => {
    const selectionChanged = previousSelectionKey.current !== selectionKey;
    previousSelectionKey.current = selectionKey;

    if (selectionChanged) {
      const selected = selectionIds.at(-1) ?? "";
      const nextFocusId = entries.has(selected) ? selected : "";
      if (removedActiveId !== undefined) focusAfterRender.current = nextFocusId;
      setFocusMoved(false);
      setFocusedId(nextFocusId);
      return;
    }

    const unavailableFocusId = !entries.has(focusedId) ? focusedId : removedActiveId;
    if (unavailableFocusId !== undefined) {
      let parentId = previousParents.current.get(unavailableFocusId);
      const visited = new Set<string>();
      while (parentId !== undefined && !visibleIds.has(parentId) && !visited.has(parentId)) {
        visited.add(parentId);
        parentId = entries.get(parentId)?.parentId ?? previousParents.current.get(parentId);
      }
      const nextFocusId = parentId !== undefined && visibleIds.has(parentId) ? parentId : "";
      if (removedActiveId !== undefined) focusAfterRender.current = nextFocusId;
      setFocusedId(nextFocusId);
    }
  }, [selectionKey, visibleKey, focusedId, document.controls]);

  useEffect(() => {
    previousParents.current = new Map(
      [...entries].map(([id, entry]) => [id, entry.parentId ?? ""]),
    );
  }, [document.controls]);

  useEffect(() => {
    const id = focusAfterRender.current;
    if (id === undefined || id !== focusedId) return;
    itemRefs.current.get(id)?.focus();
    focusAfterRender.current = undefined;
  }, [focusedId, visibleKey]);

  const selectedFocusId = selectionIds.at(-1) ?? "";
  const rovingId = focusMoved
    ? visibleIds.has(focusedId)
      ? focusedId
      : ""
    : visibleIds.has(selectedFocusId)
      ? selectedFocusId
      : visibleIds.has(focusedId)
        ? focusedId
        : "";
  const focusItem = (id: string) => {
    setFocusMoved(true);
    if (focusedId === id) itemRefs.current.get(id)?.focus();
    else {
      focusAfterRender.current = id;
      setFocusedId(id);
    }
  };

  const isDescendant = (candidateId: string, ancestorId: string) => {
    let current = entries.get(candidateId);
    const visited = new Set<string>();
    while (current?.parentId !== undefined && !visited.has(current.id)) {
      visited.add(current.id);
      if (current.parentId === ancestorId) return true;
      current = entries.get(current.parentId);
    }
    return false;
  };

  const toggleExpanded = (id: string) => {
    const willCollapse = expandedIds.has(id);
    if (willCollapse && (focusedId === id || isDescendant(focusedId, id))) focusItem(id);
    else if (!willCollapse) focusItem(id);
    setExpandedIds((current) => {
      const next = new Set(current);
      if (willCollapse) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const select = (id: string) => {
    focusItem(id);
    onSelect(id || undefined);
  };

  const onTreeKeyDown = (event: KeyboardEvent) => {
    if (
      busy ||
      event.isComposing ||
      (event.target as Element).closest(
        "input, textarea, select, [contenteditable]:not([contenteditable='false'])",
      )
    )
      return;
    const target = (event.target as Element).closest<HTMLElement>("[role='treeitem']");
    if (!target) return;
    const id = target.dataset.structureId ?? "";
    const index = visibleEntries.findIndex(({ entry }) => entry.id === id);
    const entry = entries.get(id);
    if (!entry || index < 0) return;

    if (
      (event.key === "Backspace" || event.key === "Delete") &&
      !event.repeat &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      !event.shiftKey &&
      id &&
      selectionIds.includes(id) &&
      onDelete?.()
    ) {
      event.preventDefault();
      return;
    }

    switch (event.key) {
      case "ArrowDown":
        event.preventDefault();
        focusItem(visibleEntries[Math.min(index + 1, visibleEntries.length - 1)].entry.id);
        break;
      case "ArrowUp":
        event.preventDefault();
        focusItem(visibleEntries[Math.max(index - 1, 0)].entry.id);
        break;
      case "Home":
        event.preventDefault();
        focusItem(visibleEntries[0].entry.id);
        break;
      case "End":
        event.preventDefault();
        focusItem(visibleEntries[visibleEntries.length - 1].entry.id);
        break;
      case "ArrowRight":
        event.preventDefault();
        if (entry.children.length && !expandedIds.has(id)) toggleExpanded(id);
        else if (entry.children.length) focusItem(entry.children[0].id);
        break;
      case "ArrowLeft":
        event.preventDefault();
        if (entry.children.length && expandedIds.has(id)) toggleExpanded(id);
        else if (entry.parentId !== undefined) focusItem(entry.parentId);
        break;
      case "Enter":
      case " ":
      case "Spacebar":
        event.preventDefault();
        select(id);
        break;
    }
  };

  const renderEntry = (entry: TreeEntry, level: number) => {
    const isRoot = entry.control === undefined;
    const hasChildren = entry.children.length > 0;
    const expanded = expandedIds.has(entry.id);
    const selected = isRoot ? selectionIds.length === 0 : selectionIds.includes(entry.id);
    const label = isRoot
      ? document.name || document.caption || strings.structure
      : entry.control?.name || entry.id;

    return (
      <div
        key={entry.id}
        class="structure-tree-item"
        role="treeitem"
        data-structure-id={entry.id}
        aria-label={label}
        aria-level={level}
        aria-selected={selected}
        aria-expanded={hasChildren ? expanded : undefined}
        tabIndex={rovingId === entry.id ? 0 : -1}
        onContextMenu={(event) => {
          if (!eventEditable || !onOpenEvent) return;
          event.preventDefault();
          event.stopPropagation();
          setContextMenu({ x: event.clientX, y: event.clientY, targetId: entry.control?.id });
        }}
        ref={(element) => {
          if (element) itemRefs.current.set(entry.id, element);
          else itemRefs.current.delete(entry.id);
        }}
      >
        <div class="structure-tree-row" onClick={() => select(entry.id)}>
          {hasChildren ? (
            <span
              class="structure-tree-disclosure"
              aria-hidden="true"
              onClick={(event) => {
                event.stopPropagation();
                toggleExpanded(entry.id);
              }}
            >
              <Icon name={expanded ? "chevron-down" : "chevron-right"} />
            </span>
          ) : (
            <span
              class="structure-tree-disclosure structure-tree-disclosure-empty"
              aria-hidden="true"
            />
          )}
          <Icon name={iconName(entry)} />
          <span class="structure-tree-label">{label}</span>
          {!isRoot && <span class="structure-tree-type">{entry.control?.type}</span>}
        </div>
        {hasChildren && expanded && (
          <div class="structure-tree-children" role="group">
            {entry.children.map((child) => renderEntry(child, level + 1))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div
      class="structure-tree"
      role="tree"
      aria-label={strings.structure}
      aria-multiselectable={selectionIds.length > 1}
      aria-busy={busy}
      onKeyDown={onTreeKeyDown}
    >
      {renderEntry(root, 1)}
      {contextMenu && (
        <div
          ref={contextMenuRef}
          class="designer-context-menu"
          role="menu"
          style={{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }}
          onContextMenu={(event) => {
            event.preventDefault();
            event.stopPropagation();
          }}
        >
          <button
            role="menuitem"
            type="button"
            disabled={
              busy || eventPending || selectionIds.length > 1 || !eventEditable || !onOpenEvent
            }
            onClick={() => {
              const targetId = contextMenu.targetId;
              setContextMenu(undefined);
              onOpenEvent?.(targetId, true);
            }}
          >
            {eventPending ? strings.eventPending : strings.openEvent}
          </button>
        </div>
      )}
    </div>
  );
}
