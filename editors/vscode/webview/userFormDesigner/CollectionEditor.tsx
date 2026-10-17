import { useState } from "preact/hooks";
import type {
  DesignerDocument,
  DesignerStrings,
  PropertyGridData,
  SemanticOperation,
} from "../../src/userFormEditor/protocol";
import { createPage, moveToPage } from "../../src/userFormEditor/collections";
import {
  canEditGeometry,
  isDesignerPage,
  orderedChildren,
  selectionPath,
} from "../../src/userFormEditor/geometry";
import { subtreeIds } from "../../src/userFormEditor/toolbox";
import { PropertyGrid } from "./PropertyGrid";
import type { PropertyDraftStore, PropertyReply } from "./PropertyGrid";

export function CollectionEditor({
  document,
  selectedId,
  pageEditable,
  tabEditable,
  parentEditable,
  editable,
  busy,
  strings,
  onSelect,
  onActivePage,
  onCommit,
  propertyGrid,
  drafts,
  reply,
  version,
  snapping,
  onInvalidMove,
}: {
  document: DesignerDocument;
  selectedId?: string;
  pageEditable: boolean;
  tabEditable: boolean;
  parentEditable: boolean;
  editable: boolean;
  busy: boolean;
  strings: DesignerStrings;
  onSelect: (id: string) => void;
  onActivePage: (owner: string, page: string) => void;
  onCommit: (ops: SemanticOperation[]) => number | undefined;
  propertyGrid?: PropertyGridData;
  drafts: PropertyDraftStore;
  reply?: PropertyReply;
  version: number;
  snapping: boolean;
  onInvalidMove: () => void;
}) {
  const [tabSelection, setTabSelection] = useState<{
    owner: string;
    name: string;
    index: number;
  }>();
  const selected = document.controls.find((c) => c.id === selectedId);
  const owner =
    selected &&
    selectionPath(document, selected.id)
      .reverse()
      .find((c) => c.type.toLowerCase() === "multipage");
  const pages = owner ? orderedChildren(document, owner.id) : [];
  const tabTargets = selected ? propertyGrid?.tabs?.[selected.id] : undefined;
  // Use authored metadata, rather than observed tabs synthesized for rendering.
  const tabs = (tabTargets ?? []).map((t) => ({
    name: String(t.values.name?.value ?? ""),
    caption: t.values.caption?.value,
  }));
  const foundTab =
    tabSelection && tabSelection.owner === selectedId
      ? tabs.findIndex((t) => t.name === tabSelection.name)
      : -1;
  const tabIndex =
    foundTab >= 0
      ? foundTab
      : Math.min(
          tabSelection && tabSelection.owner === selectedId ? tabSelection.index : 0,
          tabs.length - 1,
        );
  const pageDisabled = !pageEditable || busy;
  const tabDisabled = !tabEditable || busy;
  const allPages = document.controls.filter(
    (c) =>
      isDesignerPage(document, c) &&
      selected &&
      !subtreeIds(document.controls, selected.id).has(c.id) &&
      c.id !== selected.parentId,
  );
  return (
    <>
      {owner && (
        <details class="collection-editor" aria-label={strings.pages} open>
          <summary>
            {strings.pages}: {owner.name}
          </summary>
          {!pageEditable && <p class="notice">{strings.collectionReadOnly}</p>}
          <button
            disabled={pageDisabled}
            onClick={() => {
              const control = createPage(document, owner.id);
              if (control) onCommit([{ type: "addPage", control }]);
            }}
          >
            {strings.addPage}
          </button>
          <ul>
            {pages.map((page, index) => (
              <li key={page.id}>
                <button
                  aria-pressed={selectedId === page.id}
                  disabled={busy}
                  onClick={() => {
                    onActivePage(owner.id, page.id);
                    onSelect(page.id);
                  }}
                >
                  {page.caption ?? page.name}
                </button>
                <button
                  aria-label={`${strings.moveUp}: ${page.name}`}
                  disabled={pageDisabled || index === 0}
                  onClick={() =>
                    onCommit([{ type: "reorderPage", controlId: page.id, index: index - 1 }])
                  }
                >
                  ↑
                </button>
                <button
                  aria-label={`${strings.moveDown}: ${page.name}`}
                  disabled={pageDisabled || index === pages.length - 1}
                  onClick={() =>
                    onCommit([{ type: "reorderPage", controlId: page.id, index: index + 1 }])
                  }
                >
                  ↓
                </button>
                <button
                  aria-label={`${strings.deleteItem}: ${page.name}`}
                  disabled={pageDisabled}
                  onClick={() =>
                    onCommit([{ type: "removePage", controlId: page.id, cascade: true }])
                  }
                >
                  {strings.deleteItem}
                </button>
              </li>
            ))}
          </ul>
        </details>
      )}
      {selected?.type.toLowerCase() === "tabstrip" && (
        <details class="collection-editor" aria-label={strings.tabs} open>
          <summary>
            {strings.tabs}: {selected.name}
          </summary>
          {!tabEditable && <p class="notice">{strings.collectionReadOnly}</p>}
          <button
            disabled={tabDisabled || !tabTargets}
            onClick={() => {
              const names = new Set(tabs.map((t) => t.name.toLowerCase()));
              let suffix = 1;
              while (names.has(`tab${suffix}`)) suffix++;
              onCommit([
                {
                  type: "addTab",
                  controlId: selected.id,
                  tab: { name: `Tab${suffix}`, caption: `Tab${suffix}` },
                },
              ]);
            }}
          >
            {strings.addTab}
          </button>
          <ul>
            {tabs.map((tab, index) => (
              <li key={tab.name}>
                <button
                  disabled={busy}
                  aria-pressed={tabIndex === index}
                  onClick={() => setTabSelection({ owner: selected.id, name: tab.name, index })}
                >
                  {tab.caption ?? tab.name}
                </button>
                <button
                  aria-label={`${strings.moveUp}: ${tab.name}`}
                  disabled={tabDisabled || index === 0}
                  onClick={() =>
                    onCommit([
                      { type: "reorderTab", controlId: selected.id, index, toIndex: index - 1 },
                    ])
                  }
                >
                  ↑
                </button>
                <button
                  aria-label={`${strings.moveDown}: ${tab.name}`}
                  disabled={tabDisabled || index === tabs.length - 1}
                  onClick={() =>
                    onCommit([
                      { type: "reorderTab", controlId: selected.id, index, toIndex: index + 1 },
                    ])
                  }
                >
                  ↓
                </button>
                <button
                  aria-label={`${strings.deleteItem}: ${tab.name}`}
                  disabled={tabDisabled}
                  onClick={() => onCommit([{ type: "removeTab", controlId: selected.id, index }])}
                >
                  {strings.deleteItem}
                </button>
              </li>
            ))}
          </ul>
          {tabTargets?.[tabIndex] && (
            <PropertyGrid
              key={JSON.stringify([selected.id, tabs[tabIndex].name])}
              target={tabTargets[tabIndex]}
              controlId={`tab:${selected.id}:${tabIndex}`}
              draftId={JSON.stringify(["tab", selected.id, tabs[tabIndex].name])}
              title={tabs[tabIndex].name}
              version={version}
              editable={tabEditable}
              busy={busy}
              strings={strings}
              drafts={drafts}
              reply={reply}
              onCommit={(field, value) =>
                value === null || typeof value !== "object"
                  ? onCommit([
                      {
                        type: "setTabProperty",
                        controlId: selected.id,
                        index: tabIndex,
                        field,
                        value,
                      },
                    ])
                  : undefined
              }
            />
          )}
        </details>
      )}
      {selected && canEditGeometry(selected) && (
        <details class="collection-editor">
          <summary>{strings.moveToPage}</summary>
          <label>
            {strings.moveToPage}
            <select
              aria-label={strings.moveToPage}
              value=""
              disabled={!pageEditable || !parentEditable || !editable || busy || !allPages.length}
              onChange={(event) => {
                const pageId = event.currentTarget.value;
                if (!pageId) return;
                const operation = moveToPage(document, selected.id, pageId, snapping);
                if (operation) onCommit([operation]);
                else onInvalidMove();
                event.currentTarget.value = "";
              }}
            >
              <option value="">{strings.choosePage}</option>
              {allPages.map((page) => (
                <option value={page.id}>
                  {selectionPath(document, page.id)
                    .map((c) => c.name)
                    .join(" > ")}
                </option>
              ))}
            </select>
          </label>
        </details>
      )}
    </>
  );
}
