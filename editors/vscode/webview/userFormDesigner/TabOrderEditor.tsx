import { useEffect, useMemo, useState } from "preact/hooks";
import { isDesignerFrame, isDesignerPage } from "../../src/userFormEditor/geometry";
import type {
  DesignerDocument,
  DesignerStrings,
  PropertyGridData,
  SemanticOperation,
} from "../../src/userFormEditor/protocol";
import {
  eligibleTabOrder,
  siblingSelection,
  tabOrderOperation,
} from "../../src/userFormEditor/productivity";

export function TabOrderEditor({
  document,
  selectionIds,
  propertyGrid,
  strings,
  tabEnabled,
  busy,
  onOperation,
}: {
  document: DesignerDocument;
  selectionIds: string[];
  propertyGrid?: PropertyGridData;
  strings: DesignerStrings;
  tabEnabled: boolean;
  busy: boolean;
  onOperation: (operation: SemanticOperation) => void;
}) {
  const controls = siblingSelection(document, selectionIds);
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
  return (
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
                ? (document.controls.find((control) => control.id === parentId)?.name ?? parentId)
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
