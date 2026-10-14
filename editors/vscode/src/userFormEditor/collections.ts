import type {
  DesignerDocument,
  PageOperation,
  ParentOperation,
  SemanticOperation,
} from "./protocol";
import type { NewControl } from "./toolbox";
import {
  isDesignerPage,
  orderedChildren,
  parentBounds,
  parentOrigin,
  placeGeometry,
  worldGeometry,
} from "./geometry";
import { subtreeIds } from "./toolbox";

export type ActivePages = Record<string, string>;

/** Place the owner and its initial collection as one draft/history action. */
export function initialControlOperations(
  document: DesignerDocument,
  control: NewControl,
  uuid: () => string = () => crypto.randomUUID(),
): SemanticOperation[] | undefined {
  const operations: SemanticOperation[] = [{ type: "addControl", control }];
  if (control.type !== "MultiPage") return operations;
  const allocation = {
    ...document,
    controls: [...document.controls, { ...control, approximate: false }],
  };
  for (let index = 0; index < 2; index++) {
    const page = createPage(allocation, control.id, uuid);
    if (!page) return;
    operations.push({ type: "addPage", control: page });
    allocation.controls.push({ ...page, left: 0, top: 0, width: 0, height: 0, approximate: true });
  }
  return operations;
}

// This projection is used only for rendering and hit testing, never source edits.
export function displayPages(document: DesignerDocument, active: ActivePages): DesignerDocument {
  return {
    ...document,
    controls: document.controls.map((control) => {
      if (control.type.toLowerCase() !== "multipage" || !active[control.id]) return control;
      const index = orderedChildren(document, control.id).findIndex(
        (p) => p.id === active[control.id],
      );
      return index < 0 ? control : { ...control, selectedIndex: index };
    }),
  };
}

export function createPage(
  document: DesignerDocument,
  parentId: string,
  uuid: () => string = () => crypto.randomUUID(),
): Extract<PageOperation, { type: "addPage" }>["control"] | undefined {
  const parent = document.controls.find((c) => c.id === parentId);
  if (
    !parent ||
    parent.type.toLowerCase() !== "multipage" ||
    (parent.progId && parent.progId.toLowerCase() !== "forms.multipage.1")
  )
    return;
  const names = new Set(
    [document.name, ...document.controls.map((c) => c.name)].map((n) => n.toLowerCase()),
  );
  let suffix = 1;
  while (names.has(`page${suffix}`)) suffix++;
  for (let attempt = 0; attempt < 100; attempt++) {
    const id = `control-${uuid()}`;
    if (!document.controls.some((c) => c.id === id))
      return { id, parentId, type: "Page", name: `Page${suffix}`, caption: `Page${suffix}` };
  }
}

export function moveToPage(
  document: DesignerDocument,
  controlId: string,
  pageId: string,
  snapping = false,
): ParentOperation | undefined {
  const control = document.controls.find((c) => c.id === controlId);
  const page = document.controls.find((c) => c.id === pageId);
  if (
    !control ||
    !page ||
    !isDesignerPage(document, page) ||
    control.type.toLowerCase() === "page" ||
    control.parentId === pageId ||
    subtreeIds(document.controls, controlId).has(pageId)
  )
    return;
  const world = worldGeometry(document, control);
  const origin = parentOrigin(document, pageId);
  const geometry = placeGeometry(
    { ...world, left: world.left - origin.left, top: world.top - origin.top },
    parentBounds(document, { ...control, parentId: pageId }),
    snapping,
  );
  if (!geometry) return;
  return { type: "setParent", controlId, parentId: pageId, left: geometry.left, top: geometry.top };
}
