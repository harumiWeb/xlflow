import { render } from "preact";
import { useEffect, useRef, useState } from "preact/hooks";
import type {
  DesignerDocument,
  HostMessage,
  WebviewMessage,
  PropertyGridData,
  SemanticOperation,
} from "../../src/userFormEditor/protocol";
import { DesignerCanvas } from "./DesignerCanvas";
import { PropertyGrid } from "./PropertyGrid";
import { Toolbox } from "./Toolbox";
import { CollectionEditor } from "./CollectionEditor";
import type { ActivePages } from "../../src/userFormEditor/collections";
import { createControl, controlMetadata } from "../../src/userFormEditor/toolbox";
import type { ToolboxType } from "../../src/userFormEditor/toolbox";
import type { PropertyDraftStore, PropertyReply } from "./PropertyGrid";
import { designerStrings } from "../../src/userFormEditor/protocol";
import "./app.css";
import { orderedChildren, zoomLevels } from "../../src/userFormEditor/geometry";
import {
  defaultDuplicateParent,
  defaultPasteParent,
  siblingSelection,
} from "../../src/userFormEditor/productivity";
import { ProductivityToolbar } from "./ProductivityToolbar";
import type { ProductivityClipboardAction } from "./ProductivityToolbar";

declare function acquireVsCodeApi(): {
  postMessage(message: WebviewMessage): void;
  getState(): { scrollX?: number; scrollY?: number } | undefined;
  setState(state: { scrollX: number; scrollY: number }): void;
};
const bridge = acquireVsCodeApi();
export function App() {
  const [strings, setStrings] = useState(designerStrings);
  const [document, setDocument] = useState<DesignerDocument>();
  const [error, setError] = useState<string>();
  const [editError, setEditError] = useState<string>();
  const [propertyError, setPropertyError] = useState<string>();
  const [editable, setEditable] = useState(false);
  const [propertyEditable, setPropertyEditable] = useState(false);
  const [structuralEditable, setStructuralEditable] = useState(false);
  const [parentEditable, setParentEditable] = useState(false);
  const [pageEditable, setPageEditable] = useState(false);
  const [tabEditable, setTabEditable] = useState(false);
  const [productivityEditable, setProductivityEditable] = useState(false);
  const [activePages, setActivePages] = useState<ActivePages>({});
  const [tool, setTool] = useState<ToolboxType>();
  const [propertyGrid, setPropertyGrid] = useState<PropertyGridData>();
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const selectedId = selectedIds.at(-1);
  const setSelectedId = (id?: string) => setSelectedIds(id ? [id] : []);
  const [documentVersion, setDocumentVersion] = useState(0);
  const [reply, setReply] = useState<PropertyReply>();
  const [busy, setBusy] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [snapping, setSnapping] = useState(false);
  const [clipboardAvailable, setClipboardAvailable] = useState(false);
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number }>();
  const version = useRef(0);
  const nextRequest = useRef(0);
  const pending = useRef<PropertyReply>();
  const pendingAdd = useRef<string>();
  const pendingStructural = useRef(false);
  const pendingActive = useRef<{ owner: string; page: string }>();
  const pendingSelection = useRef<{ id: string; parentId?: string }>();
  const pendingSelectionIds = useRef<string[]>();
  const pendingClipboardVersion = useRef<number>();
  const documentAfterClipboard = useRef(false);
  const propertiesPane = useRef<HTMLDivElement>(null);
  const currentDocument = useRef<DesignerDocument>();
  const drafts = useRef<PropertyDraftStore>({ values: new Map(), replies: new Map() });
  const restored = useRef(false);
  useEffect(() => {
    if (!document || restored.current) return;
    restored.current = true;
    const saved = bridge.getState();
    if (saved) window.scrollTo(saved.scrollX ?? 0, saved.scrollY ?? 0);
  }, [document]);
  useEffect(() => {
    const listener = (event: MessageEvent<HostMessage>) => {
      const message = event.data;
      if (message.type === "localization") setStrings(message.strings);
      if (message.type === "document") {
        if (
          pending.current &&
          pending.current.field === undefined &&
          !pendingStructural.current &&
          message.version !== version.current
        )
          pending.current = undefined;
        version.current = message.version;
        currentDocument.current = message.document;
        if (
          pendingClipboardVersion.current !== undefined &&
          message.version !== pendingClipboardVersion.current
        )
          documentAfterClipboard.current = true;
        setContextMenu(undefined);
        setDocumentVersion(message.version);
        setDocument(message.document);
        setEditable(message.editable === true);
        setPropertyGrid(message.propertyGrid);
        setPropertyEditable(message.propertyEditable === true);
        setStructuralEditable(message.structuralEditable === true);
        setParentEditable(message.parentEditable === true);
        setPageEditable(message.pageEditable === true);
        setTabEditable(message.tabEditable === true);
        setProductivityEditable(message.productivityEditable === true);
        if (pendingAdd.current) {
          const added = message.document.controls.find((c) => c.id === pendingAdd.current);
          if (added?.type.toLowerCase() === "page" && added.parentId)
            setActivePages((current) => ({ ...current, [added.parentId!]: added.id }));
        }
        const selection = pendingSelection.current;
        const selectedControl =
          selection &&
          message.document.controls.find(
            (c) =>
              c.id === selection.id &&
              (selection.parentId === undefined || (c.parentId ?? "") === selection.parentId),
          );
        if (selectedControl && pendingActive.current) {
          const { owner, page } = pendingActive.current;
          setActivePages((current) => ({ ...current, [owner]: page }));
        }
        setPropertyError(message.propertyError?.message);
        const pendingIds = pendingSelectionIds.current;
        if (pendingIds !== undefined) {
          setSelectedIds(
            pendingIds.filter((id) => message.document.controls.some((c) => c.id === id)),
          );
          pendingSelectionIds.current = undefined;
        } else if (selectedControl) setSelectedIds([selectedControl.id]);
        else
          setSelectedIds((ids) => {
            const existing = ids.filter((id) =>
              message.document!.controls.some((c) => c.id === id),
            );
            if (existing.length <= 1) return existing;
            const controls = siblingSelection(message.document!, existing);
            return controls.length === existing.length ? existing : existing.slice(-1);
          });
        if (selectedControl) pendingSelection.current = undefined;
        if (!pending.current) setBusy(false);
        setError(undefined);
      }
      if (message.type === "invalidDocument") {
        setError(message.error.message);
        setEditable(false);
        setPropertyEditable(false);
        setStructuralEditable(false);
        setParentEditable(false);
        setPageEditable(false);
        setTabEditable(false);
        setProductivityEditable(false);
        setContextMenu(undefined);
        setTool(undefined);
        version.current = message.version;
        setDocumentVersion(message.version);
        setBusy(false);
      }
      if (message.type === "editingUnavailable") {
        setEditable(false);
        setPropertyEditable(false);
        setStructuralEditable(false);
        setParentEditable(false);
        setPageEditable(false);
        setTabEditable(false);
        setProductivityEditable(false);
      }
      if (message.type === "editResult" && message.requestId === pending.current?.requestId) {
        const response = { ...pending.current, requestId: message.requestId, error: message.error };
        if (response.field !== undefined) drafts.current.replies.set(response.requestId, response);
        setReply(response);
        if (!message.error && message.selectionIds !== undefined) {
          if (documentAfterClipboard.current) {
            const availableIds = message.selectionIds.filter((id) =>
              currentDocument.current?.controls.some((control) => control.id === id),
            );
            if (
              message.selectionIds.length === 0 ||
              availableIds.length === message.selectionIds.length
            ) {
              setSelectedIds(availableIds);
              pendingSelectionIds.current = undefined;
            } else pendingSelectionIds.current = message.selectionIds;
          } else pendingSelectionIds.current = message.selectionIds;
        }
        if (message.error) pendingSelectionIds.current = undefined;
        pending.current = undefined;
        if (!message.error && pendingAdd.current) {
          setSelectedId(pendingAdd.current);
          setTool(undefined);
        }
        if (!message.error && pendingActive.current) {
          const { owner, page } = pendingActive.current;
          setActivePages((current) => ({ ...current, [owner]: page }));
        }
        if (message.error) pendingSelection.current = undefined;
        pendingActive.current = undefined;
        pendingAdd.current = undefined;
        pendingStructural.current = false;
        pendingClipboardVersion.current = undefined;
        documentAfterClipboard.current = false;
        setBusy(false);
        if (message.error && message.error.code !== "editCancelled") {
          setEditError(message.error.message);
          setBusy(false);
        } else setEditError(undefined);
      }
      if (message.type === "clipboardState") setClipboardAvailable(message.available);
    };
    window.addEventListener("message", listener);
    const save = () => bridge.setState({ scrollX: window.scrollX, scrollY: window.scrollY });
    window.addEventListener("scroll", save);
    bridge.postMessage({ type: "ready" });
    return () => {
      window.removeEventListener("message", listener);
      window.removeEventListener("scroll", save);
    };
  }, []);
  const commit = (operations: SemanticOperation[]) => {
    if (pending.current || busy || error) return;
    const operation = operations[0];
    const property =
      operation.type === "setFormProperty" ||
      operation.type === "setControlProperty" ||
      operation.type === "setTabProperty" ||
      operation.type === "setTabOrder";
    const structural =
      operation.type === "addControl" ||
      operation.type === "removeControl" ||
      operation.type === "removeControls" ||
      operation.type === "pasteControls";
    const page = ["addPage", "removePage", "reorderPage"].includes(operation.type);
    const tab = ["addTab", "removeTab", "reorderTab", "setTabProperty"].includes(operation.type);
    const productivity = [
      "moveControls",
      "alignControls",
      "distributeControls",
      "sameSizeControls",
      "removeControls",
      "reorderControls",
      "setTabOrder",
      "pasteControls",
    ].includes(operation.type);
    if (productivity && !productivityEditable) return;
    if (
      page
        ? !pageEditable
        : tab
          ? !tabEditable
          : structural
            ? !structuralEditable
            : property
              ? !propertyEditable
              : !editable
    )
      return;
    if (
      (operation.type === "setParent" ||
        (operation.type === "addControl" && operation.control.parentId) ||
        (operation.type === "pasteControls" &&
          operation.parentId &&
          document?.controls
            .find((control) => control.id === operation.parentId)
            ?.type.toLowerCase() !== "page")) &&
      !parentEditable
    )
      return;
    if (
      operation.type === "pasteControls" &&
      operation.parentId &&
      document?.controls
        .find((control) => control.id === operation.parentId)
        ?.type.toLowerCase() === "page" &&
      !pageEditable
    )
      return;
    const requestId = ++nextRequest.current;
    pending.current = {
      requestId,
      field: "field" in operation ? operation.field : undefined,
      controlId:
        operation.type === "setTabProperty"
          ? `tab:${operation.controlId}:${operation.index}`
          : operation.type === "setControlProperty"
            ? operation.controlId
            : undefined,
    };
    pendingAdd.current =
      operation.type === "addControl" || operation.type === "addPage"
        ? operation.control.id
        : undefined;
    pendingSelection.current =
      operation.type === "addControl" || operation.type === "addPage"
        ? { id: operation.control.id }
        : operation.type === "setParent"
          ? { id: operation.controlId, parentId: operation.parentId }
          : undefined;
    pendingStructural.current =
      structural || page || tab || operation.type === "setParent" || productivity;
    pendingActive.current = undefined;
    if (operation.type === "addPage" && operation.control.parentId)
      pendingActive.current = { owner: operation.control.parentId, page: operation.control.id };
    setEditError(undefined);
    setBusy(true);
    bridge.postMessage({ type: "edit", requestId, version: version.current, operations });
    if (document && operation.type === "setParent") {
      const page = document.controls.find(
        (c) => c.id === operation.parentId && c.type.toLowerCase() === "page",
      );
      if (page?.parentId) pendingActive.current = { owner: page.parentId, page: page.id };
    }
    if (
      document &&
      operation.type === "setControlProperty" &&
      operation.field === "selectedIndex" &&
      typeof operation.value === "number"
    ) {
      const page = orderedChildren(document, operation.controlId)[operation.value];
      if (page?.type.toLowerCase() === "page")
        pendingActive.current = { owner: operation.controlId, page: page.id };
    }
    return requestId;
  };
  const currentProductivitySelection = document
    ? siblingSelection(document, selectedIds).map((control) => control.id)
    : [];
  const selectedCanDelete =
    !!selectedId &&
    !!document?.controls.some(
      (control) => control.id === selectedId && control.type.toLowerCase() !== "page",
    );
  const canDelete =
    structuralEditable &&
    !error &&
    !busy &&
    (productivityEditable
      ? currentProductivitySelection.length > 0 || selectedCanDelete
      : selectedCanDelete);
  const remove = () => {
    if (!canDelete || !selectedId) return;
    if (productivityEditable && currentProductivitySelection.length > 0)
      commit([{ type: "removeControls", controlIds: currentProductivitySelection, cascade: true }]);
    else commit([{ type: "removeControl", controlId: selectedId, cascade: true }]);
  };
  const canPasteInto = (parentId: string) => {
    if (!parentId) return true;
    const parent = document?.controls.find((control) => control.id === parentId);
    if (!parent) return false;
    return parentEditable && (parent.type.toLowerCase() !== "page" || pageEditable);
  };
  const requestClipboard = (
    action: ProductivityClipboardAction,
    controlIds: string[],
    parentId: string,
  ) => {
    if (!document || !productivityEditable || busy || error) return;
    if (action === "paste") {
      if (!clipboardAvailable || !structuralEditable || !canPasteInto(parentId)) return;
    } else {
      if (!siblingSelection(document, controlIds).length) return;
      if (action !== "copy" && !structuralEditable) return;
      if (action === "duplicate" && !canPasteInto(parentId)) return;
    }
    const requestId = ++nextRequest.current;
    pending.current = { requestId };
    pendingClipboardVersion.current = version.current;
    documentAfterClipboard.current = false;
    pendingStructural.current = true;
    pendingSelection.current = undefined;
    pendingAdd.current = undefined;
    pendingActive.current = undefined;
    setEditError(undefined);
    setContextMenu(undefined);
    setBusy(true);
    bridge.postMessage({
      type: "clipboard",
      requestId,
      version: version.current,
      action,
      controlIds: action === "paste" ? [] : controlIds,
      parentId,
    });
  };
  const handleShortcut = (action: "copy" | "cut" | "paste" | "duplicate" | "delete") => {
    if (!productivityEditable || busy || error) return false;
    const parentId = document ? defaultPasteParent(document, selectedIds) : "";
    if (action === "delete") {
      if (!canDelete) return false;
      remove();
      return true;
    }
    if (action === "paste") {
      if (!clipboardAvailable || !structuralEditable || !canPasteInto(parentId)) return false;
      requestClipboard(action, [], parentId);
      return true;
    }
    const ids = currentProductivitySelection;
    if (!ids.length || (action !== "copy" && !structuralEditable)) return false;
    const targetParent =
      action === "duplicate" && document ? defaultDuplicateParent(document, ids) : parentId;
    if (action === "duplicate" && !canPasteInto(targetParent)) return false;
    requestClipboard(action, ids, targetParent);
    return true;
  };
  const openProperties = () => {
    setContextMenu(undefined);
    window.requestAnimationFrame(() =>
      propertiesPane.current
        ?.querySelector<HTMLElement>(
          ".property-grid input, .property-grid select, .property-grid textarea, .property-grid button",
        )
        ?.focus(),
    );
  };
  return (
    <>
      <header>
        {strings.header}
        <button onClick={() => bridge.postMessage({ type: "openText" })}>{strings.openText}</button>
      </header>
      {document && productivityEditable && (
        <ProductivityToolbar
          document={document}
          selectionIds={selectedIds}
          propertyGrid={propertyGrid}
          strings={strings}
          enabled={editable && !error}
          structuralEnabled={structuralEditable && !error}
          duplicateEditable={!error && canPasteInto(defaultDuplicateParent(document, selectedIds))}
          tabEnabled={propertyEditable && !error}
          clipboardAvailable={clipboardAvailable}
          canPaste={
            structuralEditable && !error && canPasteInto(defaultPasteParent(document, selectedIds))
          }
          busy={busy}
          contextMenu={contextMenu}
          onOperation={(operation) => commit([operation])}
          onClipboard={requestClipboard}
          onProperties={openProperties}
          onCloseMenu={() => setContextMenu(undefined)}
        />
      )}
      <div class="designer-toolbar">
        <button
          type="button"
          class="delete-control"
          hidden={productivityEditable}
          disabled={!canDelete || !!tool}
          onClick={remove}
        >
          {strings.deleteControl}
        </button>
        <label class="grid-toggle">
          <input
            type="checkbox"
            checked={showGrid}
            onChange={(event) => setShowGrid(event.currentTarget.checked)}
          />
          {strings.grid}
        </label>
        <label class="snap-toggle">
          <input
            type="checkbox"
            checked={snapping}
            disabled={busy}
            onChange={(event) => setSnapping(event.currentTarget.checked)}
          />
          {strings.snap}
        </label>
        <label>
          {strings.zoom}
          <select
            value={zoom}
            disabled={busy}
            onChange={(event) => setZoom(Number(event.currentTarget.value))}
          >
            {zoomLevels.map((level) => (
              <option value={level}>{Math.round(level * 100)}%</option>
            ))}
          </select>
        </label>
      </div>
      {(error || editError || propertyError) && (
        <div role="alert" class="error">
          {error || editError || propertyError}
          {document && error && <p>{strings.lastValid}</p>}
        </div>
      )}
      {!document && !error && <p>{strings.loading}</p>}
      {document &&
        !editable &&
        !propertyEditable &&
        !structuralEditable &&
        !pageEditable &&
        !tabEditable &&
        !error &&
        !busy && <p class="notice">{strings.readOnly}</p>}
      {document && (
        <>
          {(document.approximate || document.controls.some((c) => c.approximate)) && (
            <p class="notice">{strings.approximate}</p>
          )}
          {document.warnings.map((warning, index) => (
            <p class="notice" key={index}>
              {warning}
            </p>
          ))}
          <div class="designer-workspace">
            <Toolbox
              parentEditable={parentEditable}
              pageEditable={pageEditable}
              selected={tool}
              disabled={!structuralEditable || busy || !!error}
              strings={strings}
              onSelect={setTool}
            />
            <DesignerCanvas
              document={document}
              approximateBounds={strings.approximateBounds}
              editable={editable}
              parentEditable={parentEditable && !error}
              pageEditable={pageEditable && !error}
              productivityEditable={productivityEditable && !error}
              activePages={activePages}
              onActivePageChange={(owner, page) =>
                setActivePages((current) => ({ ...current, [owner]: page }))
              }
              busy={busy}
              zoom={zoom}
              showGrid={showGrid}
              snapping={snapping}
              cannotMove={strings.cannotMove}
              hierarchy={strings.hierarchy}
              invalidDrop={strings.invalidDrop}
              dropTargetLabel={strings.dropTarget}
              placementSize={controlMetadata.find((m) => m.type === tool)}
              selectedId={selectedId}
              selectedIds={selectedIds}
              onSelectionChangeMany={setSelectedIds}
              onCommit={commit}
              onContextMenu={(id, x, y) => {
                if (productivityEditable) setContextMenu({ x, y });
                else setSelectedId(id);
              }}
              onShortcut={handleShortcut}
              placing={!!tool && structuralEditable && !error}
              onCancelPlacement={() => setTool(undefined)}
              onDelete={canDelete ? remove : undefined}
              onPlace={(left, top, parentId) => {
                if (!tool) return;
                const control = createControl(
                  document,
                  tool,
                  left,
                  top,
                  snapping,
                  undefined,
                  parentId,
                  pageEditable,
                );
                if (!control) {
                  setEditError(parentId ? strings.cannotInsertContainer : strings.cannotInsert);
                  return;
                }
                commit([{ type: "addControl", control }]);
              }}
            />
            <div class="properties-pane" ref={propertiesPane}>
              <CollectionEditor
                document={document}
                selectedId={selectedId}
                pageEditable={pageEditable && !error}
                tabEditable={tabEditable && !error}
                parentEditable={parentEditable && !error}
                editable={editable && !error}
                busy={busy}
                strings={strings}
                onSelect={setSelectedId}
                onActivePage={(owner, page) =>
                  setActivePages((current) => ({ ...current, [owner]: page }))
                }
                onCommit={commit}
                propertyGrid={propertyGrid}
                drafts={drafts.current}
                reply={reply}
                version={documentVersion}
                snapping={snapping}
                onInvalidMove={() => setEditError(strings.invalidDrop)}
              />
              {propertyGrid &&
              (selectedId ? propertyGrid.controls[selectedId] : propertyGrid.form) ? (
                <PropertyGrid
                  drafts={drafts.current}
                  key={selectedId === undefined ? "form" : `control:${selectedId}`}
                  target={selectedId ? propertyGrid.controls[selectedId] : propertyGrid.form}
                  controlId={selectedId}
                  title={document.controls.find((c) => c.id === selectedId)?.name ?? document.name}
                  version={documentVersion}
                  editable={propertyEditable && !error}
                  busy={busy}
                  strings={strings}
                  reply={reply}
                  onCommit={(field, value) =>
                    commit([
                      selectedId
                        ? { type: "setControlProperty", controlId: selectedId, field, value }
                        : { type: "setFormProperty", field, value },
                    ])
                  }
                />
              ) : (
                <aside class="property-grid">
                  <h2>{strings.properties}</h2>
                  <p class="notice">{strings.propertyReadOnly}</p>
                </aside>
              )}
            </div>
          </div>
        </>
      )}
    </>
  );
}

render(<App />, document.getElementById("app")!);
