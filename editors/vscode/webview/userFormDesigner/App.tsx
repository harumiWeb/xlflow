import { render } from "preact";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import type {
  DesignerDocument,
  HostMessage,
  WebviewMessage,
  PropertyGridData,
  SemanticOperation,
  PropertyDescriptor,
  PropertyValue,
} from "../../src/userFormEditor/protocol";
import { DesignerCanvas } from "./DesignerCanvas";
import { PropertyGrid } from "./PropertyGrid";
import { PropertyEditorHost } from "./propertyEditors/EditorHost";
import type { PictureReply } from "./propertyEditors/EditorHost";
import type { EditorSession, EditorPreview, PictureSelection } from "./propertyEditors/types";
import { objectValue } from "./propertyEditors/types";
import { metadataAppearanceDocument, previewDocument } from "./propertyEditors/appearance";
import { Toolbox } from "./Toolbox";
import { CollectionEditor } from "./CollectionEditor";
import type { ActivePages } from "../../src/userFormEditor/collections";
import { initialControlOperations } from "../../src/userFormEditor/collections";
import { createControl, controlMetadata } from "../../src/userFormEditor/toolbox";
import type { ToolboxType } from "../../src/userFormEditor/toolbox";
import type { PropertyDraftStore, PropertyReply } from "./PropertyGrid";
import { designerStrings } from "../../src/userFormEditor/protocol";
import "@vscode/codicons/dist/codicon.css";
import "./app.css";
import { Icon } from "./Icon";
import { TabOrderEditor } from "./TabOrderEditor";
import { StructureTree } from "./StructureTree";
import {
  orderedChildren,
  zoomLevels,
  selectionPath,
  isDesignerPage,
} from "../../src/userFormEditor/geometry";
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
  const [editorVersion, setEditorVersion] = useState(0);
  const [editorEpoch, setEditorEpoch] = useState(0);
  const [editor, setEditor] = useState<EditorSession>();
  const [editorPreview, setEditorPreview] = useState<EditorPreview>();
  const [pictureReply, setPictureReply] = useState<PictureReply>();
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
  const [history, setHistory] = useState({ canUndo: false, canRedo: false });
  const [draftState, setDraftState] = useState({
    dirty: false,
    conflicted: false,
    saving: false,
    error: undefined as string | undefined,
  });
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [snapping, setSnapping] = useState(false);
  const [clipboardAvailable, setClipboardAvailable] = useState(false);
  const [rightTab, setRightTab] = useState<"properties" | "structure" | "tabOrder">("properties");
  const [mobilePane, setMobilePane] = useState<"toolbox" | "properties">();
  const [revealId, setRevealId] = useState<string>();
  const canvasScroll = useRef<HTMLDivElement>(null);
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
    if (saved && canvasScroll.current) {
      canvasScroll.current.scrollLeft = saved.scrollX ?? 0;
      canvasScroll.current.scrollTop = saved.scrollY ?? 0;
    }
  }, [document]);
  useEffect(() => {
    const listener = (event: MessageEvent<HostMessage>) => {
      const message = event.data;
      if (message.type === "historyState") setHistory(message);
      if (message.type === "draftState")
        setDraftState({
          dirty: message.dirty,
          conflicted: message.conflicted,
          saving: message.saving,
          error: message.error,
        });
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
        setEditorVersion(message.propertyEditorVersion ?? 0);
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
      if (message.type === "pictureSelection") setPictureReply(message);
      if (message.type === "invalidDocument") {
        setEditorEpoch((epoch) => epoch + 1);
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
        setEditorEpoch((epoch) => epoch + 1);
        setEditorVersion(0);
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
    bridge.postMessage({ type: "ready" });
    return () => {
      window.removeEventListener("message", listener);
    };
  }, []);
  const commit = (operations: SemanticOperation[]) => {
    if (pending.current || busy || error || draftState.saving || draftState.conflicted) return;
    const operation = operations[0];
    if (operations.some((op) => op.type === "addPage") && !pageEditable) return;
    if (
      operation.type === "addControl" &&
      ((operation.control.type === "MultiPage" && !pageEditable) ||
        (operation.control.type === "TabStrip" && !tabEditable))
    )
      return;
    const property =
      operation.type === "setGeometry" ||
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
    if (operation.type === "setGeometry" && editorVersion !== 1) return;
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
      field:
        "field" in operation
          ? operation.field
          : operation.type === "setGeometry"
            ? "$geometry"
            : undefined,
      controlId:
        operation.type === "setTabProperty"
          ? `tab:${operation.controlId}:${operation.index}`
          : operation.type === "setControlProperty" || operation.type === "setGeometry"
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
    !draftState.saving &&
    !draftState.conflicted &&
    !editor &&
    (productivityEditable
      ? currentProductivitySelection.length > 0 || selectedCanDelete
      : selectedCanDelete);
  const remove = () => {
    if (!canDelete || !selectedId || pending.current) return;
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
    if (
      !document ||
      !productivityEditable ||
      busy ||
      error ||
      draftState.saving ||
      draftState.conflicted
    )
      return;
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
    if (action === "delete") {
      if (!canDelete) return false;
      remove();
      return true;
    }
    if (
      !productivityEditable ||
      busy ||
      error ||
      editor ||
      draftState.saving ||
      draftState.conflicted
    )
      return false;
    const parentId = document ? defaultPasteParent(document, selectedIds) : "";
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
    setRightTab("properties");
    setMobilePane("properties");
    window.requestAnimationFrame(() =>
      propertiesPane.current
        ?.querySelector<HTMLElement>(
          ".property-value > input, .property-value > select, .property-value-menu summary",
        )
        ?.focus(),
    );
  };
  const historyDisabled = busy || draftState.saving || draftState.conflicted || !!editor;
  const requestHistory = (action: "undo" | "redo") => {
    if (
      historyDisabled ||
      pending.current ||
      !(action === "undo" ? history.canUndo : history.canRedo)
    )
      return;
    const requestId = ++nextRequest.current;
    pending.current = { requestId };
    pendingStructural.current = true;
    setBusy(true);
    setEditError(undefined);
    setContextMenu(undefined);
    setTool(undefined);
    bridge.postMessage({ type: "history", action, requestId, version: version.current });
  };
  const selectStructure = (id?: string) => {
    if (busy || draftState.saving || !document) return;
    setTool(undefined);
    setSelectedId(id);
    const ancestors = selectionPath(document, id);
    setActivePages((current) => {
      const next = { ...current };
      for (const control of ancestors)
        if (isDesignerPage(document, control) && control.parentId)
          next[control.parentId] = control.id;
      return next;
    });
    setRevealId(id);
  };
  useEffect(() => {
    if (!revealId || !canvasScroll.current) return;
    const target = Array.from(
      canvasScroll.current.querySelectorAll<HTMLElement>(".control[data-control-id]"),
    ).find((element) => element.dataset.controlId === revealId);
    target?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    setRevealId(undefined);
  }, [revealId, activePages]);
  const selectedControl = document?.controls.find((control) => control.id === selectedId);
  // Canvas treats a new document as an external update and cancels gestures.
  // Selection, focus, and toolbar renders must preserve this object identity.
  const canvasDocument = useMemo(
    () =>
      document
        ? previewDocument(
            metadataAppearanceDocument(document, propertyGrid),
            editor && editor.version === documentVersion && editor.epoch === editorEpoch
              ? editorPreview
              : undefined,
          )
        : undefined,
    [document, propertyGrid, editor, editorPreview, documentVersion, editorEpoch],
  );
  const selectedTarget = selectedId ? propertyGrid?.controls[selectedId] : propertyGrid?.form;
  const positionFields =
    selectedTarget?.descriptors.filter(
      (descriptor) =>
        (selectedId ? ["left", "top", "width", "height"] : ["width", "height"]).includes(
          descriptor.field,
        ) &&
        (descriptor.valueType === "number" || descriptor.valueType === "integer"),
    ) ?? [];
  const canOpenEditor =
    editorVersion === 1 &&
    propertyEditable &&
    !error &&
    !busy &&
    !draftState.saving &&
    !draftState.conflicted &&
    !editor &&
    selectedIds.length <= 1;
  const openEditor = (descriptor: PropertyDescriptor) => {
    if (!canOpenEditor || pending.current || !selectedTarget) return;
    setEditor({
      descriptor,
      controlId: selectedId,
      version: documentVersion,
      epoch: editorEpoch,
      returnFocus: window.document.activeElement as HTMLElement,
      picturePreview:
        selectedControl?.picturePreview ??
        (selectedId === undefined ? document?.picturePreview : undefined),
      value: selectedTarget.values[descriptor.field]?.present
        ? selectedTarget.values[descriptor.field].value
        : undefined,
      title: descriptor.field[0].toUpperCase() + descriptor.field.slice(1),
    });
  };
  const openPosition = () => {
    if (!canOpenEditor || pending.current || !selectedTarget || !positionFields.length) return;
    setEditor({
      descriptor: {
        field: "$geometry",
        valueType: "object",
        required: false,
        nullable: false,
        children: positionFields,
      },
      returnFocus: window.document.activeElement as HTMLElement,
      controlId: selectedId,
      version: documentVersion,
      epoch: editorEpoch,
      title: "Position",
      value: Object.fromEntries(
        positionFields
          .filter((field) => selectedTarget.values[field.field]?.present)
          .map((field) => [field.field, selectedTarget.values[field.field].value]),
      ),
    });
  };
  const closeEditor = () => {
    setEditor(undefined);
    setEditorPreview(undefined);
    bridge.postMessage({ type: "pictureCancel" });
  };
  const commitEditor = (value: PropertyValue) => {
    if (!editor) return;
    if (editor.descriptor.field === "$geometry") {
      const previous = objectValue(editor.value);
      const changed = Object.fromEntries(
        Object.entries(objectValue(value)).filter(
          ([field, next]) => next !== previous[field] && typeof next === "number",
        ),
      );
      if (!Object.keys(changed).length) {
        closeEditor();
        return;
      }
      return commit([
        {
          type: "setGeometry",
          ...(editor.controlId === undefined ? {} : { controlId: editor.controlId }),
          ...changed,
        },
      ]);
    }
    return commit([
      editor.controlId === undefined
        ? { type: "setFormProperty", field: editor.descriptor.field, value }
        : {
            type: "setControlProperty",
            controlId: editor.controlId,
            field: editor.descriptor.field,
            value,
          },
    ]);
  };
  const pickPicture = () => {
    if (!editor?.controlId || editor.version !== version.current || busy) return;
    const requestId = ++nextRequest.current;
    bridge.postMessage({
      type: "picturePick",
      requestId,
      version: version.current,
      controlId: editor.controlId,
    });
    return requestId;
  };
  const commitPicture = (selection: PictureSelection) => {
    if (!editor?.controlId || pending.current || editor.version !== version.current || busy) return;
    const requestId = ++nextRequest.current;
    pending.current = { requestId, field: "picture", controlId: editor.controlId };
    setEditError(undefined);
    setBusy(true);
    bridge.postMessage({
      type: "pictureCommit",
      requestId,
      version: version.current,
      controlId: editor.controlId,
      token: selection.token,
    });
    return requestId;
  };
  const closePane = () => {
    const pane = mobilePane;
    setMobilePane(undefined);
    if (pane)
      window.document
        .querySelector<HTMLElement>(`.pane-toggle[data-pane="${pane}"]`)
        ?.focus({ preventScroll: true });
  };
  return (
    <main class="designer-app">
      <header>
        <span class="sr-only">{strings.header}</span>
        <strong class="designer-document-name">{document?.name ?? strings.header}</strong>
        <div class="view-switcher">
          <span class="active-view">{strings.designerView}</span>
          <button
            type="button"
            title={strings.openText}
            onClick={() => bridge.postMessage({ type: "openText" })}
          >
            <Icon name="go-to-file" />
            {strings.openText}
          </button>
        </div>
        <button
          class="primary-action"
          disabled={busy || draftState.saving || draftState.conflicted}
          onClick={() => bridge.postMessage({ type: "save" })}
        >
          {strings.save}
        </button>
        {draftState.dirty && <span class="notice">{strings.unsaved}</span>}
      </header>
      <div class="designer-notices">
        {draftState.conflicted && (
          <div role="alert" class="error">
            <p>{strings.sourceConflict}</p>
            <button
              onClick={() => bridge.postMessage({ type: "resolveConflict", choice: "source" })}
            >
              {strings.useSource}
            </button>
            <button
              onClick={() => bridge.postMessage({ type: "resolveConflict", choice: "designer" })}
            >
              {strings.useDesigner}
            </button>
          </div>
        )}
        {draftState.error && (
          <div role="alert" class="error">
            {draftState.error}
          </div>
        )}
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
          </>
        )}
      </div>
      {editor && (
        <PropertyEditorHost
          session={editor}
          version={documentVersion}
          epoch={editorEpoch}
          busy={busy || draftState.saving || draftState.conflicted}
          strings={strings}
          reply={reply}
          pictureReply={pictureReply}
          onClose={closeEditor}
          onPreview={setEditorPreview}
          onCommit={commitEditor}
          onPickPicture={pickPicture}
          onCommitPicture={commitPicture}
          onReleasePicture={() => bridge.postMessage({ type: "pictureCancel" })}
        />
      )}
      {document && (
        <>
          <div
            class={`designer-workspace ${mobilePane ? `show-${mobilePane}` : ""}`}
            onKeyDown={(event) => {
              if (
                mobilePane &&
                window.innerWidth < 900 &&
                event.key === "Escape" &&
                !event.defaultPrevented
              ) {
                event.preventDefault();
                closePane();
              }
            }}
          >
            <div class="toolbox-pane" id="designer-toolbox">
              <button
                class="pane-close"
                type="button"
                aria-label={strings.closePane}
                onClick={closePane}
              >
                <Icon name="close" />
              </button>
              <Toolbox
                parentEditable={parentEditable}
                pageEditable={pageEditable}
                tabEditable={tabEditable}
                selected={tool}
                disabled={!structuralEditable || busy || !!error}
                strings={strings}
                onSelect={(next) => {
                  setTool(next);
                  setMobilePane(undefined);
                  window.document
                    .querySelector<HTMLElement>(".designer-viewport")
                    ?.focus({ preventScroll: true });
                }}
              />
            </div>
            <section class="canvas-pane">
              <div class="canvas-command-bar">
                {(["undo", "redo"] as const).map((action) => (
                  <button
                    type="button"
                    class={`history-${action} icon-button`}
                    title={strings[action]}
                    aria-label={strings[action]}
                    disabled={
                      historyDisabled || !(action === "undo" ? history.canUndo : history.canRedo)
                    }
                    onClick={() => requestHistory(action)}
                  >
                    <Icon name={action} />
                  </button>
                ))}
                <button
                  class="pane-toggle"
                  data-pane="toolbox"
                  type="button"
                  aria-expanded={mobilePane === "toolbox"}
                  aria-controls="designer-toolbox"
                  onClick={() => setMobilePane(mobilePane === "toolbox" ? undefined : "toolbox")}
                >
                  <Icon name="tools" />
                  {strings.toolbox}
                </button>
                {document && productivityEditable && (
                  <ProductivityToolbar
                    document={document}
                    selectionIds={selectedIds}
                    strings={strings}
                    enabled={editable && !error}
                    structuralEnabled={structuralEditable && !error}
                    duplicateEditable={
                      !error && canPasteInto(defaultDuplicateParent(document, selectedIds))
                    }
                    clipboardAvailable={clipboardAvailable}
                    canPaste={
                      structuralEditable &&
                      !error &&
                      canPasteInto(defaultPasteParent(document, selectedIds))
                    }
                    busy={busy || draftState.saving}
                    contextMenu={contextMenu}
                    onOperation={(operation) => commit([operation])}
                    onClipboard={requestClipboard}
                    onProperties={openProperties}
                    onCloseMenu={() => setContextMenu(undefined)}
                  />
                )}
                {!productivityEditable && (
                  <button
                    type="button"
                    class="delete-control icon-button"
                    title={strings.deleteControl}
                    aria-label={strings.deleteControl}
                    disabled={!canDelete || !!tool}
                    onClick={remove}
                  >
                    <Icon name="trash" />
                  </button>
                )}
                <button
                  class="pane-toggle"
                  data-pane="properties"
                  type="button"
                  aria-expanded={mobilePane === "properties"}
                  aria-controls="designer-properties"
                  onClick={() =>
                    setMobilePane(mobilePane === "properties" ? undefined : "properties")
                  }
                >
                  <Icon name="settings-gear" />
                  {strings.properties}
                </button>
              </div>
              <div
                class="canvas-scroll"
                ref={canvasScroll}
                onScroll={(event) =>
                  bridge.setState({
                    scrollX: event.currentTarget.scrollLeft,
                    scrollY: event.currentTarget.scrollTop,
                  })
                }
              >
                <DesignerCanvas
                  document={canvasDocument!}
                  approximateBounds={strings.approximateBounds}
                  editable={editable}
                  formResizable={
                    propertyEditable &&
                    !error &&
                    ["width", "height"].every((field) =>
                      propertyGrid?.form.descriptors.some(
                        (descriptor) =>
                          descriptor.field === field && descriptor.valueType === "number",
                      ),
                    )
                  }
                  onFormResize={commit}
                  resizeWidthLabel={strings.resizeWidth}
                  resizeHeightLabel={strings.resizeHeight}
                  resizeBothLabel={strings.resizeBoth}
                  parentEditable={parentEditable && !error}
                  pageEditable={pageEditable && !error}
                  productivityEditable={productivityEditable && !error}
                  activePages={activePages}
                  onActivePageChange={(owner, page) =>
                    setActivePages((current) => ({ ...current, [owner]: page }))
                  }
                  busy={busy || draftState.saving || !!editor}
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
                    const operations = initialControlOperations(document, control);
                    if (!operations) {
                      setEditError(strings.cannotInsert);
                      return;
                    }
                    commit(operations);
                  }}
                />
              </div>
            </section>
            <aside class="properties-pane" id="designer-properties" ref={propertiesPane}>
              <button
                class="pane-close"
                type="button"
                aria-label={strings.closePane}
                onClick={closePane}
              >
                <Icon name="close" />
              </button>
              <div class="right-pane-tabs" role="tablist" aria-label={strings.properties}>
                {(["properties", "structure", "tabOrder"] as const).map((tab, index, tabs) => (
                  <button
                    id={`designer-tab-${tab}`}
                    type="button"
                    role="tab"
                    aria-selected={rightTab === tab}
                    aria-controls={`designer-panel-${tab}`}
                    tabIndex={rightTab === tab ? 0 : -1}
                    onClick={() => setRightTab(tab)}
                    onKeyDown={(event) => {
                      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
                      event.preventDefault();
                      const next =
                        event.key === "Home"
                          ? 0
                          : event.key === "End"
                            ? tabs.length - 1
                            : (index + (event.key === "ArrowRight" ? 1 : -1) + tabs.length) %
                              tabs.length;
                      setRightTab(tabs[next]);
                      propertiesPane.current
                        ?.querySelector<HTMLElement>(`#designer-tab-${tabs[next]}`)
                        ?.focus();
                    }}
                  >
                    {strings[tab]}
                  </button>
                ))}
              </div>
              <div
                class="right-pane-content"
                id="designer-panel-properties"
                role="tabpanel"
                aria-labelledby="designer-tab-properties"
                hidden={rightTab !== "properties"}
              >
                {propertyGrid &&
                (selectedId ? propertyGrid.controls[selectedId] : propertyGrid.form) ? (
                  <PropertyGrid
                    drafts={drafts.current}
                    key={selectedId === undefined ? "form" : `control:${selectedId}`}
                    target={selectedId ? propertyGrid.controls[selectedId] : propertyGrid.form}
                    controlId={selectedId}
                    title={selectedControl?.name ?? document.name}
                    targetType={selectedControl?.type ?? "UserForm"}
                    version={documentVersion}
                    editable={propertyEditable && !error}
                    busy={busy || draftState.saving}
                    strings={strings}
                    reply={reply}
                    editorEnabled={canOpenEditor}
                    onOpenEditor={openEditor}
                    onPosition={
                      editorVersion === 1 && positionFields.length > 0 && selectedIds.length <= 1
                        ? openPosition
                        : undefined
                    }
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
                <CollectionEditor
                  document={document}
                  selectedId={selectedId}
                  pageEditable={pageEditable && !error}
                  tabEditable={tabEditable && !error}
                  parentEditable={parentEditable && !error}
                  editable={editable && !error}
                  busy={busy || draftState.saving}
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
              </div>
              <div
                class="right-pane-content"
                id="designer-panel-structure"
                role="tabpanel"
                aria-labelledby="designer-tab-structure"
                hidden={rightTab !== "structure"}
              >
                <StructureTree
                  document={document}
                  selectionIds={selectedIds}
                  strings={strings}
                  busy={busy || draftState.saving || draftState.conflicted || !!editor}
                  onSelect={selectStructure}
                  onDelete={() => handleShortcut("delete")}
                />
              </div>
              <div
                class="right-pane-content"
                id="designer-panel-tabOrder"
                role="tabpanel"
                aria-labelledby="designer-tab-tabOrder"
                hidden={rightTab !== "tabOrder"}
              >
                <TabOrderEditor
                  document={document}
                  selectionIds={selectedIds}
                  propertyGrid={propertyGrid}
                  strings={strings}
                  tabEnabled={productivityEditable && propertyEditable && !error}
                  busy={busy || draftState.saving}
                  onOperation={(operation) => commit([operation])}
                />
              </div>
            </aside>
          </div>
          <footer class="designer-statusbar">
            <span aria-live="polite">
              {strings.selectionCount}: {selectedIds.length}
            </span>
            {selectedIds.length === 1 && selectedControl && (
              <span class="selection-bounds" aria-label={strings.selectionBounds}>
                x:{selectedControl.left} y:{selectedControl.top} w:{selectedControl.width} h:
                {selectedControl.height}
              </span>
            )}
            <div class="designer-toolbar">
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
          </footer>
        </>
      )}
    </main>
  );
}

render(<App />, document.getElementById("app")!);
