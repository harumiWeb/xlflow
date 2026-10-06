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
import type { PropertyDraftStore, PropertyReply } from "./PropertyGrid";
import { designerStrings } from "../../src/userFormEditor/protocol";
import "./app.css";
import { zoomLevels } from "../../src/userFormEditor/geometry";

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
  const [propertyGrid, setPropertyGrid] = useState<PropertyGridData>();
  const [selectedId, setSelectedId] = useState<string>();
  const [documentVersion, setDocumentVersion] = useState(0);
  const [reply, setReply] = useState<PropertyReply>();
  const [busy, setBusy] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [snapping, setSnapping] = useState(false);
  const version = useRef(0);
  const nextRequest = useRef(0);
  const pending = useRef<PropertyReply>();
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
          message.version !== version.current
        )
          pending.current = undefined;
        version.current = message.version;
        setDocumentVersion(message.version);
        setDocument(message.document);
        setEditable(message.editable === true);
        setPropertyGrid(message.propertyGrid);
        setPropertyEditable(message.propertyEditable === true);
        setPropertyError(message.propertyError?.message);
        setSelectedId((id) =>
          message.document.controls.some((c) => c.id === id) ? id : undefined,
        );
        if (!pending.current) setBusy(false);
        setError(undefined);
      }
      if (message.type === "invalidDocument") {
        setError(message.error.message);
        setEditable(false);
        setPropertyEditable(false);
        version.current = message.version;
        setDocumentVersion(message.version);
        setBusy(false);
      }
      if (message.type === "editingUnavailable") {
        setEditable(false);
        setPropertyEditable(false);
      }
      if (message.type === "editResult" && message.requestId === pending.current?.requestId) {
        const response = { ...pending.current, requestId: message.requestId, error: message.error };
        if (response.field !== undefined) drafts.current.replies.set(response.requestId, response);
        setReply(response);
        pending.current = undefined;
        setBusy(false);
        if (message.error) {
          setEditError(message.error.message);
          setBusy(false);
        } else setEditError(undefined);
      }
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
      operation.type === "setFormProperty" || operation.type === "setControlProperty";
    if (property ? !propertyEditable : !editable) return;
    const requestId = ++nextRequest.current;
    pending.current = {
      requestId,
      field: property ? operation.field : undefined,
      controlId: operation.type === "setControlProperty" ? operation.controlId : undefined,
    };
    setEditError(undefined);
    setBusy(true);
    bridge.postMessage({ type: "edit", requestId, version: version.current, operations });
    return requestId;
  };
  return (
    <>
      <header>
        {strings.header}
        <button onClick={() => bridge.postMessage({ type: "openText" })}>{strings.openText}</button>
      </header>
      <div class="designer-toolbar">
        <label>
          <input
            type="checkbox"
            checked={showGrid}
            onChange={(event) => setShowGrid(event.currentTarget.checked)}
          />
          {strings.grid}
        </label>
        <label>
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
      {document && !editable && !propertyEditable && !error && !busy && (
        <p class="notice">{strings.readOnly}</p>
      )}
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
            <DesignerCanvas
              document={document}
              approximateBounds={strings.approximateBounds}
              editable={editable}
              busy={busy}
              zoom={zoom}
              showGrid={showGrid}
              snapping={snapping}
              cannotMove={strings.cannotMove}
              selectedId={selectedId}
              onSelectionChange={setSelectedId}
              onCommit={commit}
            />
            {propertyGrid &&
            (selectedId ? propertyGrid.controls[selectedId] : propertyGrid.form) ? (
              <PropertyGrid
                drafts={drafts.current}
                key={selectedId ?? "form"}
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
        </>
      )}
    </>
  );
}

render(<App />, document.getElementById("app")!);
