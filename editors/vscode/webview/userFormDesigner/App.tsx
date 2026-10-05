import { render } from "preact";
import { useEffect, useRef, useState } from "preact/hooks";
import type {
  DesignerDocument,
  HostMessage,
  WebviewMessage,
} from "../../src/userFormEditor/protocol";
import { DesignerCanvas } from "./DesignerCanvas";
import { designerStrings } from "../../src/userFormEditor/protocol";
import "./app.css";
import { zoomLevels } from "../../src/userFormEditor/geometry";

declare function acquireVsCodeApi(): {
  postMessage(message: WebviewMessage): void;
  getState(): { scrollX?: number; scrollY?: number } | undefined;
  setState(state: { scrollX: number; scrollY: number }): void;
};
const bridge = acquireVsCodeApi();
function App() {
  const [strings, setStrings] = useState(designerStrings);
  const [document, setDocument] = useState<DesignerDocument>();
  const [error, setError] = useState<string>();
  const [editError, setEditError] = useState<string>();
  const [editable, setEditable] = useState(false);
  const [busy, setBusy] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [snapping, setSnapping] = useState(false);
  const version = useRef(0);
  const nextRequest = useRef(0);
  const pending = useRef<number>();
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
        version.current = message.version;
        setDocument(message.document);
        setEditable(message.editable === true);
        setBusy(false);
        setError(undefined);
      }
      if (message.type === "invalidDocument") {
        setError(message.error.message);
        setEditable(false);
        setBusy(false);
      }
      if (message.type === "editingUnavailable") setEditable(false);
      if (message.type === "editResult" && message.requestId === pending.current) {
        pending.current = undefined;
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
      {(error || editError) && (
        <div role="alert" class="error">
          {error || editError}
          {document && error && <p>{strings.lastValid}</p>}
        </div>
      )}
      {!document && !error && <p>{strings.loading}</p>}
      {document && !editable && !error && !busy && <p class="notice">{strings.readOnly}</p>}
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
          <DesignerCanvas
            document={document}
            approximateBounds={strings.approximateBounds}
            editable={editable}
            busy={busy}
            zoom={zoom}
            showGrid={showGrid}
            snapping={snapping}
            cannotMove={strings.cannotMove}
            onCommit={(operations) => {
              if (!editable || busy) return;
              const requestId = ++nextRequest.current;
              pending.current = requestId;
              setEditError(undefined);
              setBusy(true);
              bridge.postMessage({ type: "edit", requestId, version: version.current, operations });
            }}
          />
        </>
      )}
    </>
  );
}

render(<App />, document.getElementById("app")!);
