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
        setDocument(message.document);
        setError(undefined);
      }
      if (message.type === "invalidDocument") setError(message.error.message);
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
      {error && (
        <div role="alert" class="error">
          {error}
          {document && <p>{strings.lastValid}</p>}
        </div>
      )}
      {!document && !error && <p>{strings.loading}</p>}
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
          <DesignerCanvas document={document} approximateBounds={strings.approximateBounds} />
        </>
      )}
    </>
  );
}

render(<App />, document.getElementById("app")!);
