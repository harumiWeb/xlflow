import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { designerDocument } from "../src/userFormEditor/model";
import { problemForRequest, problemSourceRange } from "../src/userFormEditor/problems";
import { PreviewSynchronizer } from "../src/userFormEditor/document";
import { designerStrings, isWebviewMessage } from "../src/userFormEditor/protocol";
import type {
  DesignerDocument,
  DesignerProblem,
  DesignerStrings,
  HostMessage,
  PreviewResult,
  WebviewMessage,
} from "../src/userFormEditor/protocol";
import { ProblemsPanel } from "../webview/userFormDesigner/ProblemsPanel";

const { JSDOM } = require("jsdom");
const dom = new JSDOM("<!doctype html><html><body><div id=app></div></body></html>", {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });

const root = document.getElementById("app")!;
const issue: DesignerProblem = {
  code: "UFV013",
  severity: "warning",
  message: "Persisted warning message.",
  field: "warnings[0].message",
  suggestion: "Review this persisted warning.",
  support: "snapshot_only",
  range: {
    start: { line: 4, character: 12 },
    end: { line: 4, character: 35 },
  },
};

async function problemPanelCollapsesNavigatesAndRestoresFocus() {
  const opened: [number, number][] = [];
  const strings: DesignerStrings = {
    ...designerStrings,
    problems: "Localized Problems",
    closeProblems: "Localized Close Problems",
    problemSeverityWarning: "Localized Warning",
  };
  await act(() =>
    render(
      <ProblemsPanel
        problems={[issue]}
        version={12}
        strings={strings}
        onOpenProblem={(version, index) => opened.push([version, index])}
      />,
      root,
    ),
  );

  const toggle = root.querySelector<HTMLButtonElement>("#designer-problems-toggle")!;
  const panel = root.querySelector<HTMLElement>("#designer-problems-panel")!;
  assert.equal(toggle.textContent?.trim(), "Localized Problems 1");
  assert.equal(toggle.getAttribute("aria-expanded"), "false");
  assert.match(toggle.getAttribute("aria-label") ?? "", /Localized Warning/);
  assert.ok(toggle.querySelector(".codicon-warning"), "the collapsed button shows severity");
  assert.equal(panel.hidden, true, "the Problems panel starts collapsed");

  await act(() => toggle.click());
  assert.equal(panel.hidden, false);
  assert.match(panel.textContent ?? "", /Localized Warning/);
  assert.match(panel.textContent ?? "", /warnings\[0\]\.message/);
  const item = panel.querySelector<HTMLButtonElement>(".designer-problem")!;
  item.focus();
  const outside = document.createElement("button");
  document.body.append(outside);
  const outsideEscape = new dom.window.KeyboardEvent("keydown", {
    key: "Escape",
    bubbles: true,
    cancelable: true,
  });
  await act(() => outside.dispatchEvent(outsideEscape));
  assert.equal(panel.hidden, false, "Escape outside the panel remains available to other UI");
  assert.equal(outsideEscape.defaultPrevented, false);
  outside.remove();
  const panelEscape = new dom.window.KeyboardEvent("keydown", {
    key: "Escape",
    bubbles: true,
    cancelable: true,
  });
  await act(() => item.dispatchEvent(panelEscape));
  assert.equal(panel.hidden, true, "Escape closes the panel");
  assert.equal(panelEscape.defaultPrevented, true);
  assert.equal(document.activeElement, toggle, "closing restores focus to the Problems button");

  await act(() => toggle.click());
  await act(() => item.click());
  assert.deepEqual(opened, [[12, 0]], "the host resolves the source issue by version and index");
  await act(() => render(null, root));
}

async function structuredIssuesSurviveTheModelAndPreviewProtocol() {
  const documentModel = designerDocument({ form: { name: "MainForm" }, controls: [] }, [issue]);
  assert.deepEqual(documentModel.warnings, [issue]);
  assert.notEqual(documentModel.warnings[0], issue, "the model owns a shallow issue copy");

  let delivered: unknown;
  const sync = new PreviewSynchronizer();
  const preview: PreviewResult = {
    version: 7,
    document: { form: { name: "MainForm" }, controls: [] },
    warnings: [issue],
  };
  await sync.update(
    7,
    async () => preview,
    () => 7,
    (message) => (delivered = message),
  );
  assert.equal(
    (delivered as { document: { warnings: DesignerProblem[] } }).document.warnings[0].field,
    issue.field,
  );
  assert.deepEqual(
    (delivered as { document: { warnings: DesignerProblem[] } }).document.warnings[0].range,
    issue.range,
  );
}

async function navigationRequiresTheCurrentIssueAndSourceRevision() {
  const snapshot = {
    version: 9,
    document: { warnings: [issue] },
  };
  assert.equal(problemForRequest(snapshot, 9, 9, 0), issue);
  assert.equal(
    problemForRequest(snapshot, 10, 9, 0),
    undefined,
    "stale draft versions are rejected",
  );
  assert.equal(
    problemForRequest(snapshot, 9, 9, 1),
    undefined,
    "unknown problem indexes are rejected",
  );

  let sourceRequests = 0;
  const sameTextRange = await problemSourceRange(
    issue,
    0,
    "one\ntwo\nthree\nfour\n012345678901234567890123456789012345",
    "one\ntwo\nthree\nfour\n012345678901234567890123456789012345",
    3,
    async () => {
      sourceRequests++;
      return { version: 3 };
    },
  );
  assert.deepEqual(sameTextRange, issue.range);
  assert.equal(sourceRequests, 0);

  const sourceIssue = {
    ...issue,
    range: { start: { line: 1, character: 3 }, end: { line: 1, character: 9 } },
  };
  const draftTextRange = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 5, warnings: [sourceIssue] }),
  );
  assert.deepEqual(
    draftTextRange,
    sourceIssue.range,
    "a draft range is replaced by the matching fresh LSP range for the source buffer",
  );

  const staleSourceRange = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 4, warnings: [sourceIssue] }),
  );
  assert.equal(staleSourceRange, undefined, "a mismatched source preview version cannot navigate");

  const unrelated: DesignerProblem = { ...sourceIssue, code: "OTHER" };
  const missingMatch = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 5, warnings: [unrelated] }),
  );
  assert.equal(
    missingMatch,
    undefined,
    "a draft-only issue cannot select an unrelated source range",
  );

  const differentField: DesignerProblem = { ...sourceIssue, field: "controls[0].caption" };
  const fieldMismatch = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 5, warnings: [differentField] }),
  );
  assert.equal(fieldMismatch, undefined, "a field mismatch cannot bind to a different issue");

  const missingField: DesignerProblem = { ...sourceIssue, field: undefined };
  const missingIdentity = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 5, warnings: [missingField] }),
  );
  assert.equal(missingIdentity, undefined, "missing core identity cannot be treated as a match");

  const legacySeverity: DesignerProblem = { ...sourceIssue, severity: undefined };
  const legacySeverityMatch = await problemSourceRange(
    issue,
    0,
    "draft text has shifted the warning",
    "source\nold warning here",
    5,
    async () => ({ version: 5, warnings: [legacySeverity] }),
  );
  assert.deepEqual(
    legacySeverityMatch,
    sourceIssue.range,
    "older LSP preview versions may omit severity while retaining core identity",
  );
}

function webviewCanOnlyRequestAnIndexedVersionedProblem() {
  assert.equal(isWebviewMessage({ type: "openProblem", version: 4, index: 1 }), true);
  assert.equal(isWebviewMessage({ type: "openProblem", version: 4, index: -1 }), false);
  assert.equal(
    isWebviewMessage({ type: "openProblem", version: 4, index: 1, range: issue.range }),
    false,
    "the webview cannot provide an arbitrary navigation range",
  );
}

async function designerKeepsProblemCountCollapsedPostsCurrentIdentityAndRetainsErrors() {
  const posted: WebviewMessage[] = [];
  Object.assign(globalThis, {
    acquireVsCodeApi: () => ({
      postMessage: (message: WebviewMessage) => posted.push(message),
      getState: () => undefined,
      setState: () => {},
    }),
  });
  await act(async () => {
    require("../webview/userFormDesigner/App");
    await Promise.resolve();
  });

  const send = async (message: HostMessage) => {
    await act(() =>
      window.dispatchEvent(new dom.window.MessageEvent("message", { data: message })),
    );
  };
  const strings: DesignerStrings = { ...designerStrings };
  await send({ type: "localization", strings });
  await send({
    type: "draftState",
    dirty: true,
    conflicted: false,
    saving: false,
    error: "Draft save failed; retry.",
  });

  const first: DesignerDocument = designerDocument({ form: { name: "MainForm" }, controls: [] }, [
    issue,
  ]);
  await send({ type: "document", version: 20, document: first });
  const toggle = root.querySelector<HTMLButtonElement>("#designer-problems-toggle")!;
  assert.equal(toggle.getAttribute("aria-expanded"), "false", "the integrated panel starts closed");
  assert.equal(toggle.textContent?.trim(), "Problems 1");
  assert.match(root.querySelector(".designer-notices")?.textContent ?? "", /Draft save failed/);

  const second: DesignerDocument = {
    ...first,
    warnings: [issue, { ...issue, code: "UFV020", severity: "error", field: "controls[0].tabs" }],
  };
  await send({ type: "document", version: 21, document: second });
  assert.equal(toggle.textContent?.trim(), "Problems 2");
  assert.equal(
    toggle.getAttribute("aria-expanded"),
    "false",
    "count changes do not open the panel",
  );
  assert.match(toggle.getAttribute("aria-label") ?? "", /Error/);
  assert.ok(toggle.querySelector(".codicon-error"), "the highest severity is shown on the button");

  await act(() => toggle.click());
  const firstProblem = root.querySelector<HTMLButtonElement>(".designer-problem")!;
  await act(() => firstProblem.click());
  assert.deepEqual(posted.filter((message) => message.type === "openProblem").at(-1), {
    type: "openProblem",
    version: 21,
    index: 0,
  });

  await send({
    type: "draftState",
    dirty: true,
    conflicted: true,
    saving: false,
    error: "Draft conflict retained.",
  });
  const notices = root.querySelector(".designer-notices")!;
  assert.match(notices.textContent ?? "", /conflict/i);
  assert.match(notices.textContent ?? "", /Draft conflict retained/);
  assert.ok(notices.textContent?.includes(strings.useSource));
  assert.ok(notices.textContent?.includes(strings.useDesigner));
  await act(() => render(null, root));
}

async function run() {
  await problemPanelCollapsesNavigatesAndRestoresFocus();
  await structuredIssuesSurviveTheModelAndPreviewProtocol();
  await navigationRequiresTheCurrentIssueAndSourceRevision();
  webviewCanOnlyRequestAnIndexedVersionedProblem();
  await designerKeepsProblemCountCollapsedPostsCurrentIdentityAndRetainsErrors();
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
