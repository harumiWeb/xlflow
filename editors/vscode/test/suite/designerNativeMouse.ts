import { get as httpGet } from "http";

type CdpSocket = {
  addEventListener(
    type: "open" | "error" | "close",
    listener: () => void,
    options?: { once?: boolean },
  ): void;
  addEventListener(type: "message", listener: (event: { data: unknown }) => void): void;
  send(data: string): void;
  close(): void;
};

type CdpFrame = { id: string; parentId?: string };
type FrameTreeNode = { frame: CdpFrame; childFrames?: FrameTreeNode[] };

class CdpClient {
  private nextId = 1;
  private readonly pending = new Map<
    number,
    { resolve(value: unknown): void; reject(error: Error): void; timer: NodeJS.Timeout }
  >();

  private constructor(private readonly socket: CdpSocket) {
    socket.addEventListener("message", (event) => this.receive(String(event.data)));
    socket.addEventListener("close", () => this.fail(new Error("CDP connection closed")));
  }

  static async connect(url: string): Promise<CdpClient> {
    const WebSocketConstructor = (
      globalThis as unknown as { WebSocket?: new (url: string) => CdpSocket }
    ).WebSocket;
    if (!WebSocketConstructor)
      throw new Error("Node WebSocket support is required for the CDP test");
    const socket = new WebSocketConstructor(url);
    await new Promise<void>((resolve, reject) => {
      socket.addEventListener("open", resolve, { once: true });
      socket.addEventListener(
        "error",
        () => reject(new Error("Could not connect to VS Code CDP")),
        {
          once: true,
        },
      );
    });
    return new CdpClient(socket);
  }

  send<T>(method: string, params: Record<string, unknown> = {}): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP command timed out: ${method}`));
      }, 5000);
      this.pending.set(id, {
        resolve: (value) => resolve(value as T),
        reject,
        timer,
      });
      try {
        this.socket.send(JSON.stringify({ id, method, params }));
      } catch (error) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  close(): void {
    this.socket.close();
    this.fail(new Error("CDP connection closed"));
  }

  private receive(data: string): void {
    const message = JSON.parse(data) as {
      id?: number;
      result?: unknown;
      error?: { message?: string };
    };
    if (message.id === undefined) return;
    const pending = this.pending.get(message.id);
    if (!pending) return;
    clearTimeout(pending.timer);
    this.pending.delete(message.id);
    if (message.error) pending.reject(new Error(message.error.message ?? "CDP command failed"));
    else pending.resolve(message.result);
  }

  private fail(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer);
      pending.reject(error);
    }
    this.pending.clear();
  }
}

type DesignerPoint = { x: number; y: number; hitControlId?: string; hitInForm: boolean };

export class NativeDesignerMouse {
  private constructor(
    private readonly cdp: CdpClient,
    private readonly frame: CdpFrame,
    private readonly frameTree: Map<string, CdpFrame>,
    private readonly contexts: Map<string, number>,
  ) {}

  static async connect(): Promise<NativeDesignerMouse> {
    const port = Number(process.env.XLFLOW_VSCODE_TEST_CDP_PORT);
    if (!Number.isInteger(port) || port < 1)
      throw new Error("The VS Code test CDP port is missing");
    const deadline = Date.now() + 15000;
    let lastError: unknown;
    while (Date.now() < deadline) {
      try {
        const targets = await getCdpTargets(port);
        for (const target of targets.filter((item) =>
          ["page", "iframe"].includes(item.type ?? ""),
        )) {
          if (!target.webSocketDebuggerUrl) continue;
          let cdp: CdpClient | undefined;
          try {
            cdp = await CdpClient.connect(target.webSocketDebuggerUrl);
            await Promise.all([
              cdp.send("Page.enable"),
              cdp.send("Runtime.enable"),
              cdp.send("DOM.enable"),
            ]);
            const { frameTree } = await cdp.send<{ frameTree: FrameTreeNode }>("Page.getFrameTree");
            const frames = flattenFrames(frameTree);
            const frameTreeById = new Map(frames.map((frame) => [frame.id, frame]));
            const contexts = new Map<string, number>();
            for (const frame of frames) {
              try {
                const context = await cdp.send<{ executionContextId: number }>(
                  "Page.createIsolatedWorld",
                  { frameId: frame.id, worldName: "xlflow-native-input-test" },
                );
                contexts.set(frame.id, context.executionContextId);
              } catch {
                // A frame may disappear while VS Code is refreshing its Webview.
              }
            }
            for (const frame of frames) {
              const contextId = contexts.get(frame.id);
              if (!contextId) continue;
              const result = await cdp.send<{
                result?: { value?: boolean };
              }>("Runtime.evaluate", {
                contextId,
                expression:
                  "Boolean(document.querySelector('.designer-viewport') && document.querySelector('.form'))",
                returnByValue: true,
              });
              if (result.result?.value !== true) continue;
              return new NativeDesignerMouse(cdp, frame, frameTreeById, contexts);
            }
            cdp.close();
          } catch (error) {
            cdp?.close();
            lastError = error;
          }
        }
      } catch (error) {
        lastError = error;
      }
      await new Promise((resolve) => setTimeout(resolve, 150));
    }
    const detail =
      lastError instanceof Error
        ? `${lastError.name}: ${lastError.message}${lastError.cause ? `; cause=${String(lastError.cause)}` : ""}`
        : String(lastError ?? "no page target");
    throw new Error(
      `Could not find the Designer Webview through CDP at 127.0.0.1:${port}: ${detail}`,
    );
  }

  async point(target: "control" | "form"): Promise<DesignerPoint> {
    const contextId = this.contexts.get(this.frame.id);
    if (!contextId) throw new Error("Designer Webview frame has no execution context");
    const expression = `(() => {
      const form = document.querySelector('.form');
      const target = ${JSON.stringify(target)} === 'control'
        ? document.querySelector('.control[data-control-id="label"]')
        : form;
      if (!form || !target) return null;
      const bounds = target.getBoundingClientRect();
      const formBounds = form.getBoundingClientRect();
      const x = ${JSON.stringify(target)} === 'control' ? (bounds.left + bounds.right) / 2 : formBounds.left + formBounds.width * 0.8;
      const y = ${JSON.stringify(target)} === 'control' ? (bounds.top + bounds.bottom) / 2 : formBounds.top + formBounds.height * 0.8;
      const hit = document.elementFromPoint(x, y);
      return {
        x, y,
        hitControlId: hit?.closest('[data-control-id]')?.dataset.controlId,
        hitInForm: !!hit && (hit === form || form.contains(hit))
      };
    })()`.replaceAll("\n", " ");
    const result = await this.cdp.send<{
      result?: { value?: DesignerPoint | null };
      exceptionDetails?: { text?: string };
    }>("Runtime.evaluate", { contextId, expression, returnByValue: true });
    if (result.exceptionDetails)
      throw new Error(result.exceptionDetails.text ?? "Could not measure Designer target");
    const point = result.result?.value;
    if (!point) throw new Error(`Could not find the Designer ${target} target`);
    const offset = await this.frameOffset();
    return { ...point, x: point.x + offset.x, y: point.y + offset.y };
  }

  async doubleClick(target: "control" | "form"): Promise<DesignerPoint> {
    const point = await this.point(target);
    await this.mouse("mouseMoved", point.x, point.y, "none", 0);
    for (const clickCount of [1, 2]) {
      await this.mouse("mousePressed", point.x, point.y, "left", 1, clickCount);
      await this.mouse("mouseReleased", point.x, point.y, "left", 0, clickCount);
      if (clickCount === 1) await new Promise((resolve) => setTimeout(resolve, 80));
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
    return point;
  }

  async clickSelector(selector: string): Promise<void> {
    const contextId = this.contexts.get(this.frame.id);
    const result = await this.cdp.send<{
      result?: { value?: { x: number; y: number; disabled: boolean; hit: boolean } };
    }>("Runtime.evaluate", {
      contextId,
      expression: `(() => {
        const target = document.querySelector(${JSON.stringify(selector)});
        if (!target) return null;
        const bounds = target.getBoundingClientRect();
        const x = (bounds.left + bounds.right) / 2;
        const y = (bounds.top + bounds.bottom) / 2;
        return { x, y, disabled: target.disabled === true,
          hit: target.contains(document.elementFromPoint(x, y)) };
      })()`,
      returnByValue: true,
    });
    const point = result.result?.value;
    if (!point || point.disabled || !point.hit)
      throw new Error(`Designer target is not clickable: ${selector}; ${JSON.stringify(point)}`);
    const offset = await this.frameOffset();
    await this.mouse("mouseMoved", point.x + offset.x, point.y + offset.y, "none", 0);
    await this.mouse("mousePressed", point.x + offset.x, point.y + offset.y, "left", 1, 1);
    await this.mouse("mouseReleased", point.x + offset.x, point.y + offset.y, "left", 0, 1);
    await this.cdp.send("Runtime.evaluate", {
      contextId,
      expression:
        "new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))",
      awaitPromise: true,
    });
  }

  async isVisible(selector: string): Promise<boolean> {
    const result = await this.cdp.send<{ result?: { value?: boolean } }>("Runtime.evaluate", {
      contextId: this.contexts.get(this.frame.id),
      expression: `Boolean(document.querySelector(${JSON.stringify(selector)})?.getClientRects().length)`,
      returnByValue: true,
    });
    return result.result?.value === true;
  }

  async replaceText(selector: string, text: string): Promise<void> {
    await this.clickSelector(selector);
    await this.cdp.send("Runtime.evaluate", {
      contextId: this.contexts.get(this.frame.id),
      expression: `document.querySelector(${JSON.stringify(selector)}).select()`,
    });
    await this.cdp.send("Input.insertText", { text });
  }

  async dragControl(): Promise<void> {
    const point = await this.point("control");
    await this.mouse("mouseMoved", point.x, point.y, "none", 0);
    await this.mouse("mousePressed", point.x, point.y, "left", 1, 1);
    await this.mouse("mouseMoved", point.x + 28, point.y + 20, "left", 1);
    await this.mouse("mouseReleased", point.x + 28, point.y + 20, "left", 0, 1);
    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  close(): void {
    this.cdp.close();
  }

  private async mouse(
    type: "mouseMoved" | "mousePressed" | "mouseReleased",
    x: number,
    y: number,
    button: "none" | "left",
    buttons: number,
    clickCount?: number,
  ): Promise<void> {
    await this.cdp.send("Input.dispatchMouseEvent", {
      type,
      x,
      y,
      button,
      buttons,
      ...(clickCount === undefined ? {} : { clickCount }),
    });
  }

  private async frameOffset(): Promise<{ x: number; y: number }> {
    let x = 0;
    let y = 0;
    let frame = this.frame;
    while (frame.parentId && this.frameTree.has(frame.parentId)) {
      const parentId = frame.parentId;
      const owner = await this.cdp.send<{ backendNodeId: number }>("DOM.getFrameOwner", {
        frameId: frame.id,
      });
      const box = await this.cdp.send<{
        model: { content: number[] };
      }>("DOM.getBoxModel", { backendNodeId: owner.backendNodeId });
      x += box.model.content[0];
      y += box.model.content[1];
      const parent = this.frameTree.get(parentId);
      if (!parent) throw new Error(`Missing parent frame ${parentId}`);
      frame = parent;
    }
    return { x, y };
  }
}

function getCdpTargets(port: number): Promise<{ type?: string; webSocketDebuggerUrl?: string }[]> {
  return new Promise((resolve, reject) => {
    const request = httpGet(`http://127.0.0.1:${port}/json/list`, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk: string) => (body += chunk));
      response.on("end", () => {
        if (response.statusCode !== 200) {
          reject(
            new Error(`CDP /json/list returned HTTP ${response.statusCode}: ${body.slice(0, 200)}`),
          );
          return;
        }
        try {
          resolve(JSON.parse(body) as { type?: string; webSocketDebuggerUrl?: string }[]);
        } catch (error) {
          reject(new Error(`CDP /json/list returned invalid JSON: ${String(error)}`));
        }
      });
    });
    request.setTimeout(2000, () =>
      request.destroy(new Error(`CDP request timed out on port ${port}`)),
    );
    request.on("error", reject);
  });
}

function flattenFrames(node: FrameTreeNode): CdpFrame[] {
  return [node.frame, ...(node.childFrames ?? []).flatMap(flattenFrames)];
}
