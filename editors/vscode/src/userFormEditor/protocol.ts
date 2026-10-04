export interface DesignerError {
  code: string;
  message: string;
  line?: number;
  column?: number;
}

export interface SpecControl {
  id: string;
  parentId?: string;
  type: string;
  name: string;
  caption?: string;
  text?: string;
  value?: unknown;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  zIndex?: number;
  selectedIndex?: number;
  enabled?: boolean;
  visible?: boolean;
  list?: string[];
  tabs?: { name: string; caption?: string; visible?: boolean; enabled?: boolean }[];
  observed?: Partial<Omit<SpecControl, "observed">>;
}

export interface FormDimensions {
  caption?: string;
  width?: number;
  height?: number;
  clientWidth?: number;
  clientHeight?: number;
  insideWidth?: number;
  insideHeight?: number;
}

export interface PreviewDocument {
  form: FormDimensions & { name: string; build?: FormDimensions; observed?: FormDimensions };
  controls: SpecControl[];
}

export interface PreviewResult {
  version: number;
  document?: PreviewDocument;
  warnings?: { code: string; message: string }[];
  error?: DesignerError;
}

export interface DesignerControl extends Omit<SpecControl, "observed"> {
  left: number;
  top: number;
  width: number;
  height: number;
  approximate: boolean;
}

export interface DesignerDocument {
  name: string;
  caption: string;
  width: number;
  height: number;
  approximate: boolean;
  controls: DesignerControl[];
  warnings: string[];
}

export type HostMessage =
  | { type: "document"; version: number; document: DesignerDocument }
  | { type: "invalidDocument"; version: number; error: DesignerError }
  | { type: "themeChanged" };

export type WebviewMessage = { type: "ready" } | { type: "openText" };

export function isWebviewMessage(value: unknown): value is WebviewMessage {
  return (
    typeof value === "object" &&
    value !== null &&
    "type" in value &&
    (value.type === "ready" || value.type === "openText")
  );
}
