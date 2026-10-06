export interface DesignerError {
  code: string;
  message: string;
  line?: number;
  column?: number;
  diagnostics?: {
    operationIndex: number;
    operation?: string;
    controlId?: string;
    field?: string;
    code: string;
    message: string;
    suggestion?: string;
  }[];
}

export type ScalarValue = string | number | boolean | null;
export interface PropertyDescriptor {
  field: string;
  valueType: "string" | "number" | "integer" | "boolean" | "any";
  required: boolean;
  nullable: boolean;
  allowedValues?: string[];
}
export interface PropertyState {
  present: boolean;
  value: ScalarValue;
}
export interface PropertyTarget {
  descriptors: PropertyDescriptor[];
  values: Record<string, PropertyState>;
}
export interface PropertyGridData {
  form: PropertyTarget;
  controls: Record<string, PropertyTarget>;
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
  editable?: boolean;
  propertyEditable?: boolean;
  propertyGrid?: PropertyGridData;
  propertyError?: DesignerError;
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

export const designerStrings = {
  header: "xlflow UserForm Designer",
  openText: "Open text editor",
  lastValid: "Showing the last valid document.",
  loading: "Loading FormSpec…",
  approximate:
    "Approximate preview. Missing dimensions use display defaults; Page bounds derive from MultiPage.",
  approximateBounds: "approximate bounds",
  grid: "Show grid",
  snap: "Snap to grid",
  zoom: "Zoom",
  readOnly: "Editing requires an xlflow version supporting UserForm edits.",
  cannotMove: "This control is larger than its parent. Resize it to fit before moving.",
  properties: "Properties",
  identity: "Identity",
  appearance: "Appearance",
  layout: "Layout",
  behavior: "Behavior",
  navigation: "Navigation",
  advanced: "Advanced",
  propertyName: "Name",
  propertyCaption: "Caption",
  propertyText: "Text",
  propertyValue: "Value",
  propertyLeft: "Left",
  propertyTop: "Top",
  propertyWidth: "Width",
  propertyHeight: "Height",
  propertyTabIndex: "TabIndex",
  propertySelectedIndex: "SelectedIndex",
  propertyEnabled: "Enabled",
  propertyVisible: "Visible",
  propertyTag: "Tag",
  propertyControlTipText: "ControlTipText",
  propertyAccelerator: "Accelerator",
  propertyBuildCaption: "Build caption",
  propertyBuildWidth: "Build width",
  propertyBuildHeight: "Build height",
  propertyClientWidth: "Client width",
  propertyClientHeight: "Client height",
  unset: "Not set",
  nullValue: "null",
  valueType: "Value type",
  typeString: "String",
  typeNumber: "Number",
  typeBoolean: "Boolean",
  invalidNumber: "Enter a finite number.",
  invalidInteger: "Enter a safe integer.",
  invalidValue: "Enter a valid value.",
  staleProperty: "The document changed. Re-enter this value on the updated form.",
  propertyReadOnly:
    "Property editing requires an xlflow version supporting UserForm property edits.",
};
export type DesignerStrings = typeof designerStrings;

export type HostMessage =
  | { type: "localization"; strings: DesignerStrings }
  | {
      type: "document";
      version: number;
      document: DesignerDocument;
      editable?: boolean;
      propertyEditable?: boolean;
      propertyGrid?: PropertyGridData;
      propertyError?: DesignerError;
    }
  | { type: "invalidDocument"; version: number; error: DesignerError }
  | { type: "editResult"; requestId: number; error?: DesignerError }
  | { type: "editingUnavailable" }
  | { type: "themeChanged" };

export type GeometryOperation =
  | { type: "moveControl"; controlId: string; left: number; top: number }
  | { type: "resizeControl"; controlId: string; width: number; height: number };

export type PropertyOperation =
  | { type: "setFormProperty"; field: string; value: ScalarValue }
  | { type: "setControlProperty"; controlId: string; field: string; value: ScalarValue };
export type SemanticOperation = GeometryOperation | PropertyOperation;

export function isPropertyOperation(value: unknown): value is PropertyOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  const control = op.type === "setControlProperty";
  if (!control && op.type !== "setFormProperty") return false;
  const keys = control ? ["type", "controlId", "field", "value"] : ["type", "field", "value"];
  return (
    Object.keys(op).length === keys.length &&
    Object.keys(op).every((key) => keys.includes(key)) &&
    (!control || (typeof op.controlId === "string" && op.controlId.length > 0)) &&
    typeof op.field === "string" &&
    op.field.length > 0 &&
    (op.value === null ||
      typeof op.value === "string" ||
      typeof op.value === "boolean" ||
      (typeof op.value === "number" && Number.isFinite(op.value)))
  );
}

export interface SourceTextEdit {
  range: { start: { line: number; character: number }; end: { line: number; character: number } };
  newText: string;
}
export interface EditResult {
  version: number;
  edits?: SourceTextEdit[];
  error?: DesignerError;
}
export type WebviewMessage =
  | { type: "ready" }
  | { type: "openText" }
  | { type: "edit"; requestId: number; version: number; operations: SemanticOperation[] };

export function isGeometryOperation(value: unknown): value is GeometryOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  const fields =
    op.type === "moveControl"
      ? ["left", "top"]
      : op.type === "resizeControl"
        ? ["width", "height"]
        : [];
  return (
    fields.length === 2 &&
    Object.keys(op).every((key) => ["type", "controlId", ...fields].includes(key)) &&
    fields.every((key) => typeof op[key] === "number" && Number.isFinite(op[key]))
  );
}

export function isWebviewMessage(value: unknown): value is WebviewMessage {
  if (typeof value !== "object" || value === null) return false;
  const message = value as Record<string, unknown>;
  if (message.type === "ready" || message.type === "openText")
    return Object.keys(message).length === 1;
  const operations = message.operations;
  return (
    message.type === "edit" &&
    Object.keys(message).every((key) =>
      ["type", "requestId", "version", "operations"].includes(key),
    ) &&
    Number.isSafeInteger(message.requestId) &&
    Number(message.requestId) >= 0 &&
    Number.isSafeInteger(message.version) &&
    Number(message.version) >= 0 &&
    Array.isArray(operations) &&
    operations.length >= 1 &&
    operations.length <= 2 &&
    ((operations.length === 1 && isPropertyOperation(operations[0])) ||
      (operations.every(isGeometryOperation) &&
        operations.every((op) => op.controlId === operations[0].controlId) &&
        new Set(operations.map((op) => op.type)).size === operations.length))
  );
}
