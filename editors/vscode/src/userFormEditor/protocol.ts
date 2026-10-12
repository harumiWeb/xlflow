import { controlMetadata } from "./toolbox";
import type { NewControl } from "./toolbox";

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
  tabs?: Record<string, PropertyTarget[]>;
}

export interface SpecTab {
  name: string;
  caption?: string;
  visible?: boolean;
  enabled?: boolean;
  tag?: string;
  controlTipText?: string;
  accelerator?: string;
}

export interface SpecControl {
  id: string;
  parentId?: string;
  type: string;
  progId?: string;
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
  tabs?: SpecTab[];
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
  structuralEditable?: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  tabEditable?: boolean;
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
  invalidNumberPrecision: "This number would lose precision. Use the source editor.",
  invalidInteger: "Enter a safe integer.",
  invalidValue: "Enter a valid value.",
  staleProperty: "The document changed. Re-enter this value on the updated form.",
  propertyReadOnly:
    "Property editing requires an xlflow version supporting UserForm property edits.",
  toolbox: "Toolbox",
  pointer: "Pointer",
  deleteControl: "Delete control",
  structuralReadOnly:
    "Adding and deleting controls requires an xlflow version supporting structural edits.",
  cannotInsert: "This control's default size is larger than the form.",
  placement:
    "Click the form to place the control. Escape cancels. Controls are added at the form root.",
  hierarchy: "Selection path",
  invalidDrop: "This control cannot be placed in this container.",
  dropTarget: "Drop target",
  parentReadOnly: "Changing parents requires an xlflow version supporting parent edits.",
  parentPlacement: "Click the form or a Frame to place the control. Escape cancels.",
  cannotInsertContainer: "This control's default size is larger than the container.",
  pages: "Pages",
  tabs: "Tabs",
  addPage: "Add Page",
  addTab: "Add tab",
  deleteItem: "Delete",
  moveUp: "Move up",
  moveDown: "Move down",
  moveToPage: "Move to Page",
  choosePage: "Choose destination Page",
  collectionReadOnly: "Collection editing requires an xlflow version supporting Page or tab edits.",
  pagePlacement: "Click the form, a Frame or the active Page to place the control. Escape cancels.",
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
      structuralEditable?: boolean;
      parentEditable?: boolean;
      pageEditable?: boolean;
      tabEditable?: boolean;
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
export type StructuralOperation =
  | { type: "addControl"; control: NewControl }
  | { type: "removeControl"; controlId: string; cascade: boolean };
export type ParentOperation = {
  type: "setParent";
  controlId: string;
  parentId: string;
  left: number;
  top: number;
};
export type PageOperation =
  | { type: "addPage"; control: Pick<SpecControl, "id" | "parentId" | "name" | "type" | "caption"> }
  | { type: "removePage"; controlId: string; cascade: boolean }
  | { type: "reorderPage"; controlId: string; index: number };
export type TabOperation =
  | { type: "addTab"; controlId: string; tab: SpecTab }
  | { type: "removeTab"; controlId: string; index: number }
  | { type: "reorderTab"; controlId: string; index: number; toIndex: number }
  | { type: "setTabProperty"; controlId: string; index: number; field: string; value: ScalarValue };
export type SemanticOperation =
  | GeometryOperation
  | PropertyOperation
  | StructuralOperation
  | ParentOperation
  | PageOperation
  | TabOperation;

export function isPageOperation(value: unknown): value is PageOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (op.type === "addPage") {
    if (Object.keys(op).length !== 2 || typeof op.control !== "object" || op.control === null)
      return false;
    const c = op.control as Record<string, unknown>;
    return (
      Object.keys(c).every((k) => ["id", "parentId", "name", "type", "caption"].includes(k)) &&
      typeof c.id === "string" &&
      !!c.id &&
      typeof c.parentId === "string" &&
      !!c.parentId &&
      typeof c.name === "string" &&
      /^[A-Za-z][A-Za-z0-9_]*$/.test(c.name) &&
      c.type === "Page" &&
      (!Object.hasOwn(c, "caption") || typeof c.caption === "string")
    );
  }
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  if (op.type === "removePage")
    return Object.keys(op).length === 3 && typeof op.cascade === "boolean";
  return (
    op.type === "reorderPage" &&
    Object.keys(op).length === 3 &&
    Number.isSafeInteger(op.index) &&
    Number(op.index) >= 0
  );
}

export function isTabOperation(value: unknown): value is TabOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  if (op.type === "addTab") {
    if (
      Object.keys(op).length !== 3 ||
      typeof op.tab !== "object" ||
      op.tab === null ||
      Array.isArray(op.tab)
    )
      return false;
    const tab = op.tab as Record<string, unknown>;
    return (
      typeof tab.name === "string" &&
      !!tab.name &&
      Object.entries(tab).every(([key, v]) =>
        ["enabled", "visible"].includes(key)
          ? typeof v === "boolean"
          : ["name", "caption", "tag", "controlTipText", "accelerator"].includes(key) &&
            typeof v === "string",
      )
    );
  }
  if (!Number.isSafeInteger(op.index) || Number(op.index) < 0) return false;
  if (op.type === "removeTab") return Object.keys(op).length === 3;
  if (op.type === "reorderTab")
    return (
      Object.keys(op).length === 4 && Number.isSafeInteger(op.toIndex) && Number(op.toIndex) >= 0
    );
  return (
    op.type === "setTabProperty" &&
    Object.keys(op).length === 5 &&
    typeof op.field === "string" &&
    !!op.field &&
    (op.value === null ||
      typeof op.value === "string" ||
      typeof op.value === "boolean" ||
      (typeof op.value === "number" && Number.isFinite(op.value)))
  );
}

export function isStructuralOperation(value: unknown): value is StructuralOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (op.type === "removeControl")
    return (
      Object.keys(op).length === 3 &&
      Object.keys(op).every((key) => ["type", "controlId", "cascade"].includes(key)) &&
      typeof op.controlId === "string" &&
      op.controlId.length > 0 &&
      typeof op.cascade === "boolean"
    );
  if (
    op.type !== "addControl" ||
    Object.keys(op).length !== 2 ||
    typeof op.control !== "object" ||
    op.control === null ||
    Array.isArray(op.control)
  )
    return false;
  const control = op.control as Record<string, unknown>;
  return (
    Object.keys(control).every((key) =>
      [
        "id",
        "parentId",
        "name",
        "type",
        "left",
        "top",
        "width",
        "height",
        "caption",
        "selectedIndex",
        "tabs",
      ].includes(key),
    ) &&
    typeof control.id === "string" &&
    control.id.length > 0 &&
    typeof control.name === "string" &&
    /^[A-Za-z][A-Za-z0-9_]*$/.test(control.name) &&
    (!Object.hasOwn(control, "parentId") ||
      (typeof control.parentId === "string" && control.parentId.length > 0)) &&
    controlMetadata.some((m) => m.type === control.type) &&
    ["left", "top", "width", "height"].every(
      (key) => typeof control[key] === "number" && Number.isFinite(control[key]),
    ) &&
    Number(control.width) > 0 &&
    Number(control.height) > 0 &&
    (!Object.hasOwn(control, "caption") || typeof control.caption === "string") &&
    (!Object.hasOwn(control, "selectedIndex") || control.selectedIndex === -1) &&
    (!Object.hasOwn(control, "tabs") || (Array.isArray(control.tabs) && control.tabs.length === 0))
  );
}

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

export function isParentOperation(value: unknown): value is ParentOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  return (
    Object.keys(op).length === 5 &&
    Object.keys(op).every((key) =>
      ["type", "controlId", "parentId", "left", "top"].includes(key),
    ) &&
    op.type === "setParent" &&
    typeof op.controlId === "string" &&
    op.controlId.length > 0 &&
    typeof op.parentId === "string" &&
    typeof op.left === "number" &&
    Number.isFinite(op.left) &&
    typeof op.top === "number" &&
    Number.isFinite(op.top)
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
    ((operations.length === 1 &&
      (isPropertyOperation(operations[0]) ||
        isStructuralOperation(operations[0]) ||
        isParentOperation(operations[0]) ||
        isPageOperation(operations[0]) ||
        isTabOperation(operations[0]))) ||
      (operations.every(isGeometryOperation) &&
        operations.every((op) => op.controlId === operations[0].controlId) &&
        new Set(operations.map((op) => op.type)).size === operations.length))
  );
}
