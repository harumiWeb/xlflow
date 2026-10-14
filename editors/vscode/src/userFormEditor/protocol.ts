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
  caption?: string | null;
  width?: number | null;
  height?: number | null;
}

export interface ObservedFormDimensions {
  caption?: string | null;
  width?: number | null;
  height?: number | null;
  clientWidth?: number;
  clientHeight?: number;
  insideWidth?: number;
  insideHeight?: number;
}

export interface PreviewDocument {
  form: FormDimensions & { name: string; observed?: ObservedFormDimensions };
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
  productivityEditable?: boolean;
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
  save: "Save",
  unsaved: "Unsaved Designer changes",
  sourceConflict: "YAML/JSON and Designer changes conflict. Resolve the conflict before saving.",
  useSource: "Use YAML/JSON changes",
  useDesigner: "Use Designer draft",
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
  resizeWidth: "Resize form width",
  resizeHeight: "Resize form height",
  resizeBoth: "Resize form width and height",
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
  productivity: "Productivity",
  align: "Align",
  alignLeft: "Align Left",
  alignCenter: "Align Center",
  alignRight: "Align Right",
  alignTop: "Align Top",
  alignMiddle: "Align Middle",
  alignBottom: "Align Bottom",
  distribute: "Distribute",
  distributeHorizontal: "Distribute Horizontally",
  distributeVertical: "Distribute Vertically",
  sameSize: "Same size",
  sameWidth: "Make Same Width",
  sameHeight: "Make Same Height",
  sameDimensions: "Make Same Size",
  cut: "Cut",
  copy: "Copy",
  paste: "Paste",
  duplicate: "Duplicate",
  bringToFront: "Bring to Front",
  bringForward: "Bring Forward",
  sendBackward: "Send Backward",
  sendToBack: "Send to Back",
  tabOrder: "Tab order",
  tabOrderParent: "Parent",
  applyTabOrder: "Apply tab order",
  noEditableTabOrder: "No controls with an editable TabIndex exist under this parent.",
  selectionCount: "Selected controls",
  propertiesAction: "Properties",
  selectControls: "Select controls",
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
      productivityEditable?: boolean;
      propertyGrid?: PropertyGridData;
      propertyError?: DesignerError;
    }
  | { type: "invalidDocument"; version: number; error: DesignerError }
  | { type: "editResult"; requestId: number; error?: DesignerError; selectionIds?: string[] }
  | { type: "clipboardState"; available: boolean }
  | { type: "draftState"; dirty: boolean; conflicted: boolean; saving: boolean; error?: string }
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
export interface ClipboardSnapshot {
  format: "yaml" | "json";
  text: string;
  controlIds: string[];
}
export interface CopyResult {
  version: number;
  snapshot?: ClipboardSnapshot;
  document?: PreviewDocument;
  hasPicturePaths?: boolean;
  error?: DesignerError;
}
export type ProductivityOperation =
  | { type: "moveControls"; moves: { controlId: string; left: number; top: number }[] }
  | {
      type: "alignControls";
      controlIds: string[];
      anchorId: string;
      alignment: "left" | "center" | "right" | "top" | "middle" | "bottom";
    }
  | { type: "distributeControls"; controlIds: string[]; axis: "horizontal" | "vertical" }
  | {
      type: "sameSizeControls";
      controlIds: string[];
      anchorId: string;
      dimension: "width" | "height" | "size";
    }
  | { type: "removeControls"; controlIds: string[]; cascade: boolean }
  | {
      type: "reorderControls";
      controlIds: string[];
      command: "front" | "forward" | "backward" | "back";
    }
  | { type: "setTabOrder"; parentId: string; controlIds: string[] }
  | {
      type: "pasteControls";
      snapshot: ClipboardSnapshot;
      allocations: { sourceId: string; id: string; name: string }[];
      parentId: string;
      offset: { left: number; top: number };
    };
export type SemanticOperation =
  | GeometryOperation
  | PropertyOperation
  | StructuralOperation
  | ParentOperation
  | PageOperation
  | TabOperation
  | ProductivityOperation;

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

const isStringArray = (value: unknown, minimumLength = 0): value is string[] =>
  Array.isArray(value) &&
  value.length >= minimumLength &&
  value.every((item) => typeof item === "string" && item.length > 0);

export function isProductivityOperation(value: unknown): value is ProductivityOperation {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  const op = value as Record<string, unknown>;
  const exactKeys = (keys: string[]) =>
    Object.keys(op).length === keys.length && Object.keys(op).every((key) => keys.includes(key));
  const finite = (n: unknown) => typeof n === "number" && Number.isFinite(n);
  if (op.type === "moveControls") {
    return (
      exactKeys(["type", "moves"]) &&
      Array.isArray(op.moves) &&
      op.moves.length >= 2 &&
      op.moves.every((move) => {
        if (typeof move !== "object" || move === null || Array.isArray(move)) return false;
        const item = move as Record<string, unknown>;
        return (
          Object.keys(item).length === 3 &&
          Object.keys(item).every((key) => ["controlId", "left", "top"].includes(key)) &&
          typeof item.controlId === "string" &&
          item.controlId.length > 0 &&
          finite(item.left) &&
          finite(item.top)
        );
      }) &&
      new Set((op.moves as { controlId: string }[]).map((move) => move.controlId)).size ===
        (op.moves as unknown[]).length
    );
  }
  if (op.type === "alignControls")
    return (
      exactKeys(["type", "controlIds", "anchorId", "alignment"]) &&
      isStringArray(op.controlIds, 2) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.anchorId === "string" &&
      op.controlIds.includes(op.anchorId) &&
      ["left", "center", "right", "top", "middle", "bottom"].includes(String(op.alignment))
    );
  if (op.type === "distributeControls")
    return (
      exactKeys(["type", "controlIds", "axis"]) &&
      isStringArray(op.controlIds, 3) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      ["horizontal", "vertical"].includes(String(op.axis))
    );
  if (op.type === "sameSizeControls")
    return (
      exactKeys(["type", "controlIds", "anchorId", "dimension"]) &&
      isStringArray(op.controlIds, 2) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.anchorId === "string" &&
      op.controlIds.includes(op.anchorId) &&
      ["width", "height", "size"].includes(String(op.dimension))
    );
  if (op.type === "removeControls")
    return (
      exactKeys(["type", "controlIds", "cascade"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.cascade === "boolean"
    );
  if (op.type === "reorderControls")
    return (
      exactKeys(["type", "controlIds", "command"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      ["front", "forward", "backward", "back"].includes(String(op.command))
    );
  if (op.type === "setTabOrder")
    return (
      exactKeys(["type", "parentId", "controlIds"]) &&
      typeof op.parentId === "string" &&
      isStringArray(op.controlIds) &&
      new Set(op.controlIds).size === op.controlIds.length
    );
  if (op.type === "pasteControls") {
    const snapshot = op.snapshot as Record<string, unknown> | null;
    const offset = op.offset as Record<string, unknown> | null;
    return (
      exactKeys(["type", "snapshot", "allocations", "parentId", "offset"]) &&
      typeof snapshot === "object" &&
      snapshot !== null &&
      Object.keys(snapshot).length === 3 &&
      Object.keys(snapshot).every((key) => ["format", "text", "controlIds"].includes(key)) &&
      ["yaml", "json"].includes(String(snapshot.format)) &&
      typeof snapshot.text === "string" &&
      isStringArray(snapshot.controlIds, 1) &&
      new Set(snapshot.controlIds).size === snapshot.controlIds.length &&
      Array.isArray(op.allocations) &&
      op.allocations.length >= 1 &&
      op.allocations.every((allocation) => {
        if (typeof allocation !== "object" || allocation === null || Array.isArray(allocation))
          return false;
        const item = allocation as Record<string, unknown>;
        return (
          Object.keys(item).length === 3 &&
          Object.keys(item).every((key) => ["sourceId", "id", "name"].includes(key)) &&
          typeof item.sourceId === "string" &&
          item.sourceId.length > 0 &&
          typeof item.id === "string" &&
          item.id.length > 0 &&
          typeof item.name === "string" &&
          /^[A-Za-z][A-Za-z0-9_]*$/.test(item.name)
        );
      }) &&
      typeof op.parentId === "string" &&
      typeof offset === "object" &&
      offset !== null &&
      Object.keys(offset).length === 2 &&
      Object.keys(offset).every((key) => ["left", "top"].includes(key)) &&
      finite(offset.left) &&
      finite(offset.top)
    );
  }
  return false;
}

function isClipboardMessage(value: Record<string, unknown>): boolean {
  return (
    value.type === "clipboard" &&
    Object.keys(value).length === 6 &&
    Object.keys(value).every((key) =>
      ["type", "requestId", "version", "action", "controlIds", "parentId"].includes(key),
    ) &&
    Number.isSafeInteger(value.requestId) &&
    Number(value.requestId) >= 0 &&
    Number.isSafeInteger(value.version) &&
    Number(value.version) >= 0 &&
    ["copy", "cut", "paste", "duplicate"].includes(String(value.action)) &&
    isStringArray(value.controlIds) &&
    new Set(value.controlIds).size === value.controlIds.length &&
    typeof value.parentId === "string"
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
  selectionIds?: string[];
}
export type WebviewMessage =
  | { type: "ready" }
  | { type: "openText" }
  | { type: "save" }
  | { type: "resolveConflict"; choice: "source" | "designer" }
  | {
      type: "clipboard";
      requestId: number;
      version: number;
      action: "copy" | "cut" | "paste" | "duplicate";
      controlIds: string[];
      parentId: string;
    }
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
  if (message.type === "ready" || message.type === "openText" || message.type === "save")
    return Object.keys(message).length === 1;
  if (message.type === "resolveConflict")
    return (
      Object.keys(message).length === 2 && ["source", "designer"].includes(String(message.choice))
    );
  if (message.type === "clipboard") return isClipboardMessage(message);
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
        isTabOperation(operations[0]) ||
        (isProductivityOperation(operations[0]) && operations[0].type !== "pasteControls"))) ||
      (operations.length === 2 &&
        operations.every(
          (op) =>
            isPropertyOperation(op) &&
            op.type === "setFormProperty" &&
            ["width", "height"].includes(op.field) &&
            typeof op.value === "number",
        ) &&
        new Set(operations.map((op) => op.field)).size === 2) ||
      (operations.every(isGeometryOperation) &&
        operations.every((op) => op.controlId === operations[0].controlId) &&
        new Set(operations.map((op) => op.type)).size === operations.length))
  );
}
