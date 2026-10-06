import { useEffect, useRef, useState } from "preact/hooks";
import type {
  DesignerError,
  DesignerStrings,
  PropertyDescriptor,
  PropertyTarget,
  ScalarValue,
} from "../../src/userFormEditor/protocol";

// Presentation only. The server owns types, applicability and authoring values.
const categories = {
  identity: ["name", "tag"],
  appearance: ["caption", "text", "value", "controlTipText"],
  layout: ["left", "top", "width", "height"],
  behavior: ["enabled", "visible"],
  navigation: ["tabIndex", "selectedIndex", "accelerator"],
  advanced: [
    "build.caption",
    "build.width",
    "build.height",
    "build.clientWidth",
    "build.clientHeight",
  ],
} as const;
type Category = keyof typeof categories;
const categoryOf = (field: string): Category =>
  (Object.keys(categories) as (keyof typeof categories)[]).find((key) =>
    (categories[key] as readonly string[]).includes(field),
  ) ?? "advanced";
const labels: Record<string, keyof DesignerStrings> = {
  name: "propertyName",
  caption: "propertyCaption",
  text: "propertyText",
  value: "propertyValue",
  left: "propertyLeft",
  top: "propertyTop",
  width: "propertyWidth",
  height: "propertyHeight",
  tabIndex: "propertyTabIndex",
  selectedIndex: "propertySelectedIndex",
  enabled: "propertyEnabled",
  visible: "propertyVisible",
  tag: "propertyTag",
  controlTipText: "propertyControlTipText",
  accelerator: "propertyAccelerator",
  "build.caption": "propertyBuildCaption",
  "build.width": "propertyBuildWidth",
  "build.height": "propertyBuildHeight",
  "build.clientWidth": "propertyClientWidth",
  "build.clientHeight": "propertyClientHeight",
};

export interface PropertyReply {
  requestId: number;
  controlId?: string;
  field?: string;
  error?: DesignerError;
}
interface Props {
  drafts: PropertyDraftStore;
  target: PropertyTarget;
  controlId?: string;
  title: string;
  version: number;
  editable: boolean;
  busy: boolean;
  strings: DesignerStrings;
  reply?: PropertyReply;
  onCommit: (field: string, value: ScalarValue) => number | undefined;
}
type ValueKind = "string" | "number" | "boolean" | "null";
interface Draft {
  text: string;
  kind: ValueKind;
  dirty: boolean;
  version: number;
  submitted?: number;
  attempted?: boolean;
  error?: string;
}
// Owned by App so switching selections cannot discard a failed or pending edit.
export interface PropertyDraftStore {
  values: Map<string, Draft>;
  replies: Map<number, PropertyReply>;
}
function PropertyRow({ descriptor, ...props }: Props & { descriptor: PropertyDescriptor }) {
  const { target, version, editable, busy, strings, reply, controlId, onCommit, drafts } = props;
  const draftKey = JSON.stringify([controlId ?? null, descriptor.field]);
  const source = target.values[descriptor.field];
  const fromSource = (): Draft => {
    const value = source?.present ? source.value : undefined;
    return {
      text: value == null ? "" : String(value),
      kind:
        value === null
          ? "null"
          : typeof value === "number"
            ? "number"
            : typeof value === "boolean"
              ? "boolean"
              : descriptor.valueType === "number" || descriptor.valueType === "integer"
                ? "number"
                : descriptor.valueType === "boolean"
                  ? "boolean"
                  : "string",
      dirty: false,
      version,
    };
  };
  const [draft, setDraft] = useState<Draft>(() => drafts.values.get(draftKey) ?? fromSource());
  const current = useRef(draft);
  const update = (next: Draft) => {
    current.current = next;
    if (next.dirty) drafts.values.set(draftKey, next);
    else drafts.values.delete(draftKey);
    setDraft(next);
  };
  const stale = draft.dirty && draft.version !== version && draft.submitted === undefined;
  useEffect(() => {
    const d = current.current;
    // A preview can arrive before a failed edit reply. Only the matching reply
    // may discard a submitted draft; a document update alone is not success.
    if (!d.dirty) update(fromSource());
  }, [target, version]);
  useEffect(() => {
    const response =
      current.current.submitted === undefined
        ? undefined
        : drafts.replies.get(current.current.submitted);
    if (!response || response.controlId !== controlId || response.field !== descriptor.field)
      return;
    drafts.replies.delete(response.requestId);
    if (response.error) {
      // This reply belongs to this row's single submitted operation. The server
      // may report canonical paths such as controls[0].caption or form.caption.
      const diagnostics = response.error.diagnostics?.filter((d) => d.operationIndex === 0);
      const message = diagnostics?.length
        ? diagnostics
            .map((d) => `${d.code}: ${d.message}${d.suggestion ? ` ${d.suggestion}` : ""}`)
            .join("\n")
        : `${response.error.code}: ${response.error.message}`;
      update({ ...current.current, submitted: undefined, error: message });
    } else update(fromSource());
  }, [reply]);
  const commit = () => {
    const d = current.current;
    if (
      !editable ||
      busy ||
      !d.dirty ||
      d.attempted ||
      d.submitted !== undefined ||
      d.version !== version
    )
      return;
    let value: ScalarValue = d.text;
    if (d.kind === "null") {
      if (!descriptor.nullable || descriptor.required) return;
      value = null;
    } else if (d.kind === "boolean") value = d.text === "true";
    else if (d.kind === "number") {
      value = Number(d.text);
      if (!d.text.trim() || !Number.isFinite(value)) {
        update({ ...d, error: strings.invalidNumber });
        return;
      }
      if (descriptor.valueType === "integer" && !Number.isSafeInteger(value)) {
        update({ ...d, error: strings.invalidInteger });
        return;
      }
    }
    if (
      descriptor.allowedValues &&
      !descriptor.allowedValues.some((allowed) => {
        if (d.kind === "number")
          return (
            allowed.trim() !== "" && Number.isFinite(Number(allowed)) && Number(allowed) === value
          );
        if (d.kind === "boolean")
          return (allowed === "true" || allowed === "false") && (allowed === "true") === value;
        return allowed === value;
      })
    ) {
      update({ ...d, error: strings.invalidValue });
      return;
    }
    if (source?.present && source.value === value) {
      update(fromSource());
      return;
    }
    const requestId = onCommit(descriptor.field, value);
    if (requestId !== undefined)
      update({ ...d, submitted: requestId, attempted: true, error: undefined });
  };
  const change = (text: string, kind = current.current.kind) =>
    update({ text, kind, version, dirty: true });
  const id = `property-${controlId ?? "form"}-${descriptor.field}`;
  const error = stale
    ? [strings.staleProperty, draft.error].filter(Boolean).join("\n")
    : draft.error;
  const disabled = !editable || busy;
  const kinds: ValueKind[] =
    descriptor.valueType === "any"
      ? ["string", "number", "boolean"]
      : [descriptor.valueType === "integer" ? "number" : descriptor.valueType];
  if (descriptor.nullable && !descriptor.required) kinds.push("null");
  return (
    <div class="property-row" data-property-field={descriptor.field}>
      <label for={id}>{strings[labels[descriptor.field]] ?? descriptor.field}</label>
      {!source?.present && !draft.dirty && <span class="property-state">{strings.unset}</span>}
      {(descriptor.valueType === "any" || kinds.includes("null")) && (
        <select
          aria-label={`${strings[labels[descriptor.field]] ?? descriptor.field}: ${strings.valueType}`}
          value={draft.kind}
          disabled={disabled}
          onKeyDown={(event) => {
            if (event.key === "Escape") {
              event.preventDefault();
              update(fromSource());
            }
          }}
          onChange={(event) => {
            const kind = event.currentTarget.value as ValueKind;
            change(
              kind === "boolean" ? "false" : kind === "null" ? "" : current.current.text,
              kind,
            );
            if (kind === "boolean" || kind === "null") commit();
          }}
        >
          {kinds.map((kind) => (
            <option value={kind}>
              {kind === "string"
                ? strings.typeString
                : kind === "number"
                  ? strings.typeNumber
                  : kind === "boolean"
                    ? strings.typeBoolean
                    : strings.nullValue}
            </option>
          ))}
        </select>
      )}
      {draft.kind === "null" ? (
        <span id={id} class="property-state">
          {strings.nullValue}
        </span>
      ) : descriptor.allowedValues || draft.kind === "boolean" ? (
        <select
          id={id}
          value={!source?.present && !draft.dirty ? "__unset__" : draft.text}
          disabled={disabled}
          aria-invalid={!!error}
          aria-describedby={error ? `${id}-error` : undefined}
          onChange={(event) => {
            change(event.currentTarget.value);
            commit();
          }}
          onKeyDown={(event) => {
            if (event.key === "Escape") {
              event.preventDefault();
              update(fromSource());
            }
          }}
        >
          {!source?.present && !draft.dirty && (
            <option value="__unset__" disabled>
              {strings.unset}
            </option>
          )}
          {(descriptor.allowedValues ?? ["true", "false"]).map((value) => (
            <option value={value}>{value}</option>
          ))}
        </select>
      ) : (
        <input
          id={id}
          type="text"
          inputMode={draft.kind === "number" ? "decimal" : undefined}
          value={draft.text}
          disabled={disabled}
          aria-invalid={!!error}
          aria-describedby={error ? `${id}-error` : undefined}
          onInput={(event) => change(event.currentTarget.value)}
          onBlur={commit}
          onKeyDown={(event) => {
            if (event.isComposing) return;
            if (event.key === "Enter") {
              event.preventDefault();
              commit();
            }
            if (event.key === "Escape") {
              event.preventDefault();
              update(fromSource());
            }
          }}
        />
      )}
      {error && (
        <p id={`${id}-error`} role="alert" class="property-error">
          {error}
        </p>
      )}
    </div>
  );
}

export function PropertyGrid(props: Props) {
  const groups = Object.keys(categories) as Category[];
  return (
    <aside class="property-grid" aria-label={props.strings.properties}>
      <h2>
        {props.strings.properties}: {props.title}
      </h2>
      {!props.editable && <p class="notice">{props.strings.propertyReadOnly}</p>}
      {groups.map((category) => {
        const order: readonly string[] = categories[category];
        const descriptors = props.target.descriptors
          .filter((d) => categoryOf(d.field) === category)
          .sort((a, b) => {
            const ai = order.indexOf(a.field),
              bi = order.indexOf(b.field);
            return (ai < 0 ? order.length : ai) - (bi < 0 ? order.length : bi);
          });
        return (
          descriptors.length > 0 && (
            <fieldset>
              <legend>{props.strings[category]}</legend>
              {descriptors.map((descriptor) => (
                <PropertyRow
                  key={`${props.controlId ?? "form"}:${descriptor.field}`}
                  {...props}
                  descriptor={descriptor}
                />
              ))}
            </fieldset>
          )
        );
      })}
    </aside>
  );
}
