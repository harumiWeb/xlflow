import type { EditorDescriptor } from "./types";
import { fontStyle } from "./appearance";
import { useId } from "preact/hooks";
import { validNumericInput } from "./numeric";

const candidates = [
  "Arial",
  "Calibri",
  "MS Gothic",
  "MS Mincho",
  "Meiryo",
  "Segoe UI",
  "Tahoma",
  "Yu Gothic",
];
interface Props {
  descriptor: EditorDescriptor;
  value: Record<string, unknown>;
  disabled: boolean;
  unset: string;
  sample: string;
  onChange: (field: string, value: unknown, valid?: boolean) => void;
}
export function FontEditor({ descriptor, value, disabled, unset, sample, onChange }: Props) {
  const id = useId();
  return (
    <div class="font-editor">
      <div class="property-editor-fields">
        {(descriptor.children ?? []).map((child) => {
          const current = value[child.field];
          const fieldId = `${id}-${child.field}`;
          return (
            <label class="property-editor-field" key={child.field} for={fieldId}>
              <span>{child.field[0].toUpperCase() + child.field.slice(1)}</span>
              {child.valueType === "boolean" ? (
                <input
                  id={fieldId}
                  type="checkbox"
                  disabled={disabled}
                  checked={current === true}
                  ref={(element) => {
                    if (element) element.indeterminate = current === undefined || current === null;
                  }}
                  onChange={(event) => onChange(child.field, event.currentTarget.checked)}
                />
              ) : child.allowedValues ? (
                <select
                  id={fieldId}
                  disabled={disabled}
                  value={current === undefined || current === null ? "__unset__" : String(current)}
                  onChange={(event) =>
                    onChange(
                      child.field,
                      child.valueType === "integer" || child.valueType === "number"
                        ? Number(event.currentTarget.value)
                        : event.currentTarget.value,
                    )
                  }
                >
                  {(current === undefined || current === null) && (
                    <option value="__unset__" disabled>
                      {current === null ? "null" : unset}
                    </option>
                  )}
                  {current !== undefined &&
                    current !== null &&
                    !(child.allowedValues ?? []).includes(String(current)) && (
                      <option value={String(current)}>{String(current)}</option>
                    )}
                  {(
                    child.allowedValues?.map((choice) => ({
                      value: choice,
                      label:
                        child.suggestions?.find((hint) => String(hint.value) === choice)?.label ??
                        choice,
                    })) ?? []
                  ).map((choice) => (
                    <option value={String(choice.value)}>{choice.label}</option>
                  ))}
                </select>
              ) : (
                <>
                  <input
                    id={fieldId}
                    type="text"
                    disabled={disabled}
                    placeholder={
                      current === null ? "null" : current === undefined ? unset : undefined
                    }
                    list={
                      child.field === "name"
                        ? `${id}-fonts`
                        : child.suggestions
                          ? `${id}-${child.field}-suggestions`
                          : undefined
                    }
                    inputMode={
                      child.valueType === "number" || child.valueType === "integer"
                        ? "decimal"
                        : undefined
                    }
                    value={
                      typeof current === "string" || typeof current === "number"
                        ? String(current)
                        : ""
                    }
                    onInput={(event) => {
                      const text = event.currentTarget.value;
                      const numeric = child.valueType === "number" || child.valueType === "integer";
                      const parsed = numeric ? Number(text) : text;
                      const valid = !numeric || validNumericInput(text, child);
                      onChange(child.field, valid ? parsed : text, valid);
                    }}
                  />
                  {child.field === "name" && (
                    <datalist id={`${id}-fonts`}>
                      {[...new Set([typeof current === "string" ? current : "", ...candidates])]
                        .filter(Boolean)
                        .map((name) => (
                          <option value={name} />
                        ))}
                    </datalist>
                  )}
                  {child.field !== "name" && child.suggestions && (
                    <datalist id={`${id}-${child.field}-suggestions`}>
                      {child.suggestions.map((suggestion) => (
                        <option value={String(suggestion.value)}>{suggestion.label}</option>
                      ))}
                    </datalist>
                  )}
                </>
              )}
            </label>
          );
        })}
      </div>
      <div class="font-editor-sample" style={fontStyle(value)}>
        {sample}
      </div>
    </div>
  );
}
