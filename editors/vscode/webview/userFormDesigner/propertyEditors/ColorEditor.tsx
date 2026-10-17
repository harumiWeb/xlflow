import type { EditorDescriptor } from "./types";
import { colorStyle } from "./appearance";
import { useId, useState } from "preact/hooks";

// Presentation shortcuts, not a validation schema. Arbitrary RGB and OLE values
// remain editable below; system choices always come from the LSP descriptor.
const palette = [
  "#FFFFFF",
  "#FFC0C0",
  "#FFE0C0",
  "#FFFFC0",
  "#C0FFC0",
  "#C0FFFF",
  "#C0C0FF",
  "#FFC0FF",
  "#E0E0E0",
  "#FF8080",
  "#FFC080",
  "#FFFF80",
  "#80FF80",
  "#80FFFF",
  "#8080FF",
  "#FF80FF",
  "#C0C0C0",
  "#FF0000",
  "#FF8000",
  "#FFFF00",
  "#00FF00",
  "#00FFFF",
  "#0000FF",
  "#FF00FF",
  "#808080",
  "#C00000",
  "#C06000",
  "#C0C000",
  "#00C000",
  "#00C0C0",
  "#0000C0",
  "#C000C0",
  "#404040",
  "#800000",
  "#804000",
  "#808000",
  "#008000",
  "#008080",
  "#000080",
  "#800080",
  "#000000",
  "#400000",
  "#402000",
  "#404000",
  "#004000",
  "#004040",
  "#000040",
  "#400040",
];

interface Props {
  descriptor: EditorDescriptor;
  value: Record<string, unknown>;
  disabled: boolean;
  unset: string;
  paletteLabel: string;
  systemLabel: string;
  onChange: (value: Record<string, unknown>, valid?: boolean) => void;
}
export function ColorEditor({
  descriptor,
  value,
  disabled,
  unset,
  paletteLabel,
  systemLabel,
  onChange,
}: Props) {
  const id = useId();
  const [tab, setTab] = useState(value.kind === "system" ? "system" : "palette");
  const [focusIndex, setFocusIndex] = useState(() =>
    Math.max(0, palette.indexOf(String(value.value).toUpperCase())),
  );
  const kindDescriptor = descriptor.children?.find((child) => child.field === "kind");
  const kinds =
    kindDescriptor?.allowedValues ??
    kindDescriptor?.suggestions?.map((choice) => String(choice.value)) ??
    [];
  const valueDescriptor = descriptor.children?.find((child) => child.field === "value");
  const systemDescriptor = descriptor.children?.find((child) => child.field === "system");
  const systemChoices =
    systemDescriptor?.suggestions ??
    valueDescriptor?.suggestions ??
    (systemDescriptor?.allowedValues ?? valueDescriptor?.allowedValues ?? []).map((choice) => ({
      value: choice,
      label: choice,
    }));
  return (
    <div class="color-editor">
      <div class="color-picker-tabs" role="tablist">
        {[
          { key: "palette", label: paletteLabel, available: kinds.includes("rgb") },
          { key: "system", label: systemLabel, available: kinds.includes("system") },
        ]
          .filter((item) => item.available)
          .map((item) => (
            <button
              type="button"
              role="tab"
              id={`${id}-${item.key}-tab`}
              aria-selected={tab === item.key}
              aria-controls={`${id}-${item.key}-panel`}
              disabled={disabled}
              onClick={() => setTab(item.key)}
              onKeyDown={(event) => {
                if (
                  !["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key) ||
                  event.isComposing
                )
                  return;
                const tabs = Array.from(
                  event.currentTarget.parentElement!.querySelectorAll<HTMLButtonElement>("button"),
                );
                const index = tabs.indexOf(event.currentTarget);
                const next =
                  event.key === "Home"
                    ? 0
                    : event.key === "End"
                      ? tabs.length - 1
                      : (index + (event.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length;
                event.preventDefault();
                event.stopPropagation();
                tabs[next].focus();
                setTab(tabs[next].id.endsWith("system-tab") ? "system" : "palette");
              }}
            >
              {item.label}
            </button>
          ))}
      </div>
      <div
        class={tab === "palette" ? "color-palette" : "color-system-list"}
        role="tabpanel"
        id={`${id}-${tab}-panel`}
        aria-labelledby={`${id}-${tab}-tab`}
      >
        {tab === "palette" &&
          kinds.includes("rgb") &&
          palette.map((hex, index) => (
            <button
              type="button"
              class="color-palette-swatch"
              style={{ backgroundColor: hex }}
              title={hex}
              aria-label={hex}
              aria-pressed={
                value.kind === "rgb" &&
                typeof value.value === "string" &&
                value.value.toUpperCase() === hex
              }
              tabIndex={focusIndex === index ? 0 : -1}
              onFocus={() => setFocusIndex(index)}
              disabled={disabled}
              onClick={() => onChange({ kind: "rgb", value: hex })}
              onKeyDown={(event) => {
                const offset = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -8, ArrowDown: 8 }[
                  event.key
                ];
                const next =
                  event.key === "Home"
                    ? 0
                    : event.key === "End"
                      ? palette.length - 1
                      : offset === undefined
                        ? undefined
                        : Math.max(0, Math.min(palette.length - 1, index + offset));
                if (next === undefined || event.isComposing) return;
                event.preventDefault();
                event.stopPropagation();
                event.currentTarget.parentElement
                  ?.querySelectorAll<HTMLButtonElement>("button")
                  [next]?.focus();
              }}
            />
          ))}
        {tab === "system" &&
          kinds.includes("system") &&
          systemChoices.map((choice) => (
            <button
              type="button"
              class="color-system-choice"
              aria-pressed={value.kind === "system" && value.value === choice.value}
              disabled={disabled}
              onClick={() => onChange({ kind: "system", value: String(choice.value) })}
            >
              <span
                class="color-swatch"
                style={{
                  backgroundColor: colorStyle({ kind: "system", value: String(choice.value) }),
                }}
                aria-hidden="true"
              />
              {choice.label}
            </button>
          ))}
      </div>
      <label for={`${id}-kind`}>Kind</label>
      <select
        id={`${id}-kind`}
        value={typeof value.kind === "string" ? value.kind : "__unset__"}
        disabled={disabled}
        onChange={(event) => {
          const kind = event.currentTarget.value;
          setTab(kind === "system" ? "system" : "palette");
          onChange({ ...value, kind }, false);
        }}
      >
        {value.kind === undefined && (
          <option value="__unset__" disabled>
            {unset}
          </option>
        )}
        {kinds.map((kind) => (
          <option value={kind}>{kind === "rgb" ? "RGB" : kind === "ole" ? "OLE" : "System"}</option>
        ))}
      </select>
      <label for={`${id}-value`}>Value</label>
      {value.kind === "system" && systemChoices.length > 0 ? (
        <select
          id={`${id}-value`}
          value={typeof value.value === "string" ? value.value : "__unset__"}
          disabled={disabled}
          onChange={(event) => onChange({ ...value, value: event.currentTarget.value })}
        >
          {!systemChoices.some((choice) => choice.value === value.value) && (
            <option value={typeof value.value === "string" ? value.value : "__unset__"} disabled>
              {typeof value.value === "string" ? value.value : unset}
            </option>
          )}
          {systemChoices.map((choice) => (
            <option value={String(choice.value)}>{choice.label}</option>
          ))}
        </select>
      ) : (
        <input
          id={`${id}-value`}
          type="text"
          value={typeof value.value === "string" ? value.value : ""}
          disabled={disabled}
          placeholder={
            value.kind === "rgb" ? "#RRGGBB" : value.kind === "ole" ? "0x8000000F" : unset
          }
          onInput={(event) => {
            const next = { ...value, value: event.currentTarget.value };
            // Color syntax and applicability remain server decisions.
            onChange(next);
          }}
        />
      )}
      {value.kind === "rgb" && (
        <input
          aria-label="RGB"
          type="color"
          disabled={disabled}
          value={colorStyle(value) ?? "#000000"}
          onInput={(event) => onChange({ ...value, value: event.currentTarget.value })}
        />
      )}
      <div class="color-editor-sample" aria-label={String(value.value ?? unset)}>
        <span
          class="color-swatch"
          style={{ backgroundColor: colorStyle(value) }}
          aria-hidden="true"
        />
        <span class="color-identifier">
          {typeof value.value === "string" ? value.value : unset}
        </span>
      </div>
    </div>
  );
}
