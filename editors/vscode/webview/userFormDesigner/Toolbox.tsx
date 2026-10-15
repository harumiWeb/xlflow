import { Icon } from "./Icon";
import { controlMetadata } from "../../src/userFormEditor/toolbox";
import type { ToolboxType } from "../../src/userFormEditor/toolbox";
import type { DesignerStrings } from "../../src/userFormEditor/protocol";

export function Toolbox({
  selected,
  disabled,
  parentEditable = false,
  pageEditable = false,
  tabEditable = false,
  strings,
  onSelect,
}: {
  selected?: ToolboxType;
  disabled: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  tabEditable?: boolean;
  strings: DesignerStrings;
  onSelect: (type?: ToolboxType) => void;
}) {
  const icons: Record<ToolboxType, string> = {
    Label: "symbol-string",
    TextBox: "symbol-field",
    ComboBox: "list-selection",
    ListBox: "list-flat",
    CommandButton: "symbol-event",
    CheckBox: "check",
    OptionButton: "circle",
    ToggleButton: "symbol-boolean",
    SpinButton: "arrow-swap",
    ScrollBar: "arrow-both",
    Image: "file-media",
    Frame: "symbol-class",
    MultiPage: "multiple-windows",
    TabStrip: "layout",
  };
  return (
    <aside
      class="toolbox"
      aria-label={strings.toolbox}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onSelect(undefined);
        }
      }}
    >
      <h2>{strings.toolbox}</h2>
      <button type="button" aria-pressed={!selected} onClick={() => onSelect(undefined)}>
        <Icon name="cursor" />
        {strings.pointer}
      </button>
      {(["commonControls", "containers"] as const).map((group) => (
        <section class="toolbox-group" key={group}>
          <h3>{strings[group]}</h3>
          {controlMetadata
            .filter(
              (control) =>
                ["Frame", "MultiPage", "TabStrip"].includes(control.type) ===
                (group === "containers"),
            )
            .map((control) => (
              <button
                type="button"
                key={control.type}
                data-toolbox-type={control.type}
                aria-pressed={selected === control.type}
                disabled={
                  disabled ||
                  (control.type === "MultiPage" && !pageEditable) ||
                  (control.type === "TabStrip" && !tabEditable)
                }
                onClick={() => onSelect(control.type)}
              >
                <Icon name={icons[control.type]} />
                {control.type}
              </button>
            ))}
        </section>
      ))}
      {disabled && <p class="notice">{strings.structuralReadOnly}</p>}
      {selected && (
        <p class="notice" role="status">
          {parentEditable
            ? pageEditable
              ? strings.pagePlacement
              : strings.parentPlacement
            : strings.placement}
        </p>
      )}
    </aside>
  );
}
