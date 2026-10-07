import { controlMetadata } from "../../src/userFormEditor/toolbox";
import type { ToolboxType } from "../../src/userFormEditor/toolbox";
import type { DesignerStrings } from "../../src/userFormEditor/protocol";

export function Toolbox({
  selected,
  disabled,
  strings,
  onSelect,
}: {
  selected?: ToolboxType;
  disabled: boolean;
  strings: DesignerStrings;
  onSelect: (type?: ToolboxType) => void;
}) {
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
        {strings.pointer}
      </button>
      {controlMetadata.map((control) => (
        <button
          type="button"
          key={control.type}
          data-toolbox-type={control.type}
          aria-pressed={selected === control.type}
          disabled={disabled}
          onClick={() => onSelect(control.type)}
        >
          {control.type}
        </button>
      ))}
      {disabled && <p class="notice">{strings.structuralReadOnly}</p>}
      {selected && (
        <p class="notice" role="status">
          {strings.placement}
        </p>
      )}
    </aside>
  );
}
