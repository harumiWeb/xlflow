import type { DesignerControl, DesignerDocument } from "../../src/userFormEditor/protocol";
import { pointsToPixels } from "../../src/userFormEditor/model";
const px = (points: number) => `${pointsToPixels(points)}px`;

export function ControlGlyph({ control }: { control: DesignerControl }) {
  const type = control.type.toLowerCase();
  const caption = control.caption ?? control.name;
  switch (type) {
    case "label":
    case "commandbutton":
    case "togglebutton":
      return <span>{caption}</span>;
    case "textbox":
      return <span>{control.text ?? String(control.value ?? "")}</span>;
    case "checkbox":
      return <span>☐ {caption}</span>;
    case "optionbutton":
      return <span>◯ {caption}</span>;
    case "combobox":
      return (
        <>
          <span class="combo-text">{control.text ?? control.list?.[0] ?? ""}</span>
          <span class="combo-arrow" aria-hidden="true">
            <span class="combo-chevron" />
          </span>
        </>
      );
    case "listbox":
      return (
        <span>
          {(control.list ?? []).map((item, index) => (
            <div key={index}>{item}</div>
          ))}
        </span>
      );
    case "spinbutton":
      return (
        <span class="spin">
          ▴<br />▾
        </span>
      );
    case "scrollbar":
      return <span class="scrollbar">◂ ━ ▸</span>;
    case "image":
      return <span class="placeholder">▧ Image</span>;
    case "frame":
      return <span class="frame-caption">{caption}</span>;
    case "multipage":
      return null;
    case "page":
      return null;
    case "tabstrip":
      return (
        <div class="tabs">
          {(control.tabs ?? []).map(
            (tab, index) =>
              tab.visible !== false && (
                <span key={tab.name} class={index === (control.selectedIndex ?? 0) ? "active" : ""}>
                  {tab.caption ?? tab.name}
                </span>
              ),
          )}
        </div>
      );
    default:
      return (
        <span class="placeholder">
          {control.type}: {control.name}
        </span>
      );
  }
}

export function DesignerCanvas({ document }: { document: DesignerDocument }) {
  const children = (parentId?: string): DesignerControl[] =>
    document.controls
      .filter((control) => (control.parentId || undefined) === parentId)
      .map((control, index) => ({ control, index }))
      .sort((a, b) => (a.control.zIndex ?? a.index) - (b.control.zIndex ?? b.index))
      .map(({ control }) => control);
  const draw = (control: DesignerControl) => {
    if (control.visible === false) return null;
    const type = control.type.toLowerCase();
    const pages = type === "multipage" ? children(control.id) : [];
    const selected = control.selectedIndex ?? (pages.length ? 0 : -1);
    const page = pages[selected];
    return (
      <div
        key={control.id}
        class={`control ${type} ${control.enabled === false ? "disabled" : ""}`}
        title={`${control.name} (${control.type})${control.approximate ? " — approximate bounds" : ""}`}
        style={{
          left: px(control.left),
          top: px(control.top),
          width: px(control.width),
          height: px(control.height),
        }}
      >
        <ControlGlyph control={control} />
        {type === "multipage" ? (
          <>
            <div class="tabs">
              {pages.map(
                (p, index) =>
                  p.visible !== false && (
                    <span key={p.id} class={index === selected ? "active" : ""}>
                      {p.caption ?? p.name}
                    </span>
                  ),
              )}
            </div>
            {page && page.visible !== false && (
              <div class="control page page-content" title={page.name}>
                {children(page.id).map(draw)}
              </div>
            )}
          </>
        ) : (
          children(control.id).map(draw)
        )}
      </div>
    );
  };
  return (
    <div class="form-shell">
      <div class="form-title">{document.caption}</div>
      <div class="form" style={{ width: px(document.width), height: px(document.height) }}>
        {children().map(draw)}
      </div>
    </div>
  );
}
