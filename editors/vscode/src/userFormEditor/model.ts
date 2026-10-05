import type { DesignerDocument, PreviewDocument } from "./protocol";

// These are display-only defaults from the pure-Go generation contract.
const sizes: Record<string, [number, number]> = {
  label: [72, 18],
  textbox: [120, 18],
  commandbutton: [72, 24],
  checkbox: [72, 18],
  optionbutton: [72, 18],
  togglebutton: [72, 18],
  combobox: [120, 18],
  listbox: [120, 72],
  spinbutton: [18, 36],
  scrollbar: [120, 18],
  image: [72, 72],
  frame: [144, 108],
  multipage: [240, 180],
  page: [0, 0],
  tabstrip: [240, 48],
};

export function pointsToPixels(points: number): number {
  return (points * 96) / 72;
}

export function designerDocument(spec: PreviewDocument, warnings: string[] = []): DesignerDocument {
  const form = spec.form;
  const build = form.build ?? {};
  const observed = form.observed ?? {};
  const width =
    build.clientWidth ??
    observed.clientWidth ??
    observed.insideWidth ??
    build.width ??
    form.width ??
    observed.width;
  const height =
    build.clientHeight ??
    observed.clientHeight ??
    observed.insideHeight ??
    build.height ??
    form.height ??
    observed.height;
  return {
    name: form.name,
    caption: build.caption ?? form.caption ?? observed.caption ?? form.name,
    width: width ?? 240,
    height: height ?? 180,
    approximate:
      width === undefined ||
      height === undefined ||
      (build.clientWidth === undefined &&
        observed.clientWidth === undefined &&
        observed.insideWidth === undefined) ||
      (build.clientHeight === undefined &&
        observed.clientHeight === undefined &&
        observed.insideHeight === undefined),
    warnings,
    controls: spec.controls.map((control) => {
      const effective = { ...control.observed, ...control };
      const defaults = sizes[control.type.toLowerCase()] ?? [72, 24];
      return {
        ...effective,
        left: effective.left ?? 0,
        top: effective.top ?? 0,
        width: effective.width ?? defaults[0],
        height: effective.height ?? defaults[1],
        approximate: effective.width === undefined || effective.height === undefined,
      };
    }),
  };
}
