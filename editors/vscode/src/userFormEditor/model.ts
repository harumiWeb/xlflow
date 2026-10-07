import type { DesignerDocument, PreviewDocument } from "./protocol";
import { controlDefaults } from "./toolbox";

export function pointsToPixels(points: number, zoom = 1): number {
  return (points * 96 * zoom) / 72;
}

export function pixelsToPoints(pixels: number, zoom = 1): number {
  return (pixels * 72) / (96 * zoom);
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
      const defaults = controlDefaults(control.type);
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
