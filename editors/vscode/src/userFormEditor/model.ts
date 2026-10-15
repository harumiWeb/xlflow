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
  const width = form.width;
  const height = form.height;
  return {
    font: form.font,
    foreColor: form.foreColor,
    backColor: form.backColor,
    borderStyle: form.borderStyle,
    name: form.name,
    caption: form.caption ?? form.name,
    width: width ?? 240,
    height: height ?? 180,
    approximate: width == null || height == null,
    warnings,
    controls: spec.controls.map((control) => {
      const effective = control;
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
