import type { DesignerDocument, DesignerProblem, PreviewResult, SourceRange } from "./protocol";

export function problemForRequest(
  snapshot: { version: number; document: Pick<DesignerDocument, "warnings"> } | undefined,
  currentVersion: number,
  requestVersion: number,
  index: number,
): DesignerProblem | undefined {
  if (
    !snapshot ||
    snapshot.version !== requestVersion ||
    currentVersion !== requestVersion ||
    !Number.isSafeInteger(index) ||
    index < 0
  )
    return undefined;
  return snapshot.document.warnings[index];
}

export async function problemSourceRange(
  problem: DesignerProblem,
  index: number,
  previewText: string,
  sourceText: string,
  sourceVersion: number,
  requestSourcePreview: () => Promise<PreviewResult>,
): Promise<SourceRange | undefined> {
  if (previewText === sourceText) return validRangeInText(problem.range, sourceText);

  let preview: PreviewResult;
  try {
    preview = await requestSourcePreview();
  } catch {
    return undefined;
  }
  if (preview.version !== sourceVersion) return undefined;

  const warnings = preview.warnings ?? [];
  const sameIndex = warnings[index];
  const sourceProblem = sameProblem(problem, sameIndex)
    ? sameIndex
    : warnings.find((warning) => sameProblem(problem, warning));
  return validRangeInText(sourceProblem?.range, sourceText);
}

function sameProblem(left: DesignerProblem, right: DesignerProblem | undefined): boolean {
  if (!right) return false;
  return (
    left.code === right.code &&
    left.field === right.field &&
    left.message === right.message &&
    (left.severity === undefined ||
      right.severity === undefined ||
      left.severity === right.severity)
  );
}

function validRangeInText(range: SourceRange | undefined, text: string): SourceRange | undefined {
  if (!range) return undefined;
  const lines = text.split(/\r\n|\n|\r/);
  const validPosition = (position: SourceRange["start"]) =>
    Number.isSafeInteger(position.line) &&
    Number.isSafeInteger(position.character) &&
    position.line >= 0 &&
    position.line < lines.length &&
    position.character >= 0 &&
    position.character <= lines[position.line].length;
  if (!validPosition(range.start) || !validPosition(range.end)) return undefined;
  if (
    range.end.line < range.start.line ||
    (range.end.line === range.start.line && range.end.character < range.start.character)
  )
    return undefined;
  return range;
}
