import type { DesignerStrings } from "../../../src/userFormEditor/protocol";
import { propertyEditorStrings } from "../../../src/userFormEditor/propertyEditorStrings";

export const editorStrings = propertyEditorStrings;
export function localizedEditorStrings(strings: DesignerStrings) {
  return { ...editorStrings, ...strings };
}
