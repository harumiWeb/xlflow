import type { ComponentChildren } from "preact";
import type { EditorDescriptor } from "./types";
import { FontEditor } from "./FontEditor";
import { ColorEditor } from "./ColorEditor";

export interface EditorFieldsProps {
  descriptor: EditorDescriptor;
  value: Record<string, unknown>;
  disabled: boolean;
  unset: string;
  sample: string;
  paletteLabel: string;
  systemLabel: string;
  onFieldChange: (field: string, value: unknown, valid?: boolean) => void;
  onValueChange: (value: Record<string, unknown>, valid?: boolean) => void;
}
interface EditorRegistration {
  presentation: "modal" | "nativePicker" | "inline";
  render?: (props: EditorFieldsProps) => ComponentChildren;
  approximate?: boolean;
}

// This registry chooses presentation only. The LSP descriptors own field
// applicability, children, numeric limits, and accepted values.
export const editorRegistry: Record<string, EditorRegistration> = {
  font: {
    presentation: "modal",
    approximate: true,
    render: (props) => <FontEditor {...props} onChange={props.onFieldChange} />,
  },
  color: {
    presentation: "modal",
    approximate: true,
    render: (props) => <ColorEditor {...props} onChange={props.onValueChange} />,
  },
  picture: { presentation: "nativePicker" },
  enum: { presentation: "inline" },
  boolean: { presentation: "inline" },
};
