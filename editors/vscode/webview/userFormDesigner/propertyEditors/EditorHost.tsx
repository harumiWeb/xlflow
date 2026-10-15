import { useEffect, useRef, useState } from "preact/hooks";
import type { DesignerStrings, PropertyValue } from "../../../src/userFormEditor/protocol";
import type { PropertyReply } from "../PropertyGrid";
import type { EditorSession, EditorPreview, PictureSelection } from "./types";
import { objectValue } from "./types";
import { PropertyModal } from "./Modal";
import { editorRegistry } from "./registry";
import { localizedEditorStrings } from "./strings";
import { validNumericInput } from "./numeric";

export interface PictureReply {
  requestId: number;
  selection?: PictureSelection;
  error?: { code: string; message: string };
}
interface Props {
  session: EditorSession;
  version: number;
  epoch: number;
  busy: boolean;
  strings: DesignerStrings;
  reply?: PropertyReply;
  pictureReply?: PictureReply;
  onClose: () => void;
  onPreview: (preview?: EditorPreview) => void;
  onCommit: (value: PropertyValue) => number | undefined;
  onPickPicture?: () => number | undefined;
  onCommitPicture?: (selection: PictureSelection) => number | undefined;
  onReleasePicture?: (token: string) => void;
}
export function PropertyEditorHost(props: Props) {
  const { session } = props;
  const strings = localizedEditorStrings(props.strings);
  const [value, setValue] = useState(() => objectValue(session.value));
  const [invalid, setInvalid] = useState(new Set<string>());
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<string>();
  const [submitted, setSubmitted] = useState<number>();
  const [picking, setPicking] = useState<number>();
  const [picture, setPicture] = useState<PictureSelection>();
  const callbacks = useRef(props);
  callbacks.current = props;
  const token = useRef<string>();
  const committed = useRef(false);
  const stale = session.version !== props.version || session.epoch !== props.epoch;
  const pending = submitted !== undefined;
  const isPosition = session.descriptor.field === "$geometry";
  const registration = session.descriptor.editor
    ? editorRegistry[session.descriptor.editor]
    : undefined;
  useEffect(() => {
    if (stale || invalid.size > 0) callbacks.current.onPreview(undefined);
    else
      callbacks.current.onPreview({
        controlId: session.controlId,
        values: isPosition
          ? value
          : {
              [session.descriptor.field]: value,
              ...(session.descriptor.editor === "picture" && (picture || value.remove === true)
                ? { picturePreview: value.remove === true ? "" : picture?.dataURI }
                : {}),
            },
      });
    return () => callbacks.current.onPreview(undefined);
  }, [value, stale, invalid, picture]);
  useEffect(
    () => () => {
      if (token.current && !committed.current) callbacks.current.onReleasePicture?.(token.current);
    },
    [],
  );
  useEffect(() => {
    if (submitted === undefined || props.reply?.requestId !== submitted) return;
    if (props.reply.error) {
      const diagnostics = props.reply.error.diagnostics;
      setError(
        diagnostics?.length
          ? diagnostics
              .map(
                (diagnostic) =>
                  `${diagnostic.code}: ${diagnostic.message}${diagnostic.suggestion ? ` ${diagnostic.suggestion}` : ""}`,
              )
              .join("\n")
          : `${props.reply.error.code}: ${props.reply.error.message}`,
      );
      setSubmitted(undefined);
    } else {
      committed.current = true;
      callbacks.current.onClose();
    }
  }, [props.reply]);
  useEffect(() => {
    if (picking === undefined || props.pictureReply?.requestId !== picking) return;
    setPicking(undefined);
    if (props.pictureReply.error) {
      setError(`${props.pictureReply.error.code}: ${props.pictureReply.error.message}`);
      return;
    }
    const selection = props.pictureReply.selection;
    if (selection) {
      token.current = selection.token;
      setPicture(selection);
      setValue({ path: selection.name });
      setDirty(true);
      setError(undefined);
    }
  }, [props.pictureReply]);
  const changeField = (field: string, next: unknown, valid = true) => {
    setValue((current) => ({ ...current, [field]: next }));
    setInvalid((current) => {
      const result = new Set(current);
      if (valid) result.delete(field);
      else result.add(field);
      return result;
    });
    setDirty(true);
    setError(undefined);
  };
  const confirm = () => {
    if (stale || pending || props.busy || picking !== undefined || invalid.size > 0) return;
    if (
      !dirty ||
      (!picture && JSON.stringify(value) === JSON.stringify(objectValue(session.value)))
    ) {
      props.onClose();
      return;
    }
    const request =
      picture && value.remove !== true && props.onCommitPicture
        ? props.onCommitPicture(picture)
        : props.onCommit(value as PropertyValue);
    if (request !== undefined) setSubmitted(request);
  };
  const disabled = pending || props.busy || stale;
  return (
    <PropertyModal
      title={isPosition ? strings.editorPosition : session.title}
      returnFocus={session.returnFocus}
      disabled={stale || props.busy || picking !== undefined || invalid.size > 0}
      pending={pending}
      cancelLabel={strings.editorCancel}
      okLabel={strings.editorOK}
      onCancel={props.onClose}
      onConfirm={confirm}
    >
      {stale && (
        <p role="alert" class="property-error">
          {strings.editorStale}
        </p>
      )}
      {error && (
        <p role="alert" class="property-error">
          {error}
        </p>
      )}
      {invalid.size > 0 && (
        <p role="alert" class="property-error">
          {strings.editorInvalid}
        </p>
      )}
      {registration?.render?.({
        descriptor: session.descriptor,
        value,
        disabled,
        unset: props.strings.unset,
        sample: strings.editorFontSample,
        paletteLabel: strings.editorPalette,
        systemLabel: strings.editorSystemColors,
        onFieldChange: changeField,
        onValueChange: (next, valid = true) => {
          setValue(next);
          setInvalid(valid ? new Set() : new Set(["value"]));
          setDirty(true);
          setError(undefined);
        },
      })}
      {isPosition && (
        <div class="property-editor-fields">
          {(session.descriptor.children ?? []).map((child) => (
            <label class="property-editor-field" for={`geometry-${child.field}`} key={child.field}>
              <span>{child.field[0].toUpperCase() + child.field.slice(1)}</span>
              <input
                id={`geometry-${child.field}`}
                type="text"
                inputMode="decimal"
                disabled={disabled}
                value={
                  typeof value[child.field] === "number" || typeof value[child.field] === "string"
                    ? String(value[child.field])
                    : ""
                }
                placeholder={props.strings.unset}
                onInput={(event) => {
                  const text = event.currentTarget.value,
                    number = Number(text);
                  const valid = validNumericInput(text, child);
                  changeField(child.field, valid ? number : text, valid);
                }}
              />
            </label>
          ))}
        </div>
      )}
      {session.descriptor.editor === "picture" && (
        <div class="picture-editor">
          <p>{strings.editorPictureHelp}</p>
          {(picture?.dataURI ?? session.picturePreview) && value.remove !== true && (
            <img
              src={picture?.dataURI ?? session.picturePreview}
              alt={picture?.name ?? String(value.path ?? "Picture")}
            />
          )}
          <p class="picture-editor-path">
            {typeof value.path === "string" ? value.path : props.strings.unset}
          </p>
          {picture && value.remove !== true && (
            <p>
              {picture.width} × {picture.height} px
            </p>
          )}
          <button
            type="button"
            disabled={disabled || picking !== undefined}
            onClick={() => {
              setError(undefined);
              setPicking(props.onPickPicture?.());
            }}
          >
            {picking !== undefined ? strings.editorPicturePending : strings.editorChoosePicture}
          </button>
          <button
            type="button"
            disabled={disabled || picking !== undefined}
            onClick={() => {
              setValue({ remove: true });
              setDirty(true);
              setError(undefined);
            }}
          >
            {strings.editorRemovePicture}
          </button>
        </div>
      )}
      {registration?.approximate && <p class="notice">{strings.editorApproximate}</p>}
    </PropertyModal>
  );
}
