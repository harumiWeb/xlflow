import type { ComponentChildren } from "preact";
import { useEffect, useRef } from "preact/hooks";

interface Props {
  title: string;
  children: ComponentChildren;
  disabled: boolean;
  pending: boolean;
  cancelLabel: string;
  okLabel: string;
  onCancel: () => void;
  onConfirm: () => void;
  returnFocus?: HTMLElement;
}
export function PropertyModal(props: Props) {
  const panel = useRef<HTMLDivElement>(null);
  const cancel = useRef(props.onCancel);
  cancel.current = props.onCancel;
  useEffect(() => {
    const previous = props.returnFocus ?? (document.activeElement as HTMLElement | null);
    const element = panel.current!;
    (
      element.querySelector<HTMLElement>("[data-initial-focus], input, select, button") ?? element
    ).focus();
    const keepFocus = (event: FocusEvent) => {
      if (event.target instanceof window.Node && !element.contains(event.target))
        (
          element.querySelector<HTMLElement>(
            "input:not(:disabled), select:not(:disabled), button:not(:disabled)",
          ) ?? element
        ).focus();
    };
    document.addEventListener("focusin", keepFocus);
    return () => {
      document.removeEventListener("focusin", keepFocus);
      queueMicrotask(() => {
        if (previous?.isConnected) previous.focus({ preventScroll: true });
      });
    };
  }, []);
  return (
    <div class="property-modal-backdrop" onPointerDown={(event) => event.stopPropagation()}>
      <div
        ref={panel}
        class="property-modal"
        role="dialog"
        aria-modal="true"
        aria-label={props.title}
        tabIndex={-1}
        onKeyDown={(event) => {
          if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
          if (event.key === "Escape") {
            event.preventDefault();
            event.stopPropagation();
            if (!props.pending) cancel.current();
            return;
          }
          if (event.key === "Tab") {
            const elements = Array.from(
              panel.current?.querySelectorAll<HTMLElement>(
                "input:not(:disabled), select:not(:disabled), textarea:not(:disabled), button:not(:disabled), [tabindex='0']",
              ) ?? [],
            );
            const first = elements[0],
              last = elements.at(-1);
            if (!first) {
              event.preventDefault();
              panel.current?.focus();
            } else if (
              event.shiftKey &&
              (document.activeElement === first || document.activeElement === panel.current)
            ) {
              event.preventDefault();
              last?.focus();
            } else if (!event.shiftKey && document.activeElement === last) {
              event.preventDefault();
              first.focus();
            }
            return;
          }
          if (
            event.key === "Enter" &&
            !event.altKey &&
            !event.ctrlKey &&
            !event.metaKey &&
            !(event.target as Element).closest("button, select, textarea, [data-handles-enter]")
          ) {
            event.preventDefault();
            event.stopPropagation();
            if (!props.disabled && !props.pending) props.onConfirm();
          }
        }}
      >
        <h2>{props.title}</h2>
        <div class="property-modal-body">{props.children}</div>
        <div class="property-modal-actions">
          <button type="button" onClick={props.onCancel} disabled={props.pending}>
            {props.cancelLabel}
          </button>
          <button
            class="primary-action"
            type="button"
            onClick={props.onConfirm}
            disabled={props.disabled || props.pending}
          >
            {props.okLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
