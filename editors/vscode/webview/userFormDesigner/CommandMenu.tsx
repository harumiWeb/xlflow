import { useLayoutEffect, useRef, useState } from "preact/hooks";
import type { ComponentChildren } from "preact";
import { Icon } from "./Icon";

// Menus use viewport coordinates so neither the toolbar nor a scrolling pane
// clips them. The same keyboard/focus behavior is used for each command group.
export function CommandMenu({ label, children }: { label: string; children: ComponentChildren }) {
  const [open, setOpen] = useState(false);
  const trigger = useRef<HTMLButtonElement>(null);
  const menu = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (!open || !menu.current || !trigger.current) return;
    const popup = menu.current;
    const anchor = trigger.current.getBoundingClientRect();
    const bounds = popup.getBoundingClientRect();
    popup.style.left = `${Math.max(4, Math.min(anchor.left, window.innerWidth - bounds.width - 4))}px`;
    popup.style.top = `${Math.max(4, Math.min(anchor.bottom, window.innerHeight - bounds.height - 4))}px`;
    const items = () =>
      Array.from(
        popup.querySelectorAll<HTMLButtonElement>('button[role="menuitem"]:not(:disabled)'),
      );
    items()[0]?.focus();
    const close = () => setOpen(false);
    const outside = (event: PointerEvent) => {
      if (!popup.contains(event.target as Node) && !trigger.current?.contains(event.target as Node))
        close();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        close();
        trigger.current?.focus();
      } else if (
        popup.contains(event.target as Node) &&
        ["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)
      ) {
        event.preventDefault();
        const enabled = items();
        if (!enabled.length) return;
        const index = enabled.indexOf(document.activeElement as HTMLButtonElement);
        const next =
          event.key === "Home"
            ? 0
            : event.key === "End"
              ? enabled.length - 1
              : (index + (event.key === "ArrowDown" ? 1 : -1) + enabled.length) % enabled.length;
        enabled[next]?.focus();
      } else if (event.key === "Tab") close();
    };
    window.addEventListener("pointerdown", outside, true);
    window.addEventListener("keydown", keydown, true);
    window.addEventListener("resize", close);
    window.addEventListener("scroll", close, true);
    return () => {
      window.removeEventListener("pointerdown", outside, true);
      window.removeEventListener("keydown", keydown, true);
      window.removeEventListener("resize", close);
      window.removeEventListener("scroll", close, true);
    };
  }, [open]);
  return (
    <div class="command-menu">
      <button
        ref={trigger}
        type="button"
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        {label}
        <Icon name="chevron-down" />
      </button>
      {open && (
        <div
          ref={menu}
          class="designer-context-menu"
          role="menu"
          aria-label={label}
          onClick={(event) => {
            if ((event.target as Element).closest('button[role="menuitem"]:not(:disabled)')) {
              setOpen(false);
              trigger.current?.focus();
            }
          }}
        >
          {children}
        </div>
      )}
    </div>
  );
}
