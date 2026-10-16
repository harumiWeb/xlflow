export function Icon({ name }: { name: string }) {
  const codiconName = name === "undo" ? "discard" : name;
  return <span class={`codicon codicon-${codiconName}`} aria-hidden="true" />;
}
