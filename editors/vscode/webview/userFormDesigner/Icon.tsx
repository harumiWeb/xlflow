export function Icon({ name }: { name: string }) {
  return <span class={`codicon codicon-${name}`} aria-hidden="true" />;
}
