// Decimal equality across Number()/JSON.stringify(), including exponent syntax.
// Compare shortest decimal representations, not binary rational values (0.1 is valid).
function decimalKey(text: string): string | undefined {
  const match = /^([+-]?)(?:(\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/.exec(text.trim());
  if (!match) return undefined;
  const fraction = match[3] ?? match[4] ?? "";
  let digits = ((match[2] ?? "") + fraction).replace(/^0+/, "");
  if (!digits) return "0";
  const trimmed = digits.replace(/0+$/, "");
  const exponent =
    BigInt(match[5] ?? "0") - BigInt(fraction.length) + BigInt(digits.length - trimmed.length);
  digits = trimmed;
  return `${match[1] === "-" ? "-" : ""}${digits}e${exponent}`;
}

export function numberRoundTrips(text: string, value: number): boolean {
  const original = decimalKey(text);
  return original !== undefined && Number.isFinite(value) && original === decimalKey(String(value));
}
