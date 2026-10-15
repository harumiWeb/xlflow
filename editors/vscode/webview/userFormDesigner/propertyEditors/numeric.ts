import type { PropertyDescriptor } from "../../../src/userFormEditor/protocol";
import { numberRoundTrips } from "../numbers";

export function validNumericInput(text: string, descriptor: PropertyDescriptor): boolean {
  const number = Number(text);
  return (
    text.trim() !== "" &&
    Number.isFinite(number) &&
    numberRoundTrips(text, number) &&
    (descriptor.valueType !== "integer" || Number.isSafeInteger(number)) &&
    (descriptor.minimum === undefined ||
      (descriptor.exclusiveMinimum ? number > descriptor.minimum : number >= descriptor.minimum)) &&
    (descriptor.maximum === undefined || number <= descriptor.maximum)
  );
}
