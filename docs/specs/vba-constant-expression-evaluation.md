# VBA Constant-Expression Evaluation

This specification defines the internal, Excel-free contract for deciding
whether a VBA expression has a safe static value. It is used by declaration
and compile-equivalent analysis; it is not a VBA interpreter and does not
change CLI or LSP wire formats.

## API

`internal/vba/constexpr` provides:

```go
type Result struct {
    Kind   ResultKind // Known, Unknown, or Invalid
    Value  int        // compatibility projection for integral Known values
    Typed  Value      // typed Known value
    Type   string
    Reason string
}

func Evaluate(expression string, env Environment) Result
func EvaluateValues(expression string, values map[string]Value) Result
func EvaluateInteger(expression string, constants map[string]int) Result
```

`EvaluateInteger` is a compatibility adapter. It returns `Known` only when the
typed result is integral and representable by the host `int`; otherwise it
returns `Unknown`.

## Supported subset

- decimal, hexadecimal (`&H`), octal (`&O`), floating-point, currency, and
  VBA numeric suffixes;
- quoted strings with doubled-quote escaping, date literals, `True`/`False`,
  `Empty`, `Null`, and `Nothing`;
- values supplied by an immutable environment, including Const, Enum, and
  static TypeLib/intrinsic constants;
- parentheses, unary `+`, `-`, `Not`, arithmetic (`+`, `-`, `*`, `/`, `\`,
  `Mod`, `^`), Boolean (`And`, `Or`, `Xor`, `Eqv`, `Imp`), comparisons, and
  string concatenation (`&`).

Operator precedence follows the evaluator grammar and is covered by fixtures.
`Like` and `Is` are recognized but remain `Unknown` because their semantics are
context-dependent or object/runtime-based.

## Outcome policy

`Known` means all operands and operations were modelled safely. `Unknown` is
used for unresolved or ambiguous symbols, runtime calls, conditional or
recovered source, unsupported semantics, non-integral compatibility results,
and overflow whose VBA behavior is not modelled. `Invalid` is reserved for
malformed syntax, incompatible statically known operand types, and invalid
domains such as division by zero. Consumers must not turn `Unknown` into an
error.

The evaluator uses checked integer operations and finite floating-point
results. Currency values retain four-decimal fixed-point storage. Environment
lookups are case-insensitive and deterministic even if a caller's map contains
case-only duplicate keys.

## Environment construction

Batch and LSP project snapshots build value maps from the parsed source and
static TypeLib database. Only public/friend declarations are exported across
modules. Qualified module/Enum/library names are retained; unqualified names
are retained only for a unique winner. A document's local Const/Enum values are
overlaid on the immutable project map for that document. Conditional,
recovered, unresolved, or cyclic declarations remain absent from the value map
and therefore evaluate as `Unknown`.

## Consumers and verification

- Optional parameter defaults use the typed result for existing `VB048`
  constant/type checks and keep runtime calls as nonconstant.
- Fixed declaration bounds, `ReDim` bounds, and compile-equivalent array
  adapters use the integer compatibility projection; dynamic or non-integral
  bounds remain unclassified.
- Dedicated evaluator fixtures cover suffixes, numeric boundaries,
  floating-point and currency literals, strings, Booleans, Const/Enum
  references, precedence, invalid expressions, unresolved names, and calls.
- Lint, analyze, Intel, and LSP tests verify existing diagnostic ownership and
  deterministic behavior without Excel or COM.
