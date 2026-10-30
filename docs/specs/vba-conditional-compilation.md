# Conditional-compilation analysis

Lint, analyze and LSP diagnostic parsing use a shared conditional-compilation
projection. Before parsing, provably inactive content is replaced by spaces
with exactly the same byte length. CR, LF and active source positions remain
unchanged. Resolved conditional wrappers are also replaced by spaces so a
selected declaration header can join its shared body. Unknown wrappers remain
available to structure and uncertainty checks. Original files, formatter input and workbook import source
are not rewritten by this projection.

## Constant environment

Project arguments can be declared in `xlflow.toml`:

```toml
[vba.conditional_constants]
DEBUG = "False"
FeatureLevel = "2"
VBA7 = "True"
Win64 = "True"
Win32 = "True"
Mac = "False"
```

Keys are case-insensitive VBA identifiers; duplicate case variants are errors.
Values are TOML strings containing independently evaluable VBA literal
expressions. Boolean, integral, floating-point, currency, string and Empty
values are accepted. Date, Null, Nothing, unresolved names and unsupported
expressions are configuration errors. The table is unordered: one entry
cannot reference another. No compiler arguments are read from or written to
the workbook by analysis. Match this table to the target VBA project and
restart the LSP after changing project configuration.

An active module-level `#Const` is evaluated in source order and overrides the
project value within that module. It does not leak into other modules.
Inactive `#Const` declarations do not affect later conditions. Conditional
assignments are merged at `#End If`; different possible values remain unknown.

Undefined ordinary identifiers evaluate as VBA Empty, which coerces to zero.
The target built-ins `VBA6`, `VBA7`, `Win16`, `Win32`, `Win64` and `Mac` remain
unknown unless explicitly configured. Analysis never guesses target constants
from the OS or architecture of the xlflow process.

## Evaluation and fallback

The explicit conditional evaluator reuses the shared constant-expression
parser. It supports Boolean and numeric conditions, arithmetic, comparisons,
Boolean/Empty numeric coercions and integral logical operators. `Not` binds
after comparisons and before `And`. These coercions do not change ordinary
declaration-constant evaluation.

Decimal Currency literals with at most four fractional digits retain exact
scaled integer values, including comparisons between adjacent large values.
Currency literal rounding, exponent notation, arithmetic and exponentiation
remain unknown rather than selecting a branch through floating-point rounding.

Unsupported coercions, runtime calls, string comparisons, signed-boundary
hexadecimal/octal literal widths, Single arithmetic rounding, mixed-precision
Single/Currency comparisons, wide integer comparison precision and fractional
bitwise operands remain
unknown. Both candidate branches of an unknown condition remain analyzable.
Malformed or unclosed conditional directive structures, and unsupported
continued directive headers, conservatively retain the original document so
syntax diagnostics cannot be hidden. Source-encoding validation also remains
applicable to the complete authored source, including inactive branches.

This change preserves conservative handling of unresolved conditional
declarations in resolution, CFG and dataflow. It does not turn an unknown
target into confirmed VBE compilation evidence. A retained candidate branch
with unsupported syntax may therefore still produce VB014 or analyze_failed.

VBA266 keeps its existing conservative lexical usedness policy: mentions in
the authored source, including conditional directives and inactive branches,
can suppress an unused-private-constant warning. Its declaration candidates
still come from the projected active source.

## Regression contracts

The tests cover inactive prose from Issue #966, active errors, nested and
alternate branches, case-insensitive project constants, module scope and
override behavior, uncertain assignment joins, malformed structures, CRLF and
Unicode byte positions, and unchanged raw source. Editor coverage also
protects changes to conditional constants in unsaved buffers and prevents
stale parse artifacts from surviving a branch switch.

The local VBE oracle supplies language observations separately from parser
recovery diagnostics. Oracle infrastructure failures never justify fixture
promotion or analyzer semantic changes.
