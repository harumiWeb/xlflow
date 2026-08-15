# Static-analysis diagnostic catalog

Generated from the canonical rule registry at `internal/staticanalysis/rules/registry.json`. Run `pnpm docs:generate-reference` after changing rule metadata. Do not edit this page by hand.

Use [`xlflow rules`](../commands/rules) to inspect the same metadata from an installed xlflow binary. `VBA000` is a synthetic analysis-failure diagnostic and is intentionally outside the registry; UserForm `FRM...` and `UFY...` diagnostics are outside this catalog.

`Blocks source preflight` describes the registry default. A project may list a blocking diagnostic in `[preflight].allowed_diagnostics` to let workbook automation proceed without disabling the diagnostic or changing its severity; applied waivers are reported as command warnings.

| ID                  | Family  | Severity    | Scope           | Default | Title                                              |
| ------------------- | ------- | ----------- | --------------- | ------- | -------------------------------------------------- |
| [`VB001`](#vb001)   | lint    | warning     | file-local      | yes     | Missing Option Explicit                            |
| [`VB002`](#vb002)   | lint    | warning     | procedure-local | yes     | Select usage                                       |
| [`VB003`](#vb003)   | lint    | warning     | procedure-local | yes     | Activate usage                                     |
| [`VB004`](#vb004)   | lint    | warning     | procedure-local | yes     | Broad On Error Resume Next                         |
| [`VB005`](#vb005)   | lint    | warning     | file-local      | yes     | Implicit Variant                                   |
| [`VB006`](#vb006)   | lint    | warning     | file-local      | yes     | Public module field                                |
| [`VB007`](#vb007)   | lint    | warning     | file-local      | yes     | Automation-hostile GUI boundary                    |
| [`VB008`](#vb008)   | lint    | error       | file-local      | yes     | Typographic quote                                  |
| [`VB009`](#vb009)   | lint    | error       | file-local      | yes     | C-style quote escape                               |
| [`VB010`](#vb010)   | lint    | error       | file-local      | yes     | Unterminated procedure                             |
| [`VB011`](#vb011)   | lint    | error       | file-local      | yes     | Unexpected procedure terminator                    |
| [`VB012`](#vb012)   | lint    | error       | file-local      | yes     | Mismatched procedure terminator                    |
| [`VB013`](#vb013)   | lint    | error       | file-local      | yes     | Invalid line continuation                          |
| [`VB014`](#vb014)   | lint    | error       | file-local      | yes     | Parser recovery                                    |
| [`VB015`](#vb015)   | lint    | error       | file-local      | yes     | Continuation limit exceeded                        |
| [`VB018`](#vb018)   | lint    | warning     | project-wide    | no      | Scope shadowing                                    |
| [`VB019`](#vb019)   | lint    | warning     | file-local      | yes     | Mixed declarator typing                            |
| [`VB020`](#vb020)   | lint    | warning     | procedure-local | yes     | Unused local variable                              |
| [`VB021`](#vb021)   | lint    | warning     | project-wide    | no      | Unused private procedure                           |
| [`VB022`](#vb022)   | lint    | warning     | procedure-local | yes     | Confusing call syntax                              |
| [`VB023`](#vb023)   | lint    | warning     | procedure-local | yes     | Invalid For Each control type                      |
| [`VB026`](#vb026)   | lint    | warning     | procedure-local | yes     | Dangerous Resume                                   |
| [`VB027`](#vb027)   | lint    | warning     | procedure-local | no      | Ambiguous nested With member                       |
| [`VB028`](#vb028)   | lint    | error       | project-wide    | yes     | Bare dialog call with XlflowUI                     |
| [`VB029`](#vb029)   | lint    | error       | project-wide    | yes     | Undeclared variable                                |
| [`VB030`](#vb030)   | lint    | warning     | procedure-local | yes     | Argument mismatch                                  |
| [`VB031`](#vb031)   | lint    | error       | file-local      | yes     | Missing module name attribute                      |
| [`VB032`](#vb032)   | lint    | error       | file-local      | yes     | Repeated Debug.Print shorthand                     |
| [`VB033`](#vb033)   | lint    | warning     | procedure-local | yes     | Unknown member                                     |
| [`VB034`](#vb034)   | lint    | warning     | procedure-local | yes     | Read-only property assignment                      |
| [`VB035`](#vb035)   | lint    | warning     | procedure-local | yes     | Write-only property read                           |
| [`VB036`](#vb036)   | lint    | warning     | procedure-local | yes     | Set required                                       |
| [`VB037`](#vb037)   | lint    | error       | procedure-local | yes     | Set not allowed                                    |
| [`VB038`](#vb038)   | lint    | warning     | procedure-local | yes     | Incompatible assignment                            |
| [`VB039`](#vb039)   | lint    | warning     | procedure-local | yes     | Method has no return value                         |
| [`VB040`](#vb040)   | lint    | warning     | file-local      | yes     | Unknown documented parameter                       |
| [`VB041`](#vb041)   | lint    | warning     | file-local      | yes     | Duplicate documented parameter                     |
| [`VB042`](#vb042)   | lint    | warning     | file-local      | yes     | Returns documentation on Sub                       |
| [`VB043`](#vb043)   | lint    | warning     | file-local      | yes     | Orphan documentation comment                       |
| [`VB044`](#vb044)   | lint    | warning     | procedure-local | no      | Procedure-name constant mismatch                   |
| [`VB045`](#vb045)   | lint    | error       | procedure-local | yes     | Deterministic argument binding error               |
| [`VB046`](#vb046)   | lint    | error       | file-local      | yes     | Duplicate declaration                              |
| [`VB047`](#vb047)   | lint    | error       | file-local      | yes     | Invalid declaration placement                      |
| [`VB048`](#vb048)   | lint    | error       | procedure-local | yes     | Invalid procedure parameter declaration            |
| [`VB049`](#vb049)   | lint    | error       | file-local      | yes     | Inconsistent Property accessor contract            |
| [`VB050`](#vb050)   | lint    | error       | file-local      | yes     | Invalid module declaration context                 |
| [`VB051`](#vb051)   | lint    | error       | procedure-local | yes     | Invalid Me context                                 |
| [`VB052`](#vb052)   | lint    | error       | interprocedural | yes     | Invalid procedure call target                      |
| [`VB053`](#vb053)   | lint    | error       | interprocedural | yes     | Ambiguous Enum member                              |
| [`VB054`](#vb054)   | lint    | error       | interprocedural | yes     | Undeclared RaiseEvent target                       |
| [`VB055`](#vb055)   | lint    | error       | procedure-local | yes     | Duplicate procedure label                          |
| [`VB056`](#vb056)   | lint    | error       | procedure-local | yes     | Undefined procedure label                          |
| [`VB057`](#vb057)   | lint    | error       | procedure-local | yes     | Mismatched Next control variable                   |
| [`VB058`](#vb058)   | lint    | error       | procedure-local | yes     | Invalid Exit statement                             |
| [`VB059`](#vb059)   | lint    | error       | procedure-local | yes     | Invalid call syntax                                |
| [`VB060`](#vb060)   | lint    | error       | project-wide    | yes     | Assignment to constant                             |
| [`VB061`](#vb061)   | lint    | error       | file-local      | yes     | Invalid constant array bounds                      |
| [`VB062`](#vb062)   | lint    | error       | procedure-local | yes     | Invalid conditional branch syntax                  |
| [`VB063`](#vb063)   | lint    | error       | procedure-local | yes     | Invalid Select/Case branch syntax                  |
| [`VB064`](#vb064)   | lint    | error       | procedure-local | yes     | Invalid Open mode syntax                           |
| [`VB065`](#vb065)   | lint    | error       | procedure-local | yes     | Invalid TypeOf syntax                              |
| [`VB066`](#vb066)   | lint    | warning     | file-local      | yes     | Procedure terminator style mismatch                |
| [`VBA101`](#vba101) | analyze | warning     | procedure-local | yes     | Object assignment missing Set                      |
| [`VBA102`](#vba102) | analyze | warning     | procedure-local | yes     | Object-returning call assignment missing Set       |
| [`VBA103`](#vba103) | analyze | warning     | procedure-local | yes     | Object function return missing Set                 |
| [`VBA104`](#vba104) | analyze | error       | procedure-local | yes     | Excel object member mismatch                       |
| [`VBA105`](#vba105) | analyze | error       | project-wide    | yes     | Removed XlflowLog helper                           |
| [`VBA106`](#vba106) | analyze | error       | project-wide    | yes     | Removed XlflowSetTraceFile helper                  |
| [`VBA201`](#vba201) | analyze | warning     | procedure-local | yes     | Unchecked Range.Find result                        |
| [`VBA202`](#vba202) | analyze | warning     | interprocedural | yes     | Object use before Set                              |
| [`VBA203`](#vba203) | analyze | warning     | interprocedural | yes     | Application state not restored                     |
| [`VBA204`](#vba204) | analyze | warning     | procedure-local | yes     | Error-handler fallthrough                          |
| [`VBA205`](#vba205) | analyze | warning     | procedure-local | yes     | Ambiguous Excel object scope                       |
| [`VBA206`](#vba206) | analyze | warning     | interprocedural | yes     | Unsafe ByRef argument                              |
| [`VBA207`](#vba207) | analyze | warning     | procedure-local | no      | Unguarded keyed access                             |
| [`VBA208`](#vba208) | analyze | warning     | procedure-local | yes     | Invalid ReDim Preserve dimension                   |
| [`VBA209`](#vba209) | analyze | warning     | procedure-local | yes     | Object or array comparison mistake                 |
| [`VBA210`](#vba210) | analyze | warning     | procedure-local | no      | Missing Function or Property Get return assignment |
| [`VBA211`](#vba211) | analyze | error       | procedure-local | yes     | Expanded Excel member mismatch                     |
| [`VBA212`](#vba212) | analyze | warning     | procedure-local | yes     | Non-short-circuit object guard                     |
| [`VBA213`](#vba213) | analyze | warning     | procedure-local | no      | Dictionary iteration value misuse                  |
| [`VBA214`](#vba214) | analyze | warning     | procedure-local | yes     | Leaked On Error Resume Next scope                  |
| [`VBA215`](#vba215) | analyze | warning     | procedure-local | yes     | Omitted stateful Excel call arguments              |
| [`VBA216`](#vba216) | analyze | error       | procedure-local | yes     | Worksheet root mismatch                            |
| [`VBA217`](#vba217) | analyze | warning     | procedure-local | yes     | Unstable last-row boundary                         |
| [`VBA218`](#vba218) | analyze | warning     | interprocedural | yes     | Unhandled Excel API failure contract               |
| [`VBA219`](#vba219) | analyze | warning     | procedure-local | yes     | Unreleased workbook or VBA file handle             |
| [`VBA220`](#vba220) | analyze | warning     | interprocedural | yes     | Event handler re-entry hazard                      |
| [`VBA221`](#vba221) | analyze | warning     | interprocedural | yes     | Application state changed by local helper          |
| [`VBA222`](#vba222) | analyze | warning     | project-wide    | yes     | Public API type safety                             |
| [`VBA223`](#vba223) | analyze | warning     | file-local      | yes     | Likely hardcoded secret                            |
| [`VBA224`](#vba224) | analyze | warning     | procedure-local | yes     | Untrusted data reaches a sensitive API             |
| [`VBA225`](#vba225) | analyze | warning     | interprocedural | yes     | Excel cell access inside loop                      |
| [`VBA226`](#vba226) | analyze | warning     | procedure-local | yes     | Unsafe Range.Value array shape assumption          |
| [`VBA227`](#vba227) | analyze | warning     | interprocedural | yes     | Array lifecycle and dimension safety               |
| [`VBA228`](#vba228) | analyze | error       | interprocedural | yes     | ByRef type mismatch                                |
| [`VBA229`](#vba229) | analyze | error       | procedure-local | yes     | Unresolved local As type name                      |
| [`VBA230`](#vba230) | analyze | warning     | interprocedural | yes     | Dictionary CompareMode changed after insertion     |
| [`VBA231`](#vba231) | analyze | warning     | interprocedural | yes     | Repeated Dictionary loop materialization           |
| [`VBA232`](#vba232) | analyze | warning     | procedure-local | yes     | Inconsistent Dictionary key normalization          |
| [`VBA233`](#vba233) | analyze | warning     | project-wide    | yes     | Undefined late-bound Dictionary constant           |
| [`VBA234`](#vba234) | analyze | warning     | interprocedural | yes     | Collection mutation during iteration               |
| [`VBA235`](#vba235) | analyze | warning     | procedure-local | yes     | Collection index origin confusion                  |
| [`VBA236`](#vba236) | analyze | warning     | procedure-local | yes     | Unsafe command construction                        |
| [`VBA237`](#vba237) | analyze | warning     | interprocedural | yes     | Suppressed error propagation                       |
| [`VBA238`](#vba238) | analyze | warning     | procedure-local | yes     | Loop-invariant Excel object resolution             |
| [`VBA239`](#vba239) | analyze | warning     | procedure-local | yes     | Unsafe SQL construction                            |
| [`VBA240`](#vba240) | analyze | warning     | project-wide    | no      | Risky module-level mutable state                   |
| [`VBA241`](#vba241) | analyze | warning     | procedure-local | yes     | Repeated ReDim Preserve inside loop                |
| [`VBA242`](#vba242) | analyze | information | procedure-local | no      | Expensive full-range operation                     |
| [`VBA243`](#vba243) | analyze | information | procedure-local | no      | Value2 performance opportunity                     |
| [`VBA244`](#vba244) | analyze | information | project-wide    | yes     | Recursive and cyclic procedure dependency          |
| [`VBA245`](#vba245) | analyze | warning     | procedure-local | yes     | Unsafe destructive file and path operation         |
| [`VBA246`](#vba246) | analyze | warning     | procedure-local | yes     | Unsafe HTTP or TLS configuration                   |
| [`VBA247`](#vba247) | analyze | warning     | procedure-local | yes     | Missing or unlimited HTTP timeout                  |
| [`VBA248`](#vba248) | analyze | warning     | procedure-local | no      | Opaque Boolean control arguments                   |
| [`VBA249`](#vba249) | analyze | error       | procedure-local | yes     | Deterministic runtime error                        |

## VB001

**Missing Option Explicit.** The module does not declare Option Explicit.

| Property                    | Value                     |
| --------------------------- | ------------------------- |
| Family                      | `lint`                    |
| Category                    | `correctness`             |
| Evidence class              | `policy`                  |
| Compile-equivalent          | no                        |
| Default severity            | `warning`                 |
| Supported severities        | `warning`                 |
| Surfaces                    | `lint`, `lsp`             |
| Scope                       | `file-local`              |
| Precision                   | `high`                    |
| Enabled by default          | yes                       |
| Configuration               | `require_option_explicit` |
| Inline suppression          | yes                       |
| Blocks source preflight     | no                        |
| Real-time editor diagnostic | yes                       |
| Fix available               | no                        |

## VB002

**Select usage.** Select makes behavior depend on Excel selection state.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `reliability`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | `forbid_select`   |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB003

**Activate usage.** Activate makes behavior depend on Excel active-object state.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `reliability`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | `forbid_activate` |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB004

**Broad On Error Resume Next.** On Error Resume Next is used without a narrow recovery scope.

| Property                    | Value                         |
| --------------------------- | ----------------------------- |
| Family                      | `lint`                        |
| Category                    | `reliability`                 |
| Evidence class              | `inference`                   |
| Compile-equivalent          | no                            |
| Default severity            | `warning`                     |
| Supported severities        | `warning`                     |
| Surfaces                    | `lint`, `lsp`                 |
| Scope                       | `procedure-local`             |
| Precision                   | `high`                        |
| Enabled by default          | yes                           |
| Configuration               | `forbid_on_error_resume_next` |
| Inline suppression          | yes                           |
| Blocks source preflight     | no                            |
| Real-time editor diagnostic | yes                           |
| Fix available               | no                            |

## VB005

**Implicit Variant.** A declaration omits an explicit As type.

| Property                    | Value                     |
| --------------------------- | ------------------------- |
| Family                      | `lint`                    |
| Category                    | `maintainability`         |
| Evidence class              | `maintainability`         |
| Compile-equivalent          | no                        |
| Default severity            | `warning`                 |
| Supported severities        | `warning`                 |
| Surfaces                    | `lint`, `lsp`             |
| Scope                       | `file-local`              |
| Precision                   | `high`                    |
| Enabled by default          | yes                       |
| Configuration               | `detect_implicit_variant` |
| Inline suppression          | yes                       |
| Blocks source preflight     | no                        |
| Real-time editor diagnostic | yes                       |
| Fix available               | no                        |

## VB006

**Public module field.** A module-level Public variable exposes mutable global state.

| Property                    | Value                         |
| --------------------------- | ----------------------------- |
| Family                      | `lint`                        |
| Category                    | `maintainability`             |
| Evidence class              | `maintainability`             |
| Compile-equivalent          | no                            |
| Default severity            | `warning`                     |
| Supported severities        | `warning`                     |
| Surfaces                    | `lint`, `lsp`                 |
| Scope                       | `file-local`                  |
| Precision                   | `high`                        |
| Enabled by default          | yes                           |
| Configuration               | `forbid_public_module_fields` |
| Inline suppression          | yes                           |
| Blocks source preflight     | no                            |
| Real-time editor diagnostic | yes                           |
| Fix available               | no                            |

## VB007

**Automation-hostile GUI boundary.** The source invokes an interactive boundary that can block automation.

| Property                    | Value                      |
| --------------------------- | -------------------------- |
| Family                      | `lint`                     |
| Category                    | `reliability`              |
| Evidence class              | `inference`                |
| Compile-equivalent          | no                         |
| Default severity            | `warning`                  |
| Supported severities        | `warning`                  |
| Surfaces                    | `lint`                     |
| Scope                       | `file-local`               |
| Precision                   | `high`                     |
| Enabled by default          | yes                        |
| Configuration               | `forbid_interactive_input` |
| Inline suppression          | yes                        |
| Blocks source preflight     | no                         |
| Real-time editor diagnostic | no                         |
| Fix available               | no                         |

## VB008

**Typographic quote.** A typographic quote can trigger an Excel VBE compile dialog.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB009

**C-style quote escape.** A C-style escaped quote is invalid VBA string syntax.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB010

**Unterminated procedure.** A Sub, Function, or Property procedure has no matching terminator.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB011

**Unexpected procedure terminator.** An End procedure statement has no matching opener.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB012

**Mismatched procedure terminator.** A procedure is closed with an End statement kind that VBA's VBE rejects.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB013

**Invalid line continuation.** A line-continuation underscore is not preceded by whitespace.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB014

**Parser recovery.** The parser recovered from an error, missing node, or unmatched source structure.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `medium`             |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB015

**Continuation limit exceeded.** A VBA logical line exceeds the supported continuation-line limit.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB018

**Scope shadowing.** A declaration shadows another visible declaration or procedure.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `maintainability`        |
| Evidence class              | `maintainability`        |
| Compile-equivalent          | no                       |
| Default severity            | `warning`                |
| Supported severities        | `warning`                |
| Surfaces                    | `lint`                   |
| Scope                       | `project-wide`           |
| Precision                   | `medium`                 |
| Enabled by default          | no                       |
| Configuration               | `detect_scope_shadowing` |
| Inline suppression          | yes                      |
| Blocks source preflight     | no                       |
| Real-time editor diagnostic | no                       |
| Fix available               | no                       |

## VB019

**Mixed declarator typing.** A multi-name declaration explicitly types only some declarators.

| Property                    | Value                                |
| --------------------------- | ------------------------------------ |
| Family                      | `lint`                               |
| Category                    | `correctness`                        |
| Evidence class              | `inference`                          |
| Compile-equivalent          | no                                   |
| Default severity            | `warning`                            |
| Supported severities        | `warning`                            |
| Surfaces                    | `lint`, `lsp`                        |
| Scope                       | `file-local`                         |
| Precision                   | `high`                               |
| Enabled by default          | yes                                  |
| Configuration               | `detect_multiple_declarator_clarity` |
| Inline suppression          | yes                                  |
| Blocks source preflight     | no                                   |
| Real-time editor diagnostic | yes                                  |
| Fix available               | no                                   |

## VB020

**Unused local variable.** A procedure-local variable is declared but never referenced.

| Property                    | Value                           |
| --------------------------- | ------------------------------- |
| Family                      | `lint`                          |
| Category                    | `maintainability`               |
| Evidence class              | `maintainability`               |
| Compile-equivalent          | no                              |
| Default severity            | `warning`                       |
| Supported severities        | `warning`                       |
| Surfaces                    | `lint`, `lsp`                   |
| Scope                       | `procedure-local`               |
| Precision                   | `high`                          |
| Enabled by default          | yes                             |
| Configuration               | `detect_unused_local_variables` |
| Inline suppression          | yes                             |
| Blocks source preflight     | no                              |
| Real-time editor diagnostic | yes                             |
| Fix available               | no                              |

## VB021

**Unused private procedure.** A private procedure is unreachable from known project roots.

| Property                    | Value                              |
| --------------------------- | ---------------------------------- |
| Family                      | `lint`                             |
| Category                    | `maintainability`                  |
| Evidence class              | `maintainability`                  |
| Compile-equivalent          | no                                 |
| Default severity            | `warning`                          |
| Supported severities        | `warning`                          |
| Surfaces                    | `lint`                             |
| Scope                       | `project-wide`                     |
| Precision                   | `medium`                           |
| Enabled by default          | no                                 |
| Configuration               | `detect_unused_private_procedures` |
| Inline suppression          | yes                                |
| Blocks source preflight     | no                                 |
| Real-time editor diagnostic | no                                 |
| Fix available               | no                                 |

## VB022

**Confusing call syntax.** A parenthesized procedure call uses ambiguous VBA call syntax.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `lint`                         |
| Category                    | `maintainability`              |
| Evidence class              | `maintainability`              |
| Compile-equivalent          | no                             |
| Default severity            | `warning`                      |
| Supported severities        | `warning`                      |
| Surfaces                    | `lint`, `lsp`                  |
| Scope                       | `procedure-local`              |
| Precision                   | `high`                         |
| Enabled by default          | yes                            |
| Configuration               | `detect_confusing_call_syntax` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | yes                            |
| Fix available               | no                             |

## VB023

**Invalid For Each control type.** A For Each control variable is undeclared or incompatible.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `lint`                         |
| Category                    | `correctness`                  |
| Evidence class              | `inference`                    |
| Compile-equivalent          | no                             |
| Default severity            | `warning`                      |
| Supported severities        | `warning`                      |
| Surfaces                    | `lint`, `lsp`                  |
| Scope                       | `procedure-local`              |
| Precision                   | `high`                         |
| Enabled by default          | yes                            |
| Configuration               | `detect_for_each_control_type` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | yes                            |
| Fix available               | no                             |

## VB026

**Dangerous Resume.** Resume is used outside a recognizable error-handler context.

| Property                    | Value                     |
| --------------------------- | ------------------------- |
| Family                      | `lint`                    |
| Category                    | `reliability`             |
| Evidence class              | `inference`               |
| Compile-equivalent          | no                        |
| Default severity            | `warning`                 |
| Supported severities        | `warning`                 |
| Surfaces                    | `lint`, `lsp`             |
| Scope                       | `procedure-local`         |
| Precision                   | `medium`                  |
| Enabled by default          | yes                       |
| Configuration               | `detect_dangerous_resume` |
| Inline suppression          | yes                       |
| Blocks source preflight     | no                        |
| Real-time editor diagnostic | yes                       |
| Fix available               | no                        |

## VB027

**Ambiguous nested With member.** An implicit Excel member inside nested With blocks has an ambiguous target.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `lint`                         |
| Category                    | `reliability`                  |
| Evidence class              | `inference`                    |
| Compile-equivalent          | no                             |
| Default severity            | `warning`                      |
| Supported severities        | `warning`                      |
| Surfaces                    | `lint`, `lsp`                  |
| Scope                       | `procedure-local`              |
| Precision                   | `medium`                       |
| Enabled by default          | no                             |
| Configuration               | `detect_nested_with_ambiguity` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | yes                            |
| Fix available               | no                             |

## VB028

**Bare dialog call with XlflowUI.** A bare dialog call can bind to an incompatible XlflowUI member.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`               |
| Scope                       | `project-wide`       |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | no                   |
| Fix available               | no                   |

## VB029

**Undeclared variable.** Option Explicit is present but a referenced assignment or loop variable is undeclared in the project-visible namespace.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `project-wide`       |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB030

**Argument mismatch.** A procedure call argument does not match the resolved signature.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB031

**Missing module name attribute.** A standard module lacks its exported Attribute VB_Name declaration.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB032

**Repeated Debug.Print shorthand.** Repeated question-mark shorthand is invalid VBA syntax.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB033

**Unknown member.** A member is not present on the resolved receiver type.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB034

**Read-only property assignment.** An assignment targets a read-only property.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB035

**Write-only property read.** An expression reads a write-only property.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB036

**Set required.** An object assignment omits the Set keyword.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB037

**Set not allowed.** A value assignment incorrectly uses the Set keyword.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `type-safety`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB038

**Incompatible assignment.** The assigned expression is incompatible with the resolved target type.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB039

**Method has no return value.** A method without a return value is used as an expression.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB040

**Unknown documented parameter.** A documentation comment names a parameter that the procedure does not declare.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `documentation`   |
| Evidence class              | `maintainability` |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `file-local`      |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB041

**Duplicate documented parameter.** A documentation comment lists the same parameter more than once.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `documentation`   |
| Evidence class              | `maintainability` |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `file-local`      |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB042

**Returns documentation on Sub.** A Sub procedure documentation comment contains a Returns section.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `documentation`   |
| Evidence class              | `maintainability` |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `file-local`      |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB043

**Orphan documentation comment.** A documentation comment is not associated with a declaration.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `documentation`   |
| Evidence class              | `maintainability` |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `file-local`      |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VB044

**Procedure-name constant mismatch.** A configured procedure-name string constant differs from its enclosing procedure name.

| Property                    | Value                     |
| --------------------------- | ------------------------- |
| Family                      | `lint`                    |
| Category                    | `maintainability`         |
| Evidence class              | `maintainability`         |
| Compile-equivalent          | no                        |
| Default severity            | `warning`                 |
| Supported severities        | `warning`                 |
| Surfaces                    | `lint`, `lsp`             |
| Scope                       | `procedure-local`         |
| Precision                   | `high`                    |
| Enabled by default          | no                        |
| Configuration               | `procedure_name_constant` |
| Inline suppression          | yes                       |
| Blocks source preflight     | no                        |
| Real-time editor diagnostic | yes                       |
| Fix available               | yes                       |

## VB045

**Deterministic argument binding error.** A resolved procedure call has a deterministic argument-count or named-argument compile error.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `procedure-local`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB046

**Duplicate declaration.** A VBA declaration repeats a name that is already declared in the same scope.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB047

**Invalid declaration placement.** A VBA declaration appears in a source position where that declaration kind is not permitted.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB048

**Invalid procedure parameter declaration.** A procedure parameter declaration violates a VBA compile-time parameter contract.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB049

**Inconsistent Property accessor contract.** Property Get, Let, and Set accessors in the same module do not form a compatible VBA accessor contract.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB050

**Invalid module declaration context.** A VBA declaration or member uses a form that is not permitted by the containing module kind.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `file-local`         |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB051

**Invalid Me context.** The VBA Me keyword is used outside an object module.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB052

**Invalid procedure call target.** A call targets a project-local symbol that is missing or known not to be callable.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `interprocedural`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB053

**Ambiguous Enum member.** A bare Enum member reference has multiple visible project or type-library candidates and no unique lexical winner.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `interprocedural`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB054

**Undeclared RaiseEvent target.** A RaiseEvent statement names an event that is not declared in the same object module.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `interprocedural`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB055

**Duplicate procedure label.** A procedure defines the same label more than once.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `procedure-local`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB056

**Undefined procedure label.** A procedure control-flow transfer targets a label that is not defined in the procedure.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `procedure-local`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB057

**Mismatched Next control variable.** A Next variable does not match the active For or For Each control variable.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `procedure-local`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB058

**Invalid Exit statement.** An Exit statement is incompatible with the enclosing procedure or active loop.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `procedure-local`        |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB059

**Invalid call syntax.** A VBA call uses a parenthesis or explicit Call form that the compiler rejects, or an expression invokes a Function without its required argument parentheses.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB060

**Assignment to constant.** An assignment targets a Const declaration.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `project-wide`           |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB061

**Invalid constant array bounds.** A fixed array declaration has a constant lower bound greater than its upper bound.

| Property                    | Value                    |
| --------------------------- | ------------------------ |
| Family                      | `lint`                   |
| Category                    | `correctness`            |
| Evidence class              | `compile-equivalent`     |
| Compile-equivalent          | yes                      |
| Default severity            | `error`                  |
| Supported severities        | `error`                  |
| Surfaces                    | `lint`, `lsp`, `analyze` |
| Scope                       | `file-local`             |
| Precision                   | `high`                   |
| Enabled by default          | yes                      |
| Configuration               | not configurable         |
| Inline suppression          | no                       |
| Blocks source preflight     | yes                      |
| Real-time editor diagnostic | yes                      |
| Fix available               | no                       |

## VB062

**Invalid conditional branch syntax.** A conditional branch statement uses a form that VBA rejects.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB063

**Invalid Select/Case branch syntax.** A Select Case branch uses a form or ordering that VBA rejects.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB064

**Invalid Open mode syntax.** An Open statement uses a file mode form that VBA rejects.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB065

**Invalid TypeOf syntax.** A TypeOf expression uses a form that VBA rejects.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `lint`               |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `lint`, `lsp`        |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VB066

**Procedure terminator style mismatch.** A VBE-accepted procedure uses a noncanonical End statement kind.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `lint`            |
| Category                    | `maintainability` |
| Evidence class              | `maintainability` |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `lint`, `lsp`     |
| Scope                       | `file-local`      |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | yes               |
| Fix available               | no                |

## VBA101

**Object assignment missing Set.** An object variable assignment likely omits Set.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `analyze`         |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `analyze`         |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | no                |
| Fix available               | no                |

## VBA102

**Object-returning call assignment missing Set.** An object-returning function assignment likely omits Set.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `analyze`         |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `analyze`         |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | no                |
| Fix available               | no                |

## VBA103

**Object function return missing Set.** An object-returning function body likely assigns its return without Set.

| Property                    | Value             |
| --------------------------- | ----------------- |
| Family                      | `analyze`         |
| Category                    | `type-safety`     |
| Evidence class              | `inference`       |
| Compile-equivalent          | no                |
| Default severity            | `warning`         |
| Supported severities        | `warning`         |
| Surfaces                    | `analyze`         |
| Scope                       | `procedure-local` |
| Precision                   | `high`            |
| Enabled by default          | yes               |
| Configuration               | not configurable  |
| Inline suppression          | yes               |
| Blocks source preflight     | no                |
| Real-time editor diagnostic | no                |
| Fix available               | no                |

## VBA104

**Excel object member mismatch.** A known Excel object does not provide the referenced member.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `analyze`            |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `analyze`            |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | no                   |
| Fix available               | no                   |

## VBA105

**Removed XlflowLog helper.** The source calls the removed XlflowLog trace helper.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `analyze`            |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `analyze`            |
| Scope                       | `project-wide`       |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | no                   |
| Fix available               | no                   |

## VBA106

**Removed XlflowSetTraceFile helper.** The source calls the removed XlflowSetTraceFile trace helper.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `analyze`            |
| Category                    | `correctness`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `analyze`            |
| Scope                       | `project-wide`       |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | no                   |
| Fix available               | no                   |

## VBA201

**Unchecked Range.Find result.** A Range.Find result is dereferenced before a Nothing check.

| Property                    | Value                             |
| --------------------------- | --------------------------------- |
| Family                      | `analyze`                         |
| Category                    | `reliability`                     |
| Evidence class              | `inference`                       |
| Compile-equivalent          | no                                |
| Default severity            | `warning`                         |
| Supported severities        | `warning`                         |
| Surfaces                    | `analyze`, `lsp`                  |
| Scope                       | `procedure-local`                 |
| Precision                   | `high`                            |
| Enabled by default          | yes                               |
| Configuration               | `detect_range_find_nothing_check` |
| Inline suppression          | yes                               |
| Blocks source preflight     | no                                |
| Real-time editor diagnostic | yes                               |
| Fix available               | no                                |

## VBA202

**Object use before Set.** An object variable may be dereferenced before a definitely non-Nothing value is proven on every reachable path.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `analyze`                      |
| Category                    | `reliability`                  |
| Evidence class              | `inference`                    |
| Compile-equivalent          | no                             |
| Default severity            | `warning`                      |
| Supported severities        | `warning`                      |
| Surfaces                    | `analyze`                      |
| Scope                       | `interprocedural`              |
| Precision                   | `medium`                       |
| Enabled by default          | yes                            |
| Configuration               | `detect_object_use_before_set` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | no                             |
| Fix available               | no                             |

## VBA203

**Application state not restored.** An Application state change can reach an exit without restoring its previous value.

| Property                    | Value                              |
| --------------------------- | ---------------------------------- |
| Family                      | `analyze`                          |
| Category                    | `reliability`                      |
| Evidence class              | `inference`                        |
| Compile-equivalent          | no                                 |
| Default severity            | `warning`                          |
| Supported severities        | `warning`                          |
| Surfaces                    | `analyze`                          |
| Scope                       | `interprocedural`                  |
| Precision                   | `medium`                           |
| Enabled by default          | yes                                |
| Configuration               | `detect_application_state_restore` |
| Inline suppression          | yes                                |
| Blocks source preflight     | no                                 |
| Real-time editor diagnostic | no                                 |
| Fix available               | no                                 |

## VBA204

**Error-handler fallthrough.** Normal execution can fall through into an error-handler label.

| Property                    | Value                              |
| --------------------------- | ---------------------------------- |
| Family                      | `analyze`                          |
| Category                    | `reliability`                      |
| Evidence class              | `inference`                        |
| Compile-equivalent          | no                                 |
| Default severity            | `warning`                          |
| Supported severities        | `warning`                          |
| Surfaces                    | `analyze`, `lsp`                   |
| Scope                       | `procedure-local`                  |
| Precision                   | `high`                             |
| Enabled by default          | yes                                |
| Configuration               | `detect_error_handler_fallthrough` |
| Inline suppression          | yes                                |
| Blocks source preflight     | no                                 |
| Real-time editor diagnostic | yes                                |
| Fix available               | no                                 |

## VBA205

**Ambiguous Excel object scope.** Active UI objects, unqualified worksheet members (Range, Cells, Rows, and Columns), unqualified sheet collections, positional workbook/window access, uncaptured Workbooks.Open calls, and add-in ThisWorkbook references can target an unintended workbook or worksheet.

| Property                    | Value                              |
| --------------------------- | ---------------------------------- |
| Family                      | `analyze`                          |
| Category                    | `reliability`                      |
| Evidence class              | `inference`                        |
| Compile-equivalent          | no                                 |
| Default severity            | `warning`                          |
| Supported severities        | `warning`                          |
| Surfaces                    | `analyze`                          |
| Scope                       | `procedure-local`                  |
| Precision                   | `high`                             |
| Enabled by default          | yes                                |
| Configuration               | `forbid_unqualified_excel_objects` |
| Inline suppression          | yes                                |
| Blocks source preflight     | no                                 |
| Real-time editor diagnostic | no                                 |
| Fix available               | no                                 |

## VBA206

**Unsafe ByRef argument.** A resolved project-local ByRef call receives an incompatible type, a temporary value, or an indirect property, member, or array expression whose mutation may be lost or surprising.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `runtime-safety`                 |
| Evidence class              | `runtime-safety`                 |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`                        |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `interprocedural`                |
| Precision                   | `high`                           |
| Enabled by default          | yes                              |
| Configuration               | `detect_byref_argument_mismatch` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA207

**Unguarded keyed access.** Dictionary or Collection item access has no obvious existence guard.

| Property                    | Value                                |
| --------------------------- | ------------------------------------ |
| Family                      | `analyze`                            |
| Category                    | `reliability`                        |
| Evidence class              | `inference`                          |
| Compile-equivalent          | no                                   |
| Default severity            | `warning`                            |
| Supported severities        | `warning`, `information`             |
| Surfaces                    | `analyze`                            |
| Scope                       | `procedure-local`                    |
| Precision                   | `medium`                             |
| Enabled by default          | no                                   |
| Configuration               | `detect_dictionary_collection_guard` |
| Inline suppression          | yes                                  |
| Blocks source preflight     | no                                   |
| Real-time editor diagnostic | no                                   |
| Fix available               | no                                   |

## VBA208

**Invalid ReDim Preserve dimension.** ReDim Preserve is used on a multi-dimensional array.

| Property                    | Value                             |
| --------------------------- | --------------------------------- |
| Family                      | `analyze`                         |
| Category                    | `correctness`                     |
| Evidence class              | `inference`                       |
| Compile-equivalent          | no                                |
| Default severity            | `warning`                         |
| Supported severities        | `warning`                         |
| Surfaces                    | `analyze`, `lsp`                  |
| Scope                       | `procedure-local`                 |
| Precision                   | `high`                            |
| Enabled by default          | yes                               |
| Configuration               | `detect_redim_preserve_dimension` |
| Inline suppression          | yes                               |
| Blocks source preflight     | no                                |
| Real-time editor diagnostic | yes                               |
| Fix available               | no                                |

## VBA209

**Object or array comparison mistake.** An object or array is compared with scalar equality semantics.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `type-safety`                    |
| Evidence class              | `inference`                      |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`                        |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `high`                           |
| Enabled by default          | yes                              |
| Configuration               | `detect_object_array_comparison` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA210

**Missing Function or Property Get return assignment.** A Function or Property Get may reach normal exit without a valid return assignment on every path.

| Property                    | Value                         |
| --------------------------- | ----------------------------- |
| Family                      | `analyze`                     |
| Category                    | `correctness`                 |
| Evidence class              | `inference`                   |
| Compile-equivalent          | no                            |
| Default severity            | `warning`                     |
| Supported severities        | `warning`                     |
| Surfaces                    | `analyze`                     |
| Scope                       | `procedure-local`             |
| Precision                   | `medium`                      |
| Enabled by default          | no                            |
| Configuration               | `detect_function_return_path` |
| Inline suppression          | yes                           |
| Blocks source preflight     | no                            |
| Real-time editor diagnostic | no                            |
| Fix available               | no                            |

## VBA211

**Expanded Excel member mismatch.** Expanded Excel type metadata proves an object/member mismatch.

| Property                    | Value                                 |
| --------------------------- | ------------------------------------- |
| Family                      | `analyze`                             |
| Category                    | `correctness`                         |
| Evidence class              | `compile-equivalent`                  |
| Compile-equivalent          | yes                                   |
| Default severity            | `error`                               |
| Supported severities        | `error`                               |
| Surfaces                    | `analyze`                             |
| Scope                       | `procedure-local`                     |
| Precision                   | `high`                                |
| Enabled by default          | yes                                   |
| Configuration               | `detect_excel_object_member_mismatch` |
| Inline suppression          | no                                    |
| Blocks source preflight     | yes                                   |
| Real-time editor diagnostic | no                                    |
| Fix available               | no                                    |

## VBA212

**Non-short-circuit object guard.** A Nothing or IsArray guard and a matching object, index, bound, or resolved side-effecting getter access share an eagerly evaluated And, Or, IIf, Choose, or Switch expression.

| Property                    | Value                                   |
| --------------------------- | --------------------------------------- |
| Family                      | `analyze`                               |
| Category                    | `reliability`                           |
| Evidence class              | `inference`                             |
| Compile-equivalent          | no                                      |
| Default severity            | `warning`                               |
| Supported severities        | `warning`                               |
| Surfaces                    | `analyze`, `lsp`                        |
| Scope                       | `procedure-local`                       |
| Precision                   | `high`                                  |
| Enabled by default          | yes                                     |
| Configuration               | `detect_non_short_circuit_object_guard` |
| Inline suppression          | yes                                     |
| Blocks source preflight     | no                                      |
| Real-time editor diagnostic | yes                                     |
| Fix available               | no                                      |

## VBA213

**Dictionary iteration value misuse.** A Dictionary key yielded by direct iteration is used as a value or object.

| Property                    | Value                                     |
| --------------------------- | ----------------------------------------- |
| Family                      | `analyze`                                 |
| Category                    | `type-safety`                             |
| Evidence class              | `inference`                               |
| Compile-equivalent          | no                                        |
| Default severity            | `warning`                                 |
| Supported severities        | `warning`                                 |
| Surfaces                    | `analyze`, `lsp`                          |
| Scope                       | `procedure-local`                         |
| Precision                   | `medium`                                  |
| Enabled by default          | no                                        |
| Configuration               | `detect_dictionary_iteration_value_usage` |
| Inline suppression          | yes                                       |
| Blocks source preflight     | no                                        |
| Real-time editor diagnostic | yes                                       |
| Fix available               | no                                        |

## VBA214

**Leaked On Error Resume Next scope.** On Error Resume Next remains active across an unsafe procedure scope.

| Property                    | Value                                       |
| --------------------------- | ------------------------------------------- |
| Family                      | `analyze`                                   |
| Category                    | `reliability`                               |
| Evidence class              | `inference`                                 |
| Compile-equivalent          | no                                          |
| Default severity            | `warning`                                   |
| Supported severities        | `warning`                                   |
| Surfaces                    | `analyze`                                   |
| Scope                       | `procedure-local`                           |
| Precision                   | `medium`                                    |
| Enabled by default          | yes                                         |
| Configuration               | `detect_leaked_on_error_resume_next_scopes` |
| Inline suppression          | yes                                         |
| Blocks source preflight     | no                                          |
| Real-time editor diagnostic | no                                          |
| Fix available               | no                                          |

## VBA215

**Omitted stateful Excel call arguments.** A Range.Find or Range.Replace call omits settings that Excel saves and can reuse from prior UI operations or macro calls.

| Property                    | Value                                  |
| --------------------------- | -------------------------------------- |
| Family                      | `analyze`                              |
| Category                    | `reliability`                          |
| Evidence class              | `inference`                            |
| Compile-equivalent          | no                                     |
| Default severity            | `warning`                              |
| Supported severities        | `warning`                              |
| Surfaces                    | `analyze`, `lsp`                       |
| Scope                       | `procedure-local`                      |
| Precision                   | `high`                                 |
| Enabled by default          | yes                                    |
| Configuration               | `detect_stateful_excel_call_arguments` |
| Inline suppression          | yes                                    |
| Blocks source preflight     | no                                     |
| Real-time editor diagnostic | yes                                    |
| Fix available               | no                                     |

## VBA216

**Worksheet root mismatch.** A range expression combines explicit objects rooted in different worksheets.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `reliability`                    |
| Evidence class              | `compile-equivalent`             |
| Compile-equivalent          | yes                              |
| Default severity            | `error`                          |
| Supported severities        | `error`                          |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `high`                           |
| Enabled by default          | yes                              |
| Configuration               | `detect_worksheet_root_mismatch` |
| Inline suppression          | no                               |
| Blocks source preflight     | yes                              |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA217

**Unstable last-row boundary.** A last-row calculation depends on an implicit worksheet root or a boundary pattern that can produce unstable results.

| Property                    | Value                               |
| --------------------------- | ----------------------------------- |
| Family                      | `analyze`                           |
| Category                    | `reliability`                       |
| Evidence class              | `inference`                         |
| Compile-equivalent          | no                                  |
| Default severity            | `warning`                           |
| Supported severities        | `warning`                           |
| Surfaces                    | `analyze`, `lsp`                    |
| Scope                       | `procedure-local`                   |
| Precision                   | `medium`                            |
| Enabled by default          | yes                                 |
| Configuration               | `detect_unstable_last_row_patterns` |
| Inline suppression          | yes                                 |
| Blocks source preflight     | no                                  |
| Real-time editor diagnostic | yes                                 |
| Fix available               | no                                  |

## VBA218

**Unhandled Excel API failure contract.** An Excel API call can raise a runtime error or return Variant/Error when no result exists, but its failure contract is not handled.

| Property                    | Value                                |
| --------------------------- | ------------------------------------ |
| Family                      | `analyze`                            |
| Category                    | `reliability`                        |
| Evidence class              | `inference`                          |
| Compile-equivalent          | no                                   |
| Default severity            | `warning`                            |
| Supported severities        | `warning`                            |
| Surfaces                    | `analyze`, `lsp`                     |
| Scope                       | `interprocedural`                    |
| Precision                   | `high`                               |
| Enabled by default          | yes                                  |
| Configuration               | `detect_excel_api_failure_contracts` |
| Inline suppression          | yes                                  |
| Blocks source preflight     | no                                   |
| Real-time editor diagnostic | yes                                  |
| Fix available               | no                                   |

## VBA219

**Unreleased workbook or VBA file handle.** A procedure-local Workbook or VBA file handle acquired by an explicit Open call can reach an exit without a matching Close.

| Property                    | Value                   |
| --------------------------- | ----------------------- |
| Family                      | `analyze`               |
| Category                    | `reliability`           |
| Evidence class              | `inference`             |
| Compile-equivalent          | no                      |
| Default severity            | `warning`               |
| Supported severities        | `warning`               |
| Surfaces                    | `analyze`, `lsp`        |
| Scope                       | `procedure-local`       |
| Precision                   | `high`                  |
| Enabled by default          | yes                     |
| Configuration               | `detect_resource_leaks` |
| Inline suppression          | yes                     |
| Blocks source preflight     | no                      |
| Real-time editor diagnostic | yes                     |
| Fix available               | no                      |

## VBA220

**Event handler re-entry hazard.** An Excel or UserForm event handler can trigger itself or a related event without a proven safe event guard.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `analyze`                      |
| Category                    | `reliability`                  |
| Evidence class              | `inference`                    |
| Compile-equivalent          | no                             |
| Default severity            | `warning`                      |
| Supported severities        | `warning`                      |
| Surfaces                    | `analyze`                      |
| Scope                       | `interprocedural`              |
| Precision                   | `medium`                       |
| Enabled by default          | yes                            |
| Configuration               | `detect_event_handler_reentry` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | no                             |
| Fix available               | no                             |

## VBA221

**Application state changed by local helper.** A direct project-local call can leave an Application property changed because the callee does not restore it on every exit.

| Property                    | Value                                   |
| --------------------------- | --------------------------------------- |
| Family                      | `analyze`                               |
| Category                    | `reliability`                           |
| Evidence class              | `inference`                             |
| Compile-equivalent          | no                                      |
| Default severity            | `warning`                               |
| Supported severities        | `warning`                               |
| Surfaces                    | `analyze`                               |
| Scope                       | `interprocedural`                       |
| Precision                   | `high`                                  |
| Enabled by default          | yes                                     |
| Configuration               | `detect_application_state_call_effects` |
| Inline suppression          | yes                                     |
| Blocks source preflight     | no                                      |
| Real-time editor diagnostic | no                                      |
| Fix available               | no                                      |

## VBA222

**Public API type safety.** A public API declaration or exposed type uses an inaccessible, ambiguous, or unresolved type.

| Property                    | Value                           |
| --------------------------- | ------------------------------- |
| Family                      | `analyze`                       |
| Category                    | `type-safety`                   |
| Evidence class              | `inference`                     |
| Compile-equivalent          | no                              |
| Default severity            | `warning`                       |
| Supported severities        | `warning`                       |
| Surfaces                    | `analyze`                       |
| Scope                       | `project-wide`                  |
| Precision                   | `medium`                        |
| Enabled by default          | yes                             |
| Configuration               | `detect_public_api_type_safety` |
| Inline suppression          | yes                             |
| Blocks source preflight     | no                              |
| Real-time editor diagnostic | no                              |
| Fix available               | no                              |

## VBA223

**Likely hardcoded secret.** A credential-like value appears directly in VBA source.

| Property                    | Value                      |
| --------------------------- | -------------------------- |
| Family                      | `analyze`                  |
| Category                    | `security`                 |
| Evidence class              | `policy`                   |
| Compile-equivalent          | no                         |
| Default severity            | `warning`                  |
| Supported severities        | `warning`                  |
| Surfaces                    | `analyze`, `lsp`           |
| Scope                       | `file-local`               |
| Precision                   | `high`                     |
| Enabled by default          | yes                        |
| Configuration               | `detect_hardcoded_secrets` |
| Inline suppression          | yes                        |
| Blocks source preflight     | no                         |
| Real-time editor diagnostic | yes                        |
| Fix available               | no                         |

## VBA224

**Untrusted data reaches a sensitive API.** Conservative procedure-local analysis found potentially untrusted data flowing into a sensitive API.

| Property                    | Value                        |
| --------------------------- | ---------------------------- |
| Family                      | `analyze`                    |
| Category                    | `security`                   |
| Evidence class              | `policy`                     |
| Compile-equivalent          | no                           |
| Default severity            | `warning`                    |
| Supported severities        | `warning`                    |
| Surfaces                    | `analyze`, `lsp`             |
| Scope                       | `procedure-local`            |
| Precision                   | `medium`                     |
| Enabled by default          | yes                          |
| Configuration               | `detect_untrusted_data_flow` |
| Inline suppression          | yes                          |
| Blocks source preflight     | no                           |
| Real-time editor diagnostic | yes                          |
| Fix available               | no                           |

## VBA225

**Excel cell access inside loop.** A loop repeatedly accesses Excel cells or related object-model members, causing avoidable COM round trips.

| Property                    | Value                               |
| --------------------------- | ----------------------------------- |
| Family                      | `analyze`                           |
| Category                    | `performance`                       |
| Evidence class              | `maintainability`                   |
| Compile-equivalent          | no                                  |
| Default severity            | `warning`                           |
| Supported severities        | `warning`                           |
| Surfaces                    | `analyze`, `lsp`                    |
| Scope                       | `interprocedural`                   |
| Precision                   | `medium`                            |
| Enabled by default          | yes                                 |
| Configuration               | `detect_excel_cell_access_in_loops` |
| Inline suppression          | yes                                 |
| Blocks source preflight     | no                                  |
| Real-time editor diagnostic | yes                                 |
| Fix available               | no                                  |

## VBA226

**Unsafe Range.Value array shape assumption.** A Range.Value or Range.Value2 result is consumed as a scalar or one-dimensional array, or assigned to an incompatible range shape.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `runtime-safety`                 |
| Evidence class              | `runtime-safety`                 |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`                        |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `medium`                         |
| Enabled by default          | yes                              |
| Configuration               | `detect_range_value_array_shape` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA227

**Array lifecycle and dimension safety.** An array may be unallocated, accessed with an invalid dimension or bound, or used without a provable array value; real-time array-return summaries are limited to the active document.

| Property                    | Value                           |
| --------------------------- | ------------------------------- |
| Family                      | `analyze`                       |
| Category                    | `runtime-safety`                |
| Evidence class              | `runtime-safety`                |
| Compile-equivalent          | no                              |
| Default severity            | `warning`                       |
| Supported severities        | `warning`                       |
| Surfaces                    | `analyze`, `lsp`                |
| Scope                       | `interprocedural`               |
| Precision                   | `medium`                        |
| Enabled by default          | yes                             |
| Configuration               | `detect_array_lifecycle_safety` |
| Inline suppression          | yes                             |
| Blocks source preflight     | no                              |
| Real-time editor diagnostic | yes                             |
| Fix available               | no                              |

## VBA228

**ByRef type mismatch.** A resolved project-local ByRef call passes a statically incompatible argument type that the VBE rejects at compile time.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `analyze`            |
| Category                    | `type-safety`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `analyze`, `lsp`     |
| Scope                       | `interprocedural`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VBA229

**Unresolved local As type name.** A procedure-local declaration uses a VBA type name that cannot be resolved from the project or production type database.

| Property                    | Value                |
| --------------------------- | -------------------- |
| Family                      | `analyze`            |
| Category                    | `type-safety`        |
| Evidence class              | `compile-equivalent` |
| Compile-equivalent          | yes                  |
| Default severity            | `error`              |
| Supported severities        | `error`              |
| Surfaces                    | `analyze`, `lsp`     |
| Scope                       | `procedure-local`    |
| Precision                   | `high`               |
| Enabled by default          | yes                  |
| Configuration               | not configurable     |
| Inline suppression          | no                   |
| Blocks source preflight     | yes                  |
| Real-time editor diagnostic | yes                  |
| Fix available               | no                   |

## VBA230

**Dictionary CompareMode changed after insertion.** A Dictionary CompareMode is changed after entries have been added to the same dictionary instance.

| Property                    | Value                                  |
| --------------------------- | -------------------------------------- |
| Family                      | `analyze`                              |
| Category                    | `runtime-safety`                       |
| Evidence class              | `runtime-safety`                       |
| Compile-equivalent          | no                                     |
| Default severity            | `warning`                              |
| Supported severities        | `warning`                              |
| Surfaces                    | `analyze`, `lsp`                       |
| Scope                       | `interprocedural`                      |
| Precision                   | `high`                                 |
| Enabled by default          | yes                                    |
| Configuration               | `detect_dictionary_compare_mode_order` |
| Inline suppression          | yes                                    |
| Blocks source preflight     | no                                     |
| Real-time editor diagnostic | yes                                    |
| Fix available               | no                                     |

## VBA231

**Repeated Dictionary loop materialization.** Dictionary Keys or Items is repeatedly materialized inside a loop instead of being cached or enumerated once.

| Property                    | Value                                    |
| --------------------------- | ---------------------------------------- |
| Family                      | `analyze`                                |
| Category                    | `performance`                            |
| Evidence class              | `inference`                              |
| Compile-equivalent          | no                                       |
| Default severity            | `warning`                                |
| Supported severities        | `warning`                                |
| Surfaces                    | `analyze`, `lsp`                         |
| Scope                       | `interprocedural`                        |
| Precision                   | `high`                                   |
| Enabled by default          | yes                                      |
| Configuration               | `detect_dictionary_loop_materialization` |
| Inline suppression          | yes                                      |
| Blocks source preflight     | no                                       |
| Real-time editor diagnostic | yes                                      |
| Fix available               | no                                       |

## VBA232

**Inconsistent Dictionary key normalization.** The same Dictionary key source is used with inconsistent raw, case-normalized, or trimmed forms.

| Property                    | Value                                 |
| --------------------------- | ------------------------------------- |
| Family                      | `analyze`                             |
| Category                    | `runtime-safety`                      |
| Evidence class              | `inference`                           |
| Compile-equivalent          | no                                    |
| Default severity            | `warning`                             |
| Supported severities        | `warning`                             |
| Surfaces                    | `analyze`, `lsp`                      |
| Scope                       | `procedure-local`                     |
| Precision                   | `high`                                |
| Enabled by default          | yes                                   |
| Configuration               | `detect_dictionary_key_normalization` |
| Inline suppression          | yes                                   |
| Blocks source preflight     | no                                    |
| Real-time editor diagnostic | yes                                   |
| Fix available               | no                                    |

## VBA233

**Undefined late-bound Dictionary constant.** A late-bound Dictionary CompareMode uses a Scripting enum constant that is not defined in the project.

| Property                    | Value                                    |
| --------------------------- | ---------------------------------------- |
| Family                      | `analyze`                                |
| Category                    | `correctness`                            |
| Evidence class              | `inference`                              |
| Compile-equivalent          | no                                       |
| Default severity            | `warning`                                |
| Supported severities        | `warning`                                |
| Surfaces                    | `analyze`, `lsp`                         |
| Scope                       | `project-wide`                           |
| Precision                   | `high`                                   |
| Enabled by default          | yes                                      |
| Configuration               | `detect_late_bound_dictionary_constants` |
| Inline suppression          | yes                                      |
| Blocks source preflight     | no                                       |
| Real-time editor diagnostic | yes                                      |
| Fix available               | no                                       |

## VBA234

**Collection mutation during iteration.** A Collection is mutated while a For Each loop is iterating the same collection instance.

| Property                    | Value                                  |
| --------------------------- | -------------------------------------- |
| Family                      | `analyze`                              |
| Category                    | `runtime-safety`                       |
| Evidence class              | `runtime-safety`                       |
| Compile-equivalent          | no                                     |
| Default severity            | `warning`                              |
| Supported severities        | `warning`                              |
| Surfaces                    | `analyze`, `lsp`                       |
| Scope                       | `interprocedural`                      |
| Precision                   | `high`                                 |
| Enabled by default          | yes                                    |
| Configuration               | `detect_collection_iteration_mutation` |
| Inline suppression          | yes                                    |
| Blocks source preflight     | no                                     |
| Real-time editor diagnostic | yes                                    |
| Fix available               | no                                     |

## VBA235

**Collection index origin confusion.** A one-based Collection is accessed with a zero index or an unadjusted zero-based loop or array index.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `runtime-safety`                 |
| Evidence class              | `runtime-safety`                 |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`                        |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `high`                           |
| Enabled by default          | yes                              |
| Configuration               | `detect_collection_index_origin` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA236

**Unsafe command construction.** A process-launch command may combine external input, an unquoted executable path, secrets, or unobserved execution.

| Property                    | Value                                |
| --------------------------- | ------------------------------------ |
| Family                      | `analyze`                            |
| Category                    | `security`                           |
| Evidence class              | `policy`                             |
| Compile-equivalent          | no                                   |
| Default severity            | `warning`                            |
| Supported severities        | `warning`                            |
| Surfaces                    | `analyze`, `lsp`                     |
| Scope                       | `procedure-local`                    |
| Precision                   | `medium`                             |
| Enabled by default          | yes                                  |
| Configuration               | `detect_unsafe_command_construction` |
| Inline suppression          | yes                                  |
| Blocks source preflight     | no                                   |
| Real-time editor diagnostic | yes                                  |
| Fix available               | no                                   |

## VBA237

**Suppressed error propagation.** A procedure or call boundary loses failure information by returning normally without rethrowing, returning a success flag that the caller ignores, or otherwise failing to signal failure.

| Property                    | Value                                  |
| --------------------------- | -------------------------------------- |
| Family                      | `analyze`                              |
| Category                    | `runtime-safety`                       |
| Evidence class              | `runtime-safety`                       |
| Compile-equivalent          | no                                     |
| Default severity            | `warning`                              |
| Supported severities        | `warning`                              |
| Surfaces                    | `analyze`, `lsp`                       |
| Scope                       | `interprocedural`                      |
| Precision                   | `high`                                 |
| Enabled by default          | yes                                    |
| Configuration               | `detect_error_suppression_propagation` |
| Inline suppression          | yes                                    |
| Blocks source preflight     | no                                     |
| Real-time editor diagnostic | yes                                    |
| Fix available               | no                                     |

## VBA238

**Loop-invariant Excel object resolution.** A loop repeatedly resolves an Excel object-model member chain that does not depend on the loop variable and could be cached outside the loop.

| Property                    | Value                                           |
| --------------------------- | ----------------------------------------------- |
| Family                      | `analyze`                                       |
| Category                    | `performance`                                   |
| Evidence class              | `maintainability`                               |
| Compile-equivalent          | no                                              |
| Default severity            | `warning`                                       |
| Supported severities        | `warning`                                       |
| Surfaces                    | `analyze`, `lsp`                                |
| Scope                       | `procedure-local`                               |
| Precision                   | `medium`                                        |
| Enabled by default          | yes                                             |
| Configuration               | `detect_loop_invariant_excel_object_resolution` |
| Inline suppression          | yes                                             |
| Blocks source preflight     | no                                              |
| Real-time editor diagnostic | yes                                             |
| Fix available               | no                                              |

## VBA239

**Unsafe SQL construction.** A SQL statement may combine external input, dynamic identifiers, locale-sensitive values, manual quoting, or wildcard input before execution.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `security`                       |
| Evidence class              | `policy`                         |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`                        |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `medium`                         |
| Enabled by default          | yes                              |
| Configuration               | `detect_unsafe_sql_construction` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA240

**Risky module-level mutable state.** Module-level mutable state has project-wide readers and writers that create hidden lifecycle coupling.

| Property                    | Value                       |
| --------------------------- | --------------------------- |
| Family                      | `analyze`                   |
| Category                    | `reliability`               |
| Evidence class              | `maintainability`           |
| Compile-equivalent          | no                          |
| Default severity            | `warning`                   |
| Supported severities        | `warning`                   |
| Surfaces                    | `analyze`                   |
| Scope                       | `project-wide`              |
| Precision                   | `medium`                    |
| Enabled by default          | no                          |
| Configuration               | `detect_risky_module_state` |
| Inline suppression          | yes                         |
| Blocks source preflight     | no                          |
| Real-time editor diagnostic | no                          |
| Fix available               | no                          |

## VBA241

**Repeated ReDim Preserve inside loop.** A dynamic array is repeatedly resized with ReDim Preserve inside a loop, potentially copying the existing array on every iteration.

| Property                    | Value                            |
| --------------------------- | -------------------------------- |
| Family                      | `analyze`                        |
| Category                    | `performance`                    |
| Evidence class              | `maintainability`                |
| Compile-equivalent          | no                               |
| Default severity            | `warning`                        |
| Supported severities        | `warning`, `information`         |
| Surfaces                    | `analyze`, `lsp`                 |
| Scope                       | `procedure-local`                |
| Precision                   | `medium`                         |
| Enabled by default          | yes                              |
| Configuration               | `detect_redim_preserve_in_loops` |
| Inline suppression          | yes                              |
| Blocks source preflight     | no                               |
| Real-time editor diagnostic | yes                              |
| Fix available               | no                               |

## VBA242

**Expensive full-range operation.** A costly Excel operation targets an entire row, column, worksheet, or an unbounded UsedRange where a bounded range is likely intended.

| Property                    | Value                                    |
| --------------------------- | ---------------------------------------- |
| Family                      | `analyze`                                |
| Category                    | `performance`                            |
| Evidence class              | `maintainability`                        |
| Compile-equivalent          | no                                       |
| Default severity            | `information`                            |
| Supported severities        | `information`, `warning`                 |
| Surfaces                    | `analyze`, `lsp`                         |
| Scope                       | `procedure-local`                        |
| Precision                   | `medium`                                 |
| Enabled by default          | no                                       |
| Configuration               | `detect_expensive_full_range_operations` |
| Inline suppression          | yes                                      |
| Blocks source preflight     | no                                       |
| Real-time editor diagnostic | yes                                      |
| Fix available               | no                                       |

## VBA243

**Value2 performance opportunity.** A bulk or repeated Range.Value read or write may benefit from Range.Value2 when automatic Currency or Date coercion is not required.

| Property                    | Value                                     |
| --------------------------- | ----------------------------------------- |
| Family                      | `analyze`                                 |
| Category                    | `performance`                             |
| Evidence class              | `maintainability`                         |
| Compile-equivalent          | no                                        |
| Default severity            | `information`                             |
| Supported severities        | `information`, `warning`                  |
| Surfaces                    | `analyze`, `lsp`                          |
| Scope                       | `procedure-local`                         |
| Precision                   | `medium`                                  |
| Enabled by default          | no                                        |
| Configuration               | `detect_value2_performance_opportunities` |
| Inline suppression          | yes                                       |
| Blocks source preflight     | no                                        |
| Real-time editor diagnostic | yes                                       |
| Fix available               | no                                        |

## VBA244

**Recursive and cyclic procedure dependency.** Procedures form a recursive call cycle; the complete deterministic cycle path and any dangerous reachable effects are reported.

| Property                    | Value                          |
| --------------------------- | ------------------------------ |
| Family                      | `analyze`                      |
| Category                    | `reliability`                  |
| Evidence class              | `maintainability`              |
| Compile-equivalent          | no                             |
| Default severity            | `information`                  |
| Supported severities        | `information`, `warning`       |
| Surfaces                    | `analyze`                      |
| Scope                       | `project-wide`                 |
| Precision                   | `medium`                       |
| Enabled by default          | yes                            |
| Configuration               | `detect_procedure_call_cycles` |
| Inline suppression          | yes                            |
| Blocks source preflight     | no                             |
| Real-time editor diagnostic | no                             |
| Fix available               | no                             |

## VBA245

**Unsafe destructive file and path operation.** A destructive or state-dependent file operation may use an empty, root, relative, wildcard, traversing, overwritten, or externally derived path.

| Property                    | Value                     |
| --------------------------- | ------------------------- |
| Family                      | `analyze`                 |
| Category                    | `security`                |
| Evidence class              | `policy`                  |
| Compile-equivalent          | no                        |
| Default severity            | `warning`                 |
| Supported severities        | `warning`                 |
| Surfaces                    | `analyze`, `lsp`          |
| Scope                       | `procedure-local`         |
| Precision                   | `medium`                  |
| Enabled by default          | yes                       |
| Configuration               | `detect_unsafe_file_path` |
| Inline suppression          | yes                       |
| Blocks source preflight     | no                        |
| Real-time editor diagnostic | yes                       |
| Fix available               | no                        |

## VBA246

**Unsafe HTTP or TLS configuration.** An HTTP client configuration may expose credentials, weaken TLS validation, log authorization data, or immediately execute downloaded content.

| Property                    | Value                              |
| --------------------------- | ---------------------------------- |
| Family                      | `analyze`                          |
| Category                    | `security`                         |
| Evidence class              | `policy`                           |
| Compile-equivalent          | no                                 |
| Default severity            | `warning`                          |
| Supported severities        | `warning`                          |
| Surfaces                    | `analyze`, `lsp`                   |
| Scope                       | `procedure-local`                  |
| Precision                   | `medium`                           |
| Enabled by default          | yes                                |
| Configuration               | `detect_unsafe_http_configuration` |
| Inline suppression          | yes                                |
| Blocks source preflight     | no                                 |
| Real-time editor diagnostic | yes                                |
| Fix available               | no                                 |

## VBA247

**Missing or unlimited HTTP timeout.** A long-running HTTP operation may wait indefinitely because supported client timeouts are missing or explicitly unlimited.

| Property                    | Value                         |
| --------------------------- | ----------------------------- |
| Family                      | `analyze`                     |
| Category                    | `reliability`                 |
| Evidence class              | `policy`                      |
| Compile-equivalent          | no                            |
| Default severity            | `warning`                     |
| Supported severities        | `warning`                     |
| Surfaces                    | `analyze`, `lsp`              |
| Scope                       | `procedure-local`             |
| Precision                   | `medium`                      |
| Enabled by default          | yes                           |
| Configuration               | `detect_missing_http_timeout` |
| Inline suppression          | yes                           |
| Blocks source preflight     | no                            |
| Real-time editor diagnostic | yes                           |
| Fix available               | no                            |

## VBA248

**Opaque Boolean control arguments.** A procedure call passes Boolean literals positionally in a way that obscures the behavior being requested.

| Property                    | Value                             |
| --------------------------- | --------------------------------- |
| Family                      | `analyze`                         |
| Category                    | `maintainability`                 |
| Evidence class              | `maintainability`                 |
| Compile-equivalent          | no                                |
| Default severity            | `warning`                         |
| Supported severities        | `warning`                         |
| Surfaces                    | `analyze`, `lsp`                  |
| Scope                       | `procedure-local`                 |
| Precision                   | `medium`                          |
| Enabled by default          | no                                |
| Configuration               | `detect_opaque_boolean_arguments` |
| Inline suppression          | yes                               |
| Blocks source preflight     | no                                |
| Real-time editor diagnostic | yes                               |
| Fix available               | no                                |

## VBA249

**Deterministic runtime error.** Constant evaluation, type information, control-flow, or dataflow facts prove that an expression will fail at runtime.

| Property                    | Value                                 |
| --------------------------- | ------------------------------------- |
| Family                      | `analyze`                             |
| Category                    | `runtime-safety`                      |
| Evidence class              | `runtime-error`                       |
| Compile-equivalent          | no                                    |
| Default severity            | `error`                               |
| Supported severities        | `error`                               |
| Surfaces                    | `analyze`, `lsp`                      |
| Scope                       | `procedure-local`                     |
| Precision                   | `high`                                |
| Enabled by default          | yes                                   |
| Configuration               | `detect_deterministic_runtime_errors` |
| Inline suppression          | yes                                   |
| Blocks source preflight     | no                                    |
| Real-time editor diagnostic | yes                                   |
| Fix available               | no                                    |
