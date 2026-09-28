# Rubberduck parity rule default review (2026-09-28)

The third-party corpus evaluation profile explicitly enables the new opt-in
rules without changing production configuration. Its current snapshots contain
9,237 net additional observations. Snapshot presence is not a correctness judgment:
11 newly observed VBA276 occurrences were checked against their source and
rule contract and added to the review ledger; the other new observations remain
unreviewed. Earlier ledger entries remain separate reviewed evidence.

The profile excludes VB085 and VBA277 because their policies conflict with
VB084 and VBA273, respectively. It also respects Excel host exclusions for
generic VBA and Access projects. Focused fixtures continue to cover the two
excluded rules.

The complete current observation counts for the new family are below. Each
entry is `rule: occurrences / projects`; zero means this corpus provides no
real-world example, not that the rule has been proved ineffective.

```text
VB067: 42/8    VB068: 57/4    VB069: 71/9    VB070: 0/0
VB071: 7/2     VB072: 1/1     VB073: 2/1     VB074: 143/5
VB075: 5/2     VB076: 1742/8  VB077: 0/0     VB078: 0/0
VB079: 0/0     VB080: 16/1    VB081: 372/10  VB082: 95/4
VB083: 0/0     VB084: 1/1     VB085: 0/0     VB086: 0/0
VB087: 44/3    VB088: 285/9   VB089: 253/7   VB090: 21/3
VB091: 0/0     VB092: 0/0
VBA253: 473/12 VBA254: 3582/16 VBA255: 21/2    VBA256: 52/10
VBA257: 289/12 VBA258: 5/2     VBA259: 3/1     VBA260: 0/0
VBA261: 0/0    VBA262: 0/0     VBA265: 45/4    VBA266: 369/5
VBA267: 40/2   VBA268: 1/1     VBA269: 331/16  VBA270: 1/1
VBA271: 0/0    VBA272: 0/0     VBA273: 1168/11 VBA274: 188/14
VBA275: 642/13 VBA276: 11/4    VBA277: 0/0    VBA278: 31/2
VBA279: 0/0    VBA280: 18/6    VBA281: 18/2    VBA282: 212/5
VBA283: 6/5
```

## Default decisions

The counts below are current snapshot occurrences and distinct third-party
projects. TP and FP are reviewed counts already committed to the ledger,
including historical reviews; "ledger FP" counts are forbidden evidence for
earlier analyzer defects, not current emissions.

| Rule   | Occurrences | Projects | Reviewed |  TP | Ledger FP | Default | Decision                                                                                                                                            |
| ------ | ----------: | -------: | -------: | --: | --------: | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| VBA259 |           3 |        1 |        3 |   3 |         0 | Off     | All three share one source line; insufficient independent evidence.                                                                                 |
| VBA266 |         369 |        5 |      369 | 369 |         0 | Off     | 359 reports come from two projects; default warning volume is excessive for a maintainability rule.                                                 |
| VBA270 |           1 |        1 |        1 |   1 |         0 | Off     | Needs independent Excel behavior evidence for the UDF collision consequence.                                                                        |
| VBA271 |           0 |        0 |        0 |   0 |         0 | Off     | No real-world observation.                                                                                                                          |
| VBA272 |           0 |        0 |        0 |   0 |         0 | Off     | No real-world observation.                                                                                                                          |
| VBA276 |          11 |        4 |       11 |  11 |         0 | On      | Every reviewed source declares a Property Let/Set value parameter `ByRef` despite its `ByVal` semantics.                                            |
| VBA279 |           0 |        0 |        0 |   0 |         0 | Off     | No real-world observation; the Excel document-copy hazard needs runtime evidence.                                                                   |
| VBA282 |         212 |        5 |      212 | 212 |         0 | Off     | Self-name use can deliberately address the predeclared instance; source-fact agreement alone does not prove a broadly actionable warning.           |
| VBA283 |           6 |        5 |        7 |   6 |         1 | On      | Six source-specific semantic misuses across five projects; the known Optional Variant false positive was fixed and is forbidden by ledger evidence. |
| VB091  |           0 |        0 |        0 |   0 |         0 | Off     | No real-world observation of `Stop` in this corpus, so debug-module noise is unknown.                                                               |
| VBA256 |          52 |       10 |       84 |  52 |        32 | Off     | Dead stores can preserve intentional RHS effects, and the earlier FP volume warrants a separate noise review.                                       |
| VBA258 |           5 |        2 |        5 |   5 |         0 | Off     | Discarded return is largely API/design signal.                                                                                                      |
| VBA265 |          45 |        4 |       50 |  45 |         5 | Off     | Five earlier TP labels in std-vba were corrected: IEnumVARIANT vtable callbacks are signature-constrained.                                          |
| VBA273 |        1168 |       11 |        2 |   0 |         2 | Off     | Callback declarations and explicitness preferences are unsuitable for the default policy.                                                           |
| VBA275 |         642 |       13 |        5 |   0 |         5 | Off     | `AddressOf` callback signatures cannot safely be changed to ByVal; broader ByVal advice remains a style decision.                                   |
| VBA281 |          18 |        2 |       18 |  18 |         0 | Off     | Public exposure of interface/event members is partly API policy.                                                                                    |

This pass found and fixed the VBA265 `AddressOf` false positive in a focused
fixture and five real std-vba callback parameters. The old ledger rationale
had said no external-entry constraint existed; the module actually installs
`IEnumVARIANT_Skip`, `Reset`, and `Clone` in its COM vtable with `AddressOf`.
The five entries are now forbidden FP evidence with a focused vtable regression
test, and their snapshot observations are removed. The same signature check
also exposed two VBA273 and five VBA275 callback warnings in std-vba. These
seven reports were reproduced, fixed, and recorded as forbidden evidence with
their exact ranges. No new false positive was found among the 11 newly
reviewed VBA276 occurrences. The VBA283 ledger FP was fixed in an earlier
pass; no unresolved reviewed FP remains for either promoted rule.

The remaining VB067–VB090/VB092 lint rules and VBA253–VBA255, VBA257,
VBA260–VBA262, VBA268–VBA269, VBA273–VBA275/VBA277, VBA278, and VBA280 remain
opt-in. Their reported source facts often describe valid legacy syntax,
intentional VBA idioms, initialization defaults, or API/style preferences.
In particular, `Application.Match` has a different error-return contract from
`WorksheetFunction.Match`, and implicit versus redundant explicit `ByRef`
cannot both be recommended by a single default policy. No public strict or
modernize profile is introduced here; the existing per-rule configuration
continues to serve that need without expanding the CLI contract.
The VBA273/VBA275 callback false positives found here are fixed, but the
remaining style and signature tradeoffs still need a separate source-level
review before either rule could be reconsidered for default use.

## Precision and performance interpretation

The committed reviewed ledger has 10,566 reviewed occurrences: 7,670 TP and
2,896 remediated FP, plus 90 reviewed allowed collisions. Another 9,231
snapshot occurrences are unreviewed. Reviewed-only precision is 72.6%; it is
not an estimate of precision for all emitted warnings. For VBA276 the reviewed
count is 11 TP / 0 FP, and for VBA283 it is 6 TP / 1 fixed FP.
Before this pass the ledger had 10,548 reviewed occurrences (7,664 TP,
2,884 FP, 90 allowed collisions). This pass added 11 VBA276 TP and seven
callback FP reviews, and corrected five old VBA265 TP labels to FP.

The default-on performance comparison uses the same materialized ROneCOne and
std-vba projects and production-default configuration on both sides of each
rule toggle. Each benchmark was run three times at one analysis per run on a
Windows i7-12700; the table gives the median, including cache effects equally
on both sides. VBA283 is toggled with VBA276 already enabled.

| Toggle / project  | Before s/op | After s/op |    Before B/op |     After B/op | Before allocs/op | After allocs/op |
| ----------------- | ----------: | ---------: | -------------: | -------------: | ---------------: | --------------: |
| VBA276 / ROneCOne |      12.226 |     12.082 |  6,899,904,032 |  6,911,409,208 |       63,431,781 |      63,700,396 |
| VBA276 / std-vba  |      47.966 |     47.858 | 20,701,647,504 | 20,715,421,176 |      138,436,268 |     138,824,026 |
| VBA283 / ROneCOne |      11.907 |     11.984 |  6,912,859,288 |  6,914,343,784 |       63,691,366 |      63,694,944 |
| VBA283 / std-vba  |      47.960 |     48.057 | 20,718,624,352 | 20,720,113,640 |      138,816,284 |     138,826,974 |

The `BenchmarkDefaultVBA276Promotion` benchmark initially found 17–25% more
bytes allocated because enabling VBA276 unnecessarily built interprocedural
parameter-mutation summaries. The gate was narrowed to the rules that consume
those summaries, without changing VBA276 eligibility. Its final median
allocation increase is below 0.2% in both projects; VBA283 adds below 0.03%.
The measured time differences are below 1% and do not establish a material
default-analysis regression. No LSP timing benchmark was run; batch/realtime
diagnostic parity remains covered by focused tests.
Raw logs: [VBA276 before optimization](performance/vba276-initial.txt),
[VBA276 after optimization](performance/vba276-optimized.txt), and
[VBA283](performance/vba283.txt).

Commands used for the isolated measurement:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkDefaultVBA276Promotion$' -benchmem -benchtime=1x -count=3 -timeout=25m
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/corpus -run '^$' -bench '^BenchmarkDefaultVBA283Promotion$' -benchmem -benchtime=1x -count=3 -timeout=25m
```

## Independent final review (pass 1) and remediation

An independent Orca-supervised Codex worker reviewed the committed change
read-only (verdict: approve-with-notes). One in-scope P2 was confirmed and
fixed in the follow-up commit:

- **P2 cross-module `AddressOf` signature constraint**: VBA resolves a bare
  `AddressOf` target across public procedures in every standard module, but
  `parameterPassingSignatureConstrained` only consulted the candidate
  file's own `AddressOf` calls, so a callback declared in a dedicated
  callbacks module still received VBA273/VBA275/VBA277 advice. The fix adds
  a project-wide `addressOfEntryNames` set (populated only when a
  signature-constrained parameter-passing rule is enabled), widened in the
  realtime path from `projectDocuments` the same way `dynamicEntryNames` is,
  plus a two-module regression test. VBA265 needs no change: its candidates
  are `Private`-only, and `Private` procedures cannot be `AddressOf`
  targets from another module.

The P3 notes were also addressed: a TODO documents when the corpus metrics
floors can be tightened back to exact totals, `detect_invalid_ismissing_usage`
was added to the `[analyze]` config-key table, and stale comments plus the
VBA276 spec exclusion note were corrected. The reviewer's two
unsupported/follow-up observations (realtime private-name-collision parity
for VBA265, and `Application.Run` string-literal names not constraining
parameter-passing signatures) are pre-existing design boundaries, not
defects introduced by this change; the realtime `addressOfEntryNames`
widening now shares the VBA265 scan loop so parity bookkeeping stays
symmetric.

A second independent review pass on the remediation commit found no
in-scope defect (verdict: approve-with-notes). Its three notes were:
the realtime widening loop lacked dedicated coverage (fixed by a
`SourceRealtimeFindings*ProjectContext` regression test where the
`AddressOf` call sits outside the analyzed document's type/call closure),
the name-only project-wide set can over-suppress a same-named non-target
procedure (accepted under the analyzer's documented suppression-only
fail-open convention, matching the existing `dynamicEntryNames`
contract), and the widening adds an O(project) per-refresh scan to the
realtime path under the default config (bounded, memoized in batch via
`unusedDeclOnce`, and consistent with the existing widening pattern).
