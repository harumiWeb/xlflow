# Optional VBA maintainability diagnostics

Rules VB067-VB092 are independently configured, disabled by default, inline
suppressible, and published by `xlflow lint` and LSP. They do not block source
preflight. VB091 uses warning severity; the other rules use information.

| Rule    | Meaning                 | Configuration                              |
| ------- | ----------------------- | ------------------------------------------ |
| `VB067` | Empty If branch         | `[lint].detect_empty_if`                   |
| `VB068` | Empty Else branch       | `[lint].detect_empty_else`                 |
| `VB069` | Empty Case branch       | `[lint].detect_empty_case`                 |
| `VB070` | Empty For loop          | `[lint].detect_empty_for`                  |
| `VB071` | Empty For Each loop     | `[lint].detect_empty_for_each`             |
| `VB072` | Empty Do loop           | `[lint].detect_empty_do`                   |
| `VB073` | Empty While loop        | `[lint].detect_empty_while`                |
| `VB074` | Empty procedure         | `[lint].detect_empty_procedure`            |
| `VB075` | Empty module            | `[lint].detect_empty_module`               |
| `VB076` | Legacy Call statement   | `[lint].detect_legacy_call`                |
| `VB077` | Rem comment             | `[lint].detect_rem_comment`                |
| `VB078` | Error statement         | `[lint].detect_error_statement`            |
| `VB079` | Global declaration      | `[lint].detect_global_declaration`         |
| `VB080` | Let assignment          | `[lint].detect_let_assignment`             |
| `VB081` | Identifier type suffix  | `[lint].detect_identifier_type_suffix`     |
| `VB082` | While Wend syntax       | `[lint].detect_while_wend`                 |
| `VB083` | DefType directive       | `[lint].detect_def_type`                   |
| `VB084` | Redundant Step 1        | `[lint].detect_redundant_step_one`         |
| `VB085` | Omitted Step            | `[lint].require_explicit_step`             |
| `VB086` | Redundant Option Base 0 | `[lint].detect_redundant_option_base_zero` |
| `VB087` | Module Dim declaration  | `[lint].detect_module_dim`                 |
| `VB088` | Implicit Public member  | `[lint].detect_implicit_public`            |
| `VB089` | Multiple declarations   | `[lint].detect_multiple_declarations`      |
| `VB090` | Unused label            | `[lint].detect_unused_labels`              |
| `VB091` | Stop statement          | `[lint].detect_stop_statement`             |
| `VB092` | On Local Error syntax   | `[lint].detect_on_local_error`             |

Empty branch and loop rules report only when the CST proves that the body has no
executable statement. Comments, declarations, attributes, and Option statements
are not executable. A recovered or missing CST fails open. VB074 and VB075 apply
the same test to whole procedures and modules.
For conditional procedure headers, VB074 considers all parsed branch bodies and
the shared body. VB087 treats declarations in those bodies as procedure-local.

VB081 checks identifier type suffixes on source identifiers, excluding a member
bang selector. At call sites it reports only when a matching suffixed procedure
is declared in the same source; unknown or intrinsic calls remain unreported.
VB088 reads only the declaration header's visibility modifier, and skips
conditional procedure headers whose branch modifiers may differ. VB090 checks
named and numeric labels against references within the containing procedure and
fails open across conditional compilation. VB089 yields to the default-enabled
VB019 warning when a multi-name declaration has mixed explicit typing.

VB084 and VB085 encode opposite style policies and cannot both be enabled.
Existing `VBA271` and `VBA272` retain ownership of Option Base behavioral
inconsistencies; VB086 reports only a literal `Option Base 0` directive.

The VBE compile oracle accepted `Global`, `Let`, `Error`, and `On Local Error`
in fixtures `maintainability-global`, `maintainability-let`,
`maintainability-error`, and `maintainability-on-local-error`. These observations
establish parse eligibility, not that the syntax should be discouraged.
The parser represents `Error n` as a call statement; VB078 recognizes that exact
keyword form without adding a separate CST node.

<!-- xlflow-rule-contract: {"id":"VB067","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_if","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB068","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_else","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB069","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_case","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB070","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_for","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB071","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_for_each","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB072","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_do","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB073","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_while","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB074","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_empty_procedure","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB075","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_empty_module","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB076","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_legacy_call","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB077","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_rem_comment","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB078","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_error_statement","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB079","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_global_declaration","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB080","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_let_assignment","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB081","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_identifier_type_suffix","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB082","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_while_wend","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB083","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_def_type","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB084","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_redundant_step_one","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB085","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"require_explicit_step","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB086","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_redundant_option_base_zero","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB087","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_module_dim","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB088","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_implicit_public","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB089","family":"lint","category":"maintainability","default_severity":"information","scope":"file-local","realtime":true,"configuration_key":"detect_multiple_declarations","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB090","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_unused_labels","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB091","family":"lint","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_stop_statement","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VB092","family":"lint","category":"maintainability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_on_local_error","inline_suppressible":true,"preflight_blocking":false} -->
