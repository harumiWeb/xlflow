# Workbook Operation Coordination

This spec defines the internal command-policy and canonical workbook-identity
contracts used by workbook operation coordination. Cross-process lock acquisition,
busy diagnostics, waiting options, and coordination status output are outside the
current implementation and will build on these contracts.

See ADR-0016 for the architectural rationale.

## Command Policy Contract

The Go core owns one authoritative registry of executable command descriptors.
Each descriptor has a stable command ID, selectors used by the CLI and bridge,
and one coordination policy. Consumers query this registry instead of maintaining
hardcoded workbook-command lists.

The policy vocabulary is:

| Field                 | Values                                  | Meaning                                                                        |
| --------------------- | --------------------------------------- | ------------------------------------------------------------------------------ |
| `resource_scope`      | `none`, `workbook`, `excel_instance`    | Smallest resource boundary required by the operation.                          |
| `operation_kind`      | `read`, `mutate`, `execute`, `designer` | The operation's highest-risk behavior.                                         |
| `parallel_safe`       | boolean                                 | Whether the operation may proceed without exclusive coordination at its scope. |
| `retryable_when_busy` | boolean                                 | Whether a future caller may opt into retrying after resource contention.       |
| `default_wait_policy` | `fail`, `wait`                          | Whether acquisition should fail immediately or wait by default.                |

All initial policies use `default_wait_policy: fail`. Waiting remains explicit in
future work. A resource operation is retryable only when it is non-parallel-safe
and retrying after contention is meaningful. `parallel_safe: true` means that the
operation does not need exclusive acquisition at its declared scope; it does not
introduce a shared read-lock mode.

When flags can change a command's risk, the descriptor uses the most restrictive
behavior supported by that executable command unless the registry has an explicit
selector that distinguishes the variants. WSL delegation classification is
separate from coordination policy: it determines where a command runs, not what
resource safety the command requires.

### Command Selectors

Every executable Cobra leaf command resolves to exactly one descriptor by its
normalized leaf path. Command containers, help, and generated completion commands
are not executable policy entries.

Bridge-backed commands resolve through the same registry before invocation. A
selector contains the bridge command and, for multiplexed bridge endpoints, the
request field such as `Action` that identifies the actual operation. The .NET
bridge must not duplicate the registry. Stable command IDs are internal identifiers
in this version; they are suitable for future serialization but are not yet a
public capabilities schema.

Registry enumeration is deterministic and returns defensive copies so callers
cannot mutate the authoritative definitions.

### Fail-Closed Behavior

An executable CLI command or bridge invocation without an explicit matching
descriptor returns a typed policy lookup failure. The CLI error code is
`coordination_policy_missing`. There is no implicit `resource_scope: none` or
other permissive fallback.

Coverage tests enumerate executable Cobra leaves and used bridge command/action
selectors. Adding a command without adding its policy must fail those tests.

### Initial Classification

The registry classifies commands conservatively:

- Source-only operations such as lint, format, analyze, LSP, source inspection,
  `test list`, and `form new` use `resource_scope: none`.
- Workbook inspection and synchronization reads, including pull, init, macro and
  form listing, formula pull, workbook diff, and workbook/sheet/range inspection, use
  `resource_scope: workbook`, `operation_kind: read`, and are not parallel-safe.
- Workbook creation and mutation, including new, push, rollback, save, session
  start/stop, runner changes, pack artifact generation, and worksheet/cell/UI edits, use
  `resource_scope: workbook`, `operation_kind: mutate`, and are not parallel-safe.
- VBA run and test use `resource_scope: workbook`, `operation_kind: execute`, and
  are not parallel-safe.
- UserForm migration, snapshot, build, specification application, image export,
  and Designer inspection use `resource_scope: workbook`,
  `operation_kind: designer`, and are not parallel-safe.
- Top-level status and session status are read-only observers. They use
  `resource_scope: workbook`, `operation_kind: read`, and `parallel_safe: true`
  so they can report a future busy state without first waiting for that state to
  clear.
- Environment checks, active-workbook attachment, and Excel process cleanup use
  `resource_scope: excel_instance` when they inspect or mutate state broader than
  one configured workbook.

## Canonical Workbook Identity

A workbook identity contains:

- `CanonicalPath`, a human-readable canonical Windows path for diagnostics; and
- `LockID`, an opaque deterministic identifier safe for a future cross-process
  synchronization primitive.

The API requires an explicit base directory and workbook path. A relative workbook
path is resolved against that base; identity generation never reads the process
working directory implicitly and does not require the workbook to be open or to
exist.

### Normalization

Identity generation applies these steps in order:

1. Resolve a relative workbook path against the explicit base directory and make
   the result absolute.
2. Clean redundant separators and `.` or `..` path segments.
3. For an existing path, resolve symbolic links and junctions on a best-effort
   basis. If resolution fails, retain the lexical absolute path.
4. Convert Windows extended path prefixes to their normal drive or UNC form while
   preserving UNC semantics.
5. Normalize separators to the canonical Windows form.
6. Normalize drive-letter casing for the diagnostic path.
7. Build a case-insensitive comparison key from the canonical path.

Equivalent path casing, slash direction, drive-letter casing, relative/absolute
forms, and redundant lexical segments therefore produce the same comparison key.
UNC paths retain their leading UNC semantics and are not converted to local-drive
paths.

`CanonicalPath` preserves a useful diagnostic representation; callers must use
`LockID`, not display-path string equality, as the future synchronization key.

### Lock Identifier

`LockID` is the lowercase hexadecimal SHA-256 digest of a domain separator plus
the normalized comparison key, with a fixed ASCII prefix identifying xlflow
workbook coordination. The domain separator prevents accidental reuse of the same
digest namespace for another resource type. The resulting identifier contains
only fixed ASCII prefix characters and hexadecimal digits and never includes the
workbook path.

SHA-256 collision risk is negligible for practical workbook coordination. The
operating-system lock implementation may add a primitive-specific namespace
prefix later but must preserve the opaque identity and must not embed
`CanonicalPath` in the primitive name.

### WSL Boundary

The canonicalizer accepts Windows paths; it does not translate `/mnt/...` or
other Linux paths. Under ADR-0011, WSL translates supported absolute paths with
`wslpath -w` and delegates the complete command to Windows `xlflow.exe`. Identity
generation then runs in the Windows process using the translated path. A supported
WSL-mounted path and its direct Windows representation therefore share an identity.

WSL-only paths such as `/home/...` remain unsupported for delegated workbook
commands and fail in the delegation layer before identity generation.

### Unresolved Aliases

Best-effort link resolution does not promise that every path alias converges.
Unless Windows resolves them through the normal path/link resolution above, the
following may produce distinct identities even when they ultimately reach the
same file:

- a mapped drive and its backing UNC share;
- DFS targets, DNS aliases, and alternate server names;
- hard links and 8.3 short names;
- network redirector aliases;
- paths that differ only by Unicode normalization; and
- paths in directories explicitly configured for case-sensitive lookup.

Broken, inaccessible, or not-yet-created symlink and junction paths use their
lexically normalized absolute form. Diagnostic output may expose the canonical
workbook path, but synchronization primitive names must expose only `LockID`.

## Out of Scope for This Version

This contract does not add:

- a Windows named mutex, file lock, daemon, or other cross-process lock primitive;
- lock acquisition, release, crash recovery, or owner metadata;
- `workbook_busy` or timeout diagnostics and exit-code mappings;
- `--wait`, timeout, cancellation, or FIFO queue behavior;
- coordination fields in `session status`;
- a public policy/capabilities endpoint or serialized bridge schema; or
- a separate UserForm Designer lock.

Those features must consume this registry and identity contract rather than
creating parallel policy lists or workbook key algorithms.
