# ADR-0038: HTTP Transport Safety Ownership and Reliability Taxonomy

## Status

Accepted

## Context

ADR-0025 established conservative procedure-local source-to-sink analysis and
`VBA224` for generic external-input flows. Issue #470 requires API-specific
reasoning for XMLHTTP and WinHTTP object identity, HTTP request state, TLS
options, authorization logging, ADODB.Stream download-and-launch sequences,
and timeout configuration. These facts cannot be expressed as generic taint
alone. Timeout omissions also have a different operational meaning from
credential disclosure or weakened certificate validation.

The diagnostics expose structured JSON and a configurable development-origin
exception. Those are public contracts whose ownership and redaction policy
must remain stable across batch analysis and realtime/LSP projection.

## Decision

Add two default-enabled, warning-level, non-blocking, inline-suppressible,
procedure-local diagnostics available in batch and realtime/LSP analysis:

- `VBA246` (`detect_unsafe_http_configuration`) owns transport-security,
  credential-exposure, sensitive logging, sensitive module constants used by
  HTTP requests, and download-and-launch findings.
- `VBA247` (`detect_missing_http_timeout`) owns HTTP-client reliability
  findings for missing or definitely unbounded timeouts.

The split is semantic rather than API-based. A missing timeout does not imply a
confidentiality or integrity failure and therefore remains a reliability
policy finding even when the same request also produces `VBA246`.

Both rules recognize only clients established through supported early-bound
types, `New`, known `CreateObject` ProgIDs, or simple `Set` aliases:
`MSXML2.XMLHTTP`, `MSXML2.ServerXMLHTTP`, `WinHttp.WinHttpRequest`, and their
common versioned forms. `ADODB.Stream` is recognized only for the specific
response-body/save/launch sequence. Member names such as `Open` and `Send`, or
a variable named `request`, are not sufficient evidence by themselves.

`VBA246` owns the risk kinds `plain_http_credentials`, `credentials_in_url`,
`authorization_logging`, `certificate_validation_bypass`,
`obsolete_tls_protocol`, `sensitive_module_constant`, and
`download_and_execute`. `VBA247` reports `missing` and `unbounded` timeout
states for ServerXMLHTTP and WinHTTP. Unknown option or timeout values do not
produce definite claims.

`VBA224` remains the generic compatibility fallback for supported HTTP
source/sink flows. When `VBA246` produces a specialized finding for the same
source and sink, the `VBA224` projection is suppressed. Disabling `VBA246`
restores the generic projection; unrelated generic flows remain owned by
`VBA224`.

Add `[analyze].development_http_origins` as an explicit allowlist of absolute
HTTP origins written as `http://host[:port]`; any explicit path, including
`/`, is rejected. Matching is against the normalized origin exactly; host
casing and IPv6 notation are normalized and an explicit default `:80` is
removed. Loopback origins are development endpoints without configuration. This
exception suppresses only `plain_http_credentials`; it never excuses URL
userinfo, authorization logging, TLS validation bypass, obsolete TLS,
sensitive constants, or download-and-launch behavior.

The additive public contexts are `http_security` and `http_reliability`.
They contain classification metadata only: API, risk kind or timeout state,
header name when relevant, and a credential-free normalized origin. They do
not expose header values, URL userinfo, paths, queries, or fragments. Findings
that may contain credentials redact string literals and comments from
`nearby_code` as `[REDACTED]`.

## Consequences

- Consumers can distinguish security policy from reliability policy without
  parsing messages or severity.
- Supported object identity and API-specific certainty reduce false positives
  from unrelated COM objects with similar member names.
- Users can allow an intentional local development origin without weakening
  credential-in-URL, logging, TLS, or execution checks.
- Existing `VBA224` HTTP suppressions may need migration to `VBA246` while the
  specialized rule is enabled.
- Procedure-local analysis intentionally does not prove interprocedural setup,
  sanitization, timeout configuration, or launch behavior.
- Redacted JSON and nearby source trade some debugging detail for a stable
  guarantee that diagnostics do not become a second credential-disclosure
  channel.

## Alternatives Considered

1. **Keep all HTTP findings in `VBA224`.** Rejected because TLS options,
   request credentials, timeouts, and response-stream launch sequences require
   client-specific state and would make the generic dataflow contract
   provider-specific.
2. **Use one HTTP diagnostic for both security and timeout findings.** Rejected
   because timeout omission is a reliability concern and must be independently
   configurable and reportable.
3. **Allow development endpoints by hostname pattern or path.** Rejected
   because wildcard and prefix matching can unintentionally exempt attacker-
   controlled hosts. Exact normalized origins provide an auditable boundary.
4. **Include values or complete URLs in finding context.** Rejected because
   command output, editor logs, and diagnostic telemetry would then copy the
   sensitive data the rule is intended to protect.
5. **Warn on unresolved objects or dynamic TLS flags.** Rejected because the
   first contract requires conservative API-specific checks, not speculation
   from method names or runtime-dependent values.

## Evidence

- Issue #470 requirements and acceptance criteria.
- `docs/specs/vba-source-sink-dataflow.md` and ADR-0025.
- `internal/vba/procedureir`, `internal/vba/dataflow`, and the shared analyzer
  rule registry.
- `docs/specs/http-transport-security-analysis.md` and focused `VBA246` /
  `VBA247` tests.

## Related

- Issue #470
- ADR-0024
- ADR-0025
- ADR-0033
- ADR-0034
