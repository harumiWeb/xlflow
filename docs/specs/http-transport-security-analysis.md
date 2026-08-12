# HTTP Transport Security Analysis

<!-- xlflow-rule-contract: {"id":"VBA246","family":"analyze","category":"security","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unsafe_http_configuration","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA247","family":"analyze","category":"reliability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_missing_http_timeout","inline_suppressible":true,"preflight_blocking":false} -->

`VBA246` and `VBA247` are default-enabled, warning-level, non-blocking,
procedure-local diagnostics available to batch analysis and realtime/LSP
diagnostics. `VBA246` owns transport-security and credential-exposure policy;
`VBA247` independently owns timeout reliability. Disable them with
`[analyze].disabled_rules`, their legacy-compatible Boolean keys, or an inline
suppression for an intentional local exception.

## Supported clients and identity

The analysis recognizes early-bound declarations, `As New`, `New` expressions,
known `CreateObject` ProgIDs, and simple `Set` aliases for:

- `MSXML2.XMLHTTP`, including common versioned type and ProgID forms;
- `MSXML2.ServerXMLHTTP`, including `ServerXMLHTTP60` and versioned ProgIDs;
- `WinHttp.WinHttpRequest`, including `WinHttpRequest.5.1`; and
- `ADODB.Stream` where a supported HTTP response is saved and launched.

An arbitrary object with `.Open`, `.Send`, `.SetTimeouts`, or a request-like
variable name is not an HTTP client. Unsupported receivers, dynamic member
names, unresolved aliases, interprocedural state, and runtime-only option
values remain unknown and do not establish a finding.

## `VBA246` security risks

`VBA246` reports these stable `risk_kind` values when the required values and
API state can be determined:

- `plain_http_credentials`: an `http://` request carries Open user/password
  arguments, `SetCredentials`, or an Authorization, Proxy-Authorization,
  Cookie, or API-key/token-like header;
- `credentials_in_url`: a reconstructable HTTP or HTTPS URL contains userinfo;
- `authorization_logging`: an Authorization value, a Bearer/Basic credential,
  or a simple alias of that value reaches `Debug.Print`, `Print #`, or
  `XlflowDebug.Log`;
- `certificate_validation_bypass`: a known nonzero WinHTTP
  `SslErrorIgnoreFlags`, explicit revocation-check disabling, or a known
  ServerXMLHTTP certificate-ignore option is used;
- `obsolete_tls_protocol`: a known WinHTTP `SecureProtocols` value enables
  SSL 2.0, SSL 3.0, TLS 1.0, or TLS 1.1. A known TLS 1.2/TLS 1.3-only value is
  accepted;
- `sensitive_module_constant`: a non-placeholder module-level `Const` stores a
  Bearer/Basic credential, API key, token, or comparable sensitive header value
  that is used by the HTTP request; and
- `download_and_execute`: a supported HTTP response body is written to an
  ADODB.Stream, saved to the same `.exe`, `.com`, `.bat`, `.cmd`, `.ps1`,
  `.vbs`, `.js`, `.jse`, `.wsf`, `.hta`, or `.msi` path, and then reaches a
  launcher recognized by `VBA236` in the same procedure.

Saving without launching, launching a different path, saving to an
unrecognized extension, or overwriting the tracked content/path before launch
does not satisfy `download_and_execute`. `VBA223` continues to own the fact
that a credential-like constant is stored in source; `VBA246` owns the HTTP-use
policy and does not expose the constant value.

When `VBA246` emits a specialized finding for the same HTTP source/sink flow,
the generic `VBA224` projection is suppressed. If `VBA246` is disabled,
`VBA224` remains the compatibility fallback for generic flows it already
supports. Other `VBA224` findings are unaffected.

## `VBA247` timeout reliability

`VBA247` evaluates `Send` on ServerXMLHTTP and WinHTTP receivers. A request is
reported with timeout state `missing` when no `SetTimeouts` configuration is
proven on every reaching path. It is reported as `unbounded` when a statically
known timeout component is `0` or `-1`. Finite timeout values are accepted.

A dynamic timeout value is not classified as missing or unbounded because its
runtime value is unknown. `MSXML2.XMLHTTP` is excluded because its exposed API
does not provide the same per-request timeout contract.

## Development HTTP origins

`[analyze].development_http_origins` accepts only absolute HTTP origins of the
form `http://host[:port]`. Userinfo, any path (including an explicit `/`), queries, fragments,
wildcards, and HTTPS entries are configuration errors. Host casing and IPv6
notation are normalized and an explicit default `:80` is removed before comparison; matching
is exact by scheme, host, and effective port. Subdomains and different ports
do not inherit an exception.

`localhost`, IPv4 loopback addresses in `127.0.0.0/8`, and IPv6 `::1` are
treated as development origins without configuration. A development-origin
match suppresses only `plain_http_credentials`. It does not suppress
`credentials_in_url`, `authorization_logging`, certificate or TLS findings,
`sensitive_module_constant`, or `download_and_execute`.

```toml
[analyze]
development_http_origins = ["http://dev-api.example.test:8080"]
```

## Public JSON and redaction

`VBA246` adds `http_security`; `VBA247` adds `http_reliability`. The contexts
contain only classification metadata: the recognized API, `risk_kind` or
timeout state, the header name when relevant, and a normalized origin stripped
of credentials. They never include a header value, URL userinfo, path, query,
or fragment.

For a finding that may expose a credential, string literals and comments in
`nearby_code` are replaced with `[REDACTED]`. Messages and suggestions also
describe the credential class rather than reproducing its value. This
redaction applies equally to CLI JSON and LSP-projected diagnostics.

## Non-goals

- whole-project or interprocedural request-state propagation;
- general URL reachability, hostname reputation, or certificate inspection;
- inferring TLS or timeout policy from unresolved numeric expressions;
- treating every downloaded file or every process launch as a proven chain;
- preflight blocking or Excel/VBE execution.
