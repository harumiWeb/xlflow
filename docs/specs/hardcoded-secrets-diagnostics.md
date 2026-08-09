# Hardcoded Secret Diagnostics

`VBA223` is a default-enabled, high-precision `analyze` warning for likely
credentials embedded directly in VBA source. It runs in batch analysis and the
shared real-time editor path, does not block source preflight, and supports
normal inline suppression.

## Detection policy

The rule prefers structural evidence over entropy-only matching. It reports
direct string literals used as or containing:

- secret-bearing connection-string fields such as `Password`, `Pwd`,
  `Account Key`, `Client Secret`, or `Access Token`; identity-only fields such
  as `User ID`, `UID`, and `Username` are not sufficient evidence by
  themselves;
- `Bearer` and `Basic` authorization values and URL credentials in the form
  `user:password@host`;
- PEM private-key markers;
- curated provider-shaped access keys, API keys, and webhook URLs; and
- values assigned directly to clearly credential-related names such as
  `password`, `api_key`, `access_token`, `client_secret`, `private_key`,
  `credential`, or `token`.

For a qualified assignment target, only the final storage member determines
whether the name is credential-related. Receiver names do not contribute. URL,
URI, endpoint, and resource metadata names are not credential storage even when
they contain words such as `authorization` or `token`. Structured statements
are scanned only for syntax owned by that statement; a `With` statement does
not rescan the nested statements in its body.

The initial implementation does not perform arbitrary entropy matching,
complex data-flow propagation, or infer a secret from a string concatenation
without direct structural evidence.

## Placeholders and suppression

Empty values, environment-variable references, values that merely repeat their
assignment target, and obvious examples such as `abc`, `your-password`,
`changeme`, `example`, `dummy`, `placeholder`, `test`, and template markers are
ignored where the value can be identified safely.
Intentional fixtures that use a realistic value can suppress the finding with
`xlflow:disable-line VBA223` or `xlflow:disable-next-line VBA223`.

The project-wide setting is:

```toml
[analyze]
disabled_rules = ["VBA223"]
```

## Redaction contract

The detected value is never copied into `message`, `reason`, `suggestion`,
LSP diagnostics, or JSON fields. `nearby_code` redacts string literals and
comments in the surrounding source with the fixed marker `[REDACTED]`.
The diagnostic does not expose a provider prefix, suffix, length, or partial
value.
