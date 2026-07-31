# Documentation maintenance

Run the documentation contract before every release:

```bash
pnpm docs:check
pnpm docs:build
```

The check validates command-page coverage, required command guidance sections, onboarding routes, configuration keys, curated recovery codes, internal Markdown links, and generated CLI/diagnostic/error inventories. The diagnostic catalog is generated directly from `internal/staticanalysis/rules/registry.json`; it does not require a Go build.

## Release checklist

- [ ] New commands and flags are documented.
- [ ] New configuration options are documented.
- [ ] New error codes are documented and mapped to recovery guidance.
- [ ] New diagnostic IDs and metadata are present in the canonical registry and generated catalog.
- [ ] Breaking changes have a migration note.
- [ ] Tutorials still match the current workflow.
- [ ] Command examples have been verified against the release version.
- [ ] Japanese high-impact pages are updated when the English journey changes.
