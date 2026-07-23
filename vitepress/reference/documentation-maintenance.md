# Documentation maintenance

Run the documentation contract before every release:

```bash
pnpm docs:check
pnpm docs:build
```

The check validates command-page coverage, required command guidance sections, onboarding routes, configuration keys, curated recovery codes, internal Markdown links, and generated CLI/diagnostic/error inventories.

## Release checklist

- [ ] New commands and flags are documented.
- [ ] New configuration options are documented.
- [ ] New error codes are documented and mapped to recovery guidance.
- [ ] New diagnostic IDs are present in the generated inventory.
- [ ] Breaking changes have a migration note.
- [ ] Tutorials still match the current workflow.
- [ ] Command examples have been verified against the release version.
- [ ] Japanese high-impact pages are updated when the English journey changes.
