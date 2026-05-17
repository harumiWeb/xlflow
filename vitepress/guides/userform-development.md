# UserForm Development

Recommended UserForm loop:

```bash
xlflow list forms --session --json
xlflow inspect form UserForm1 --designer --session --json
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --session --json
xlflow form build src/forms/specs/UserForm1.yaml --overwrite --session --json
xlflow form export-image UserForm1 --out artifacts/UserForm1.png --session --json
```

In sidecar mode, edit code-behind under `src/forms/code`.
