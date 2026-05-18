# xlflow completion

Generate shell completion scripts through Cobra.

## Usage

```bash
xlflow completion [bash|fish|powershell|zsh]
```

## Options and Arguments

| Option / argument | Description                     | Default |
| ----------------- | ------------------------------- | ------- |
| `bash`            | Generate Bash completion.       | -       |
| `fish`            | Generate Fish completion.       | -       |
| `powershell`      | Generate PowerShell completion. | -       |
| `zsh`             | Generate Zsh completion.        | -       |

## Examples

```bash
xlflow completion powershell > xlflow.ps1
xlflow completion bash > xlflow.bash
```

## Notes

::: tip
Completion generation is pure CLI output and does not require Excel COM.
:::

## Related

- [installation](../installation)
- [version](./version)
