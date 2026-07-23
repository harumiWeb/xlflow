# Work from WSL

Install the Linux frontend in WSL and the Windows `xlflow.exe` plus `.NET` bridge on Windows. This split is intentional: your editor and agent can stay in WSL, but Excel can run only on Windows. Keep the project under `/mnt/c`, `/mnt/d`, or another Windows-mounted path so both sides can refer to the same files.

Before using a real workbook, run `xlflow version` and `xlflow doctor --json` in WSL. A successful doctor confirms that xlflow can reach the Windows executable and Excel.

```bash
curl -fsSL https://harumiweb.github.io/xlflow/install.sh | sh
cd /mnt/c/dev/my-vba-project
xlflow doctor --json
xlflow lint --json
xlflow analyze --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --diagnostic --session --json
xlflow save --session --json
xlflow session stop --json
```

Source-only commands stay in WSL. Excel-backed commands delegate to Windows. The session commands make this boundary invisible once they work: `push`, `run`, and `save` all act on the same Windows Excel session. If discovery fails, set `XLFLOW_WINDOWS_EXE` to the Windows executable path. WSL-only paths such as `/home/user/project` cannot be translated for Excel delegation; see [WSL troubleshooting](../help/troubleshooting#wsl).
