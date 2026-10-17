import { spawnSync } from "node:child_process";
import fs from "node:fs";
import { format } from "oxfmt";

const output = "internal/agentskill/templates/xlflow/references/forms-contract.md";
const args = ["run", "./scripts/docs/userform-skill-reference"];
const windows = process.platform === "win32";
const result = spawnSync(
  windows ? "powershell.exe" : "go",
  windows
    ? ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "./scripts/dev/go.ps1", ...args]
    : args,
  { encoding: "utf8", maxBuffer: 4 * 1024 * 1024, windowsHide: true },
);
if (result.error || result.status !== 0) {
  console.error(result.error?.message ?? result.stderr);
  process.exitCode = result.status || 1;
} else {
  const formatted = await format(output, result.stdout, { endOfLine: "lf" });
  if (formatted.errors.length) {
    console.error(formatted.errors);
    process.exitCode = 1;
  } else if (process.argv.includes("--check")) {
    const current = fs.existsSync(output) ? fs.readFileSync(output, "utf8") : "";
    if (current.replaceAll("\r\n", "\n") !== formatted.code) {
      console.error("UserForm skill contract is stale; run pnpm docs:generate-form-skill");
      process.exitCode = 1;
    }
  } else {
    fs.writeFileSync(output, formatted.code);
  }
}
