import fs from "node:fs";
import path from "node:path";

const repo = path.resolve(".");
const docs = path.join(repo, "vitepress");
const requiredCommandPages = [
  "capabilities",
  "new",
  "init",
  "pack",
  "doctor",
  "attach",
  "backup",
  "list",
  "form",
  "formulas",
  "pull",
  "push",
  "rollback",
  "status",
  "session",
  "save",
  "recovery",
  "runner",
  "run",
  "export-image",
  "edit",
  "macros",
  "ui",
  "test",
  "type-db",
  "diff",
  "inspect",
  "inspect-gui",
  "lint",
  "lsp",
  "fmt",
  "analyze",
  "check",
  "module",
  "completion",
  "process",
  "skill",
  "version",
  "update",
  "generate",
];
const requiredHeadings = [
  "## When to use this command",
  "## Prerequisites",
  "## What this command reads and changes",
  "## Effect on source-of-truth state",
  "## Common workflows",
  "## Common failures",
];

const failures = [];
for (const name of requiredCommandPages) {
  const file = path.join(docs, "commands", `${name}.md`);
  if (!fs.existsSync(file)) {
    failures.push(`missing command page: ${name}`);
    continue;
  }
  const source = fs.readFileSync(file, "utf8");
  for (const heading of requiredHeadings) {
    if (!source.includes(heading)) failures.push(`${name}: missing ${heading}`);
  }
}

const requiredPages = [
  "choose-workflow.md",
  "tutorials/index.md",
  "tutorials/existing-workbook.md",
  "tutorials/ai-agent.md",
  "vscode/index.md",
  "vscode/troubleshooting.md",
  "help/troubleshooting.md",
  "help/faq.md",
];
for (const relative of requiredPages) {
  if (!fs.existsSync(path.join(docs, relative)))
    failures.push(`missing onboarding page: ${relative}`);
}

for (const name of fs.readdirSync(path.join(docs, "demos"))) {
  if (!name.endsWith(".md") || name === "index.md") continue;
  const source = fs.readFileSync(path.join(docs, "demos", name), "utf8");
  if (!source.includes("<!-- xlflow-demo-case-study -->"))
    failures.push(`demo is missing case-study sections: ${name}`);
}

const errorReference = fs.readFileSync(path.join(docs, "reference/error-codes.md"), "utf8");
for (const code of [
  "vbide_access_denied",
  "macro_not_found",
  "macro_timeout",
  "source_preflight_failed",
  "vba_compile_failed",
  "workbook_recovery_required",
  "windows_xlflow_not_found",
  "wsl_project_path_unsupported",
  "wsl_path_translation_failed",
  "windows_xlflow_execution_failed",
]) {
  if (!errorReference.includes(code)) failures.push(`error code missing from reference: ${code}`);
}

const configSource = fs.readFileSync(path.join(repo, "internal/config/config.go"), "utf8");
const configReference = fs.readFileSync(path.join(docs, "reference/config-file.md"), "utf8");
for (const [, key] of configSource.matchAll(/toml:"([^,"]+)"/g)) {
  if (!configReference.includes(key)) failures.push(`config key missing from reference: ${key}`);
}

if (failures.length) {
  console.error(failures.join("\n"));
  process.exitCode = 1;
} else {
  console.log(`documentation contract OK (${requiredCommandPages.length} command pages checked)`);
}
