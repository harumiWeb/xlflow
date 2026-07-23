import fs from "node:fs";
import path from "node:path";

const repo = path.resolve(".");
const check = process.argv.includes("--check");
const source = `${fs.readFileSync(path.join(repo, "internal/cli/root.go"), "utf8")}\n${fs.readFileSync(path.join(repo, "internal/cli/recovery.go"), "utf8")}`;
const commands = [
  ["capabilities", "capabilities"],
  ["new", "new"],
  ["init", "init"],
  ["pack", "pack"],
  ["doctor", "doctor"],
  ["attach", "attach"],
  ["backup", "backup"],
  ["list", "list"],
  ["form", "form"],
  ["formulas", "formulas"],
  ["pull", "pull"],
  ["push", "push"],
  ["rollback", "rollback"],
  ["status", "status"],
  ["session", "session"],
  ["save", "save"],
  ["recovery", "recovery"],
  ["runner", "runner"],
  ["run", "run"],
  ["export-image", "export-image"],
  ["edit", "edit"],
  ["macros", "macros"],
  ["ui", "ui"],
  ["test", "test"],
  ["type db", "type"],
  ["diff", "diff"],
  ["inspect", "inspect"],
  ["inspect-gui", "inspect-gui"],
  ["lint", "lint"],
  ["lsp", "lsp"],
  ["fmt", "fmt"],
  ["analyze", "analyze"],
  ["check", "check"],
  ["generate", "generate"],
  ["module", "module"],
  ["completion", "completion"],
  ["process", "process"],
  ["skill", "skill"],
  ["version", "version"],
  ["update", "update"],
];
const missing = commands.filter(
  ([, use]) => use !== "completion" && !new RegExp(`Use:\\s+"${use}(?:\\s|\\")`).test(source),
);
if (missing.length) {
  console.error(
    `CLI source is missing documented commands: ${missing.map(([name]) => name).join(", ")}`,
  );
  process.exitCode = 1;
}

const content = `# CLI command inventory\n\nGenerated from the Cobra command registrations in \`internal/cli/root.go\`. Run \`pnpm docs:generate-cli\` after adding a command.\n\n| Command | Documentation |\n| --- | --- |\n${commands.map(([name]) => `| \`xlflow ${name}\` | [command guide](../commands/${name.replace(" ", "-")}) |`).join("\n")}\n`;
const output = path.join(repo, "vitepress/reference/cli-command-inventory.md");
if (check) {
  if (!fs.existsSync(output) || fs.readFileSync(output, "utf8") !== content) {
    console.error("generated CLI inventory is stale: vitepress/reference/cli-command-inventory.md");
    process.exitCode = 1;
  }
} else {
  fs.writeFileSync(output, content);
}
