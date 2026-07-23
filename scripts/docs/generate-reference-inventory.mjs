import fs from "node:fs";
import path from "node:path";

const repo = path.resolve(".");
const check = process.argv.includes("--check");
const sourceFiles = [];
function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "vendor" || entry.name === ".git") continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full);
    else if (entry.name.endsWith(".go")) sourceFiles.push(full);
  }
}
walk(path.join(repo, "internal"));

const source = sourceFiles.map((file) => fs.readFileSync(file, "utf8")).join("\n");
const diagnostics = [
  ...new Set([...source.matchAll(/(?:ID|Code):\s*"([A-Z]{2,5}\d{3})"/g)].map((m) => m[1])),
].sort();
const errors = [
  ...new Set([...source.matchAll(/"([a-z][a-z0-9]*(?:_[a-z0-9]+)+)"/g)].map((m) => m[1])),
]
  .filter((code) => code.length < 80 && !code.startsWith("go_") && !code.startsWith("http_"))
  .sort();

const diagnosticPage = `# Diagnostic rule inventory\n\nGenerated from diagnostic IDs in \`internal/\`. Run \`pnpm docs:generate-reference\` after adding a rule.\n\n${diagnostics.map((code) => `- \`${code}\``).join("\n")}\n`;
const errorPage = `# Error-code inventory\n\nGenerated from structured error-code literals in \`internal/\`. Descriptions and recovery guidance remain curated in [Error Codes](./error-codes) and [Troubleshooting](../help/troubleshooting).\n\n${errors.map((code) => `- \`${code}\``).join("\n")}\n`;

const outputs = new Map([
  [path.join(repo, "vitepress/reference/diagnostics.md"), diagnosticPage],
  [path.join(repo, "vitepress/reference/error-code-inventory.md"), errorPage],
]);
let failed = false;
for (const [file, content] of outputs) {
  if (check) {
    if (!fs.existsSync(file) || fs.readFileSync(file, "utf8") !== content) {
      console.error(`generated reference is stale: ${path.relative(repo, file)}`);
      failed = true;
    }
  } else {
    fs.writeFileSync(file, content);
  }
}
if (failed) process.exitCode = 1;
