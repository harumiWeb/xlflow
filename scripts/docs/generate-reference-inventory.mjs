import fs from "node:fs";
import path from "node:path";

const repo = path.resolve(".");
const check = process.argv.includes("--check");
const registryPath = path.join(repo, "internal/staticanalysis/rules/registry.json");

if (!fs.existsSync(registryPath)) {
  console.error(
    "canonical diagnostic registry is missing: internal/staticanalysis/rules/registry.json",
  );
  process.exit(1);
}

const registryDocument = JSON.parse(fs.readFileSync(registryPath, "utf8"));
if (
  typeof registryDocument !== "object" ||
  registryDocument === null ||
  Array.isArray(registryDocument) ||
  registryDocument.schema_version !== 1 ||
  !Array.isArray(registryDocument.items)
) {
  throw new Error("diagnostic registry must use schema_version 1 with an items array");
}
const registryRules = registryDocument.items;

const requiredRuleFields = [
  "id",
  "title",
  "description",
  "family",
  "category",
  "default_severity",
  "default_enabled",
  "scope",
  "realtime",
  "precision",
  "fix_available",
  "documentation_url",
  "configurable",
  "configuration_key",
  "inline_suppressible",
  "preflight_blocking",
];
const booleanRuleFields = new Set([
  "default_enabled",
  "realtime",
  "fix_available",
  "configurable",
  "inline_suppressible",
  "preflight_blocking",
]);
for (const [index, rule] of registryRules.entries()) {
  if (typeof rule !== "object" || rule === null || Array.isArray(rule)) {
    throw new Error(`diagnostic registry item ${index} must be an object`);
  }
  const missing = requiredRuleFields.filter((field) => !(field in rule));
  if (missing.length > 0) {
    throw new Error(`diagnostic ${rule.id ?? "<unknown>"} is missing: ${missing.join(", ")}`);
  }
  for (const field of requiredRuleFields) {
    const expected = booleanRuleFields.has(field) ? "boolean" : "string";
    if (typeof rule[field] !== expected) {
      const diagnostic = typeof rule.id === "string" ? rule.id : `<item ${index}>`;
      throw new Error(`diagnostic ${diagnostic} has invalid ${field}: expected ${expected}`);
    }
  }
}
const rules = registryRules.map((rule) => ({ ...rule })).sort((a, b) => a.id.localeCompare(b.id));
const duplicateIDs = rules
  .filter((rule, index) => index > 0 && rule.id === rules[index - 1].id)
  .map((rule) => rule.id);
if (duplicateIDs.length > 0) {
  throw new Error(`diagnostic registry contains duplicate IDs: ${duplicateIDs.join(", ")}`);
}

const sourceFiles = [];
function walk(dir) {
  if (path.resolve(dir) === path.resolve(repo, "internal/staticanalysis/rules")) return;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "vendor" || entry.name === ".git") continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full);
    else if (entry.name.endsWith(".go")) sourceFiles.push(full);
  }
}
walk(path.join(repo, "internal"));

const source = sourceFiles.map((file) => fs.readFileSync(file, "utf8")).join("\n");
// Some structured metadata keys use the same snake_case shape as error-code
// literals but are not user-facing errors and do not belong in this inventory.
const excludedErrorInventoryLiterals = new Set([
  "binding_status",
  "rule_codes",
  "binding_note",
]);
const errors = [
  ...new Set([...source.matchAll(/"([a-z][a-z0-9]*(?:_[a-z0-9]+)+)"/g)].map((m) => m[1])),
]
  .filter(
    (code) =>
      code.length < 80 &&
      !code.startsWith("go_") &&
      !code.startsWith("http_") &&
      code !== "default_enabled" &&
      code !== "default_severity" &&
      !excludedErrorInventoryLiterals.has(code),
  )
  .sort();

const markdown = (value) =>
  String(value ?? "")
    .replaceAll("|", "\\|")
    .replaceAll("\n", " ");
const yesNo = (value) => (value ? "yes" : "no");
const setting = (rule) =>
  rule.configurable ? `\`${markdown(rule.configuration_key)}\`` : "not configurable";
const formatTable = (rows) => {
  const widths = rows[0].map((_, column) => Math.max(...rows.map((row) => row[column].length)));
  const formatRow = (row) =>
    `| ${row.map((cell, column) => cell.padEnd(widths[column])).join(" | ")} |`;
  return [
    formatRow(rows[0]),
    formatRow(widths.map((width) => "-".repeat(width))),
    ...rows.slice(1).map(formatRow),
  ].join("\n");
};

const summaryRows = formatTable([
  ["ID", "Family", "Severity", "Scope", "Default", "Title"],
  ...rules.map((rule) => [
    `[\`${rule.id}\`](#${rule.id.toLowerCase()})`,
    markdown(rule.family),
    markdown(rule.default_severity),
    markdown(rule.scope),
    yesNo(rule.default_enabled),
    markdown(rule.title),
  ]),
]);
const details = rules
  .map((rule) => {
    const properties = formatTable([
      ["Property", "Value"],
      ["Family", `\`${markdown(rule.family)}\``],
      ["Category", `\`${markdown(rule.category)}\``],
      ["Default severity", `\`${markdown(rule.default_severity)}\``],
      ["Scope", `\`${markdown(rule.scope)}\``],
      ["Precision", `\`${markdown(rule.precision)}\``],
      ["Enabled by default", yesNo(rule.default_enabled)],
      ["Configuration", setting(rule)],
      ["Inline suppression", yesNo(rule.inline_suppressible)],
      ["Blocks source preflight", yesNo(rule.preflight_blocking)],
      ["Real-time editor diagnostic", yesNo(rule.realtime)],
      ["Fix available", yesNo(rule.fix_available)],
    ]);
    return `## ${rule.id}

**${markdown(rule.title)}.** ${markdown(rule.description)}

${properties}
`;
  })
  .join("\n");

const diagnosticPage = `# Static-analysis diagnostic catalog

Generated from the canonical rule registry at \`internal/staticanalysis/rules/registry.json\`. Run \`pnpm docs:generate-reference\` after changing rule metadata. Do not edit this page by hand.

Use [\`xlflow rules\`](../commands/rules) to inspect the same metadata from an installed xlflow binary. \`VBA000\` is a synthetic analysis-failure diagnostic and is intentionally outside the registry; UserForm \`FRM...\` and \`UFY...\` diagnostics are outside this catalog.

${summaryRows}

${details}`;
const errorPage = `# Error-code inventory

Generated from structured error-code literals in \`internal/\`. Descriptions and recovery guidance remain curated in [Error Codes](./error-codes) and [Troubleshooting](../help/troubleshooting).

${errors.map((code) => `- \`${code}\``).join("\n")}
`;

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
