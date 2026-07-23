import fs from "node:fs";
import path from "node:path";

const root = path.resolve("vitepress/commands");
const marker = "<!-- xlflow-command-guidance -->";
const pages = fs.readdirSync(root).filter((name) => name.endsWith(".md") && name !== "index.md");

for (const page of pages) {
  const file = path.join(root, page);
  let source = fs.readFileSync(file, "utf8");
  if (source.includes(marker)) continue;

  const command = path.basename(page, ".md").replaceAll("-", " ");
  source += `\n\n${marker}\n## When to use this command\n\nUse \`xlflow ${command}\` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.\n\n## Prerequisites\n\nCheck the project configuration and run \`xlflow doctor --json\` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.\n\n## What this command reads and changes\n\nThe command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add \`--session\` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.\n\n## Effect on source-of-truth state\n\nUse \`xlflow status --json\` before and after the command. A source edit normally requires \`push\`; a workbook edit normally requires \`pull\`; a dirty live session requires \`save --session\` or an intentional discard.\n\n## Common workflows\n\nCombine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use \`--json\` in scripts and agent loops.\n\n## Common failures\n\nRead the structured \`error.code\`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.\n`;
  fs.writeFileSync(file, source);
}

console.log(`normalized ${pages.length} command pages`);
