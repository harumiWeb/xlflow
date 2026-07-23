import fs from "node:fs";
import path from "node:path";

const root = path.resolve("vitepress/demos");
const marker = "<!-- xlflow-demo-case-study -->";
const featureNotes = {
  "world-news": "HTTP data retrieval, source-first modules, and structured workbook inspection.",
  "stock-dashboard": "Formula snapshots, worksheet inspection, and visual output verification.",
  "qr-code": "External API integration, image output, and repeatable macro execution.",
  tetris: "UserForm controls, event procedures, session-backed visual iteration, and image export.",
  "invader-game": "UserForm event handlers, timer cleanup, and interactive visual verification.",
  "calendar-picker":
    "UserForm Designer artifacts, sidecar code-behind, and controlled form snapshots.",
};

for (const page of fs
  .readdirSync(root)
  .filter((name) => name.endsWith(".md") && name !== "index.md")) {
  const file = path.join(root, page);
  let source = fs.readFileSync(file, "utf8");
  if (source.includes(marker)) continue;
  const slug = path.basename(page, ".md");
  const title = slug.replaceAll("-", " ").replace(/\b\w/g, (c) => c.toUpperCase());
  const note =
    featureNotes[slug] ?? "Source-controlled VBA, workbook execution, and observable verification.";
  source += `\n\n${marker}\n## What it does\n\n${title} is a small workbook application that can be inspected, executed, and reviewed as source.\n\n## Why it is a useful xlflow example\n\nIt demonstrates ${note}\n\n## Project structure\n\nThe repository keeps VBA under \`src/\`, the workbook under \`build/\`, and project behavior in \`xlflow.toml\`.\n\n## xlflow features used\n\n- \`doctor\`, \`status\`, and \`pull\` for setup and source synchronization;\n- \`fmt\`, \`lint\`, and \`analyze\` before Excel operations;\n- sessions, \`run --diagnostic\`, \`inspect\`, and \`export-image\` for verification.\n\n## Verification strategy\n\nRun the source checks, push into a disposable workbook, execute the documented entry point, inspect the affected cells or form, export an image when layout matters, and review the Git diff.\n\n## Commands to reproduce\n\n\`\`\`bash\nxlflow doctor --json\nxlflow pull --json\nxlflow lint --json\nxlflow push --json\nxlflow run --diagnostic --json\nxlflow inspect workbook --json\nxlflow export-image --json\n\`\`\`\n\n## Repository\n\nSee the linked example repository above for the workbook and source files.\n`;
  fs.writeFileSync(file, source);
}
