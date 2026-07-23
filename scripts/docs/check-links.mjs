import fs from "node:fs";
import path from "node:path";

const repo = path.resolve(".");
const docs = path.join(repo, "vitepress");
const markdown = [];
function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === ".vitepress" || entry.name === "public") continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full);
    else if (entry.name.endsWith(".md")) markdown.push(full);
  }
}
walk(docs);

function candidates(from, target) {
  const clean = target.split("#", 1)[0].split("?", 1)[0];
  if (!clean) return [];
  if (clean.startsWith("/")) {
    const relative = clean.replace(/^\//, "");
    return [
      path.join(docs, `${relative}.md`),
      path.join(docs, relative, "index.md"),
      path.join(docs, "public", relative),
    ];
  }
  const resolved = path.resolve(path.dirname(from), clean);
  return [resolved, `${resolved}.md`, path.join(resolved, "index.md")];
}

const failures = [];
for (const file of markdown) {
  const source = fs.readFileSync(file, "utf8");
  for (const match of source.matchAll(/!?\[[^\]]*\]\(([^)\s]+)(?:\s+[^)]*)?\)/g)) {
    const target = match[1].replace(/^<|>$/g, "");
    if (/^(?:https?:|mailto:|tel:|data:|#)/i.test(target)) continue;
    if (target.includes("${") || target.includes("<path>")) continue;
    if (!candidates(file, target).some((candidate) => fs.existsSync(candidate))) {
      failures.push(`${path.relative(repo, file)} -> ${target}`);
    }
  }
}

if (failures.length) {
  console.error(failures.join("\n"));
  process.exitCode = 1;
} else {
  console.log(`markdown links OK (${markdown.length} files checked)`);
}
