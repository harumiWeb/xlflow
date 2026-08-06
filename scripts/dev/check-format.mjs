import { execFileSync, spawnSync } from "node:child_process";
import path from "node:path";

const root = process.cwd();
const windows = process.platform === "win32";

function stagedFiles() {
  const output = execFileSync("git", ["diff", "--cached", "--name-only", "--diff-filter=ACMR"], {
    cwd: root,
    encoding: "utf8",
  });
  return output
    .split(/\r?\n/)
    .map((file) => file.trim())
    .filter(Boolean)
    .map((file) => file.replaceAll(path.sep, "/"));
}

function run(label, command, args, { outputIsFailure = false } = {}) {
  const result = spawnSync(command, args, {
    cwd: root,
    encoding: "utf8",
    shell: windows,
    windowsHide: true,
  });
  if (result.error) {
    console.error(`${label}: ${result.error.message}`);
    return false;
  }
  const output = `${result.stdout ?? ""}${result.stderr ?? ""}`.trim();
  if (result.status !== 0 || (outputIsFailure && output !== "")) {
    console.error(`${label} failed${output ? `:\n${output}` : "."}`);
    return false;
  }
  return true;
}

const files = stagedFiles();
if (files.length === 0) {
  process.exit(0);
}

let ok = run("git diff --cached --check", "git", ["diff", "--cached", "--check"]);

const goFiles = files.filter((file) => file.toLowerCase().endsWith(".go"));
if (goFiles.length > 0) {
  ok = run("gofmt", "gofmt", ["-l", ...goFiles], { outputIsFailure: true }) && ok;
  ok = run("goimports", "goimports", ["-l", ...goFiles], { outputIsFailure: true }) && ok;
}

const oxfmtExtensions = new Set([".js", ".jsx", ".ts", ".tsx", ".json", ".jsonc", ".md", ".mdx"]);
const oxfmtFiles = files.filter((file) => oxfmtExtensions.has(path.extname(file).toLowerCase()));
if (oxfmtFiles.length > 0) {
  ok = run(
    "oxfmt",
    windows ? "pnpm.cmd" : "pnpm",
    ["exec", "oxfmt", "--check", ...oxfmtFiles],
  ) && ok;
}

const csharpFiles = files.filter((file) => file.toLowerCase().endsWith(".cs"));
if (csharpFiles.length > 0) {
  ok = run("dotnet format", "dotnet", [
    "format",
    "bridge/dotnet/Xlflow.ExcelBridge.sln",
    "--verify-no-changes",
    "--include",
    ...csharpFiles,
  ]) && ok;
}

if (!ok) {
  console.error("Run `task fmt` explicitly to apply repository-wide formatting, then review the diff.");
  process.exit(1);
}
