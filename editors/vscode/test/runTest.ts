import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import * as os from "os";
import * as path from "path";
import { runTests } from "@vscode/test-electron";

async function main(): Promise<void> {
  const extensionDevelopmentPath = path.resolve(__dirname, "..", "..");
  const fixtureExtensionPath = path.resolve(
    extensionDevelopmentPath,
    "test",
    "fixtures",
    "userFormDraftExtension",
  );
  const extensionTestsPath = path.resolve(__dirname, "suite", "extension.test.js");
  const temporaryRoot = await mkdtemp(path.join(os.tmpdir(), "xlflow-vscode-922-"));
  const workspaceRoot = path.join(temporaryRoot, "workspace");

  try {
    await mkdir(path.join(workspaceRoot, "src", "forms", "specs"), { recursive: true });
    await writeFile(
      path.join(workspaceRoot, "xlflow.toml"),
      '[project]\nname = "vscode-clipboard-test"\n\n[src]\nforms = "src/forms"\n',
      "utf8",
    );
    await runTests({
      extensionDevelopmentPath: [extensionDevelopmentPath, fixtureExtensionPath],
      extensionTestsPath,
      launchArgs: [workspaceRoot, "--disable-workspace-trust", "--disable-extensions"],
    });
  } finally {
    await cleanupTemporaryRoot(temporaryRoot, workspaceRoot);
  }
}

async function cleanupTemporaryRoot(temporaryRoot: string, workspaceRoot: string): Promise<void> {
  const cleanupRoot = confirmedTemporaryRoot(temporaryRoot, workspaceRoot);
  await rm(cleanupRoot, {
    recursive: true,
    force: true,
    maxRetries: 5,
    retryDelay: 100,
  });
}

function confirmedTemporaryRoot(temporaryRoot: string, workspaceRoot: string): string {
  const temporaryBase = path.resolve(os.tmpdir());
  const resolvedTemporaryRoot = path.resolve(temporaryRoot);
  const relativeTemporaryRoot = path.relative(temporaryBase, resolvedTemporaryRoot);
  if (
    !path.basename(resolvedTemporaryRoot).startsWith("xlflow-vscode-922-") ||
    !relativeTemporaryRoot ||
    relativeTemporaryRoot === ".." ||
    relativeTemporaryRoot.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relativeTemporaryRoot) ||
    path.dirname(path.resolve(workspaceRoot)) !== resolvedTemporaryRoot
  ) {
    throw new Error(`Refusing to remove an unexpected VS Code test directory: ${temporaryRoot}`);
  }
  return resolvedTemporaryRoot;
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
