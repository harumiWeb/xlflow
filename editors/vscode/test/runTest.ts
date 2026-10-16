import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import { createServer } from "net";
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
  const cdpPort = await availableTcpPort();
  process.env.XLFLOW_VSCODE_TEST_CDP_PORT = String(cdpPort);

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
      launchArgs: [
        workspaceRoot,
        "--user-data-dir",
        path.join(temporaryRoot, "user-data"),
        `--remote-debugging-port=${cdpPort}`,
        "--disable-workspace-trust",
        "--disable-extensions",
      ],
    });
  } finally {
    await cleanupTemporaryRoot(temporaryRoot, workspaceRoot);
  }
}

async function availableTcpPort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Could not reserve a CDP port");
  const { port } = address;
  await new Promise<void>((resolve, reject) => {
    server.close((error) => (error ? reject(error) : resolve()));
  });
  return port;
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
