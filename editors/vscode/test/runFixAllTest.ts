import { mkdtemp, mkdir, rm, access, writeFile } from "fs/promises";
import { runTests } from "@vscode/test-electron";
import * as os from "os";
import * as path from "path";

const debugExecutableEnvironmentVariable = "XLFLOW_VSCODE_DEBUG_PATH";

async function main(): Promise<void> {
  const extensionDevelopmentPath = path.resolve(__dirname, "..", "..");
  const extensionTestsPath = path.resolve(__dirname, "suite", "fixAll.test.js");
  const goExecutable = path.resolve(
    process.env.XLFLOW_VSCODE_DEBUG_PATH ??
      path.resolve(__dirname, "..", "..", "..", "..", ".tmp-quickfix", "xlflow.exe"),
  );
  await access(goExecutable);
  process.env[debugExecutableEnvironmentVariable] = goExecutable;

  const temporaryRoot = await mkdtemp(path.join(os.tmpdir(), "xlflow-fix-all-"));
  const workspaceRoot = path.join(temporaryRoot, "workspace");

  try {
    await mkdir(workspaceRoot, { recursive: true });
    await writeFile(
      path.join(workspaceRoot, "xlflow.toml"),
      '[project]\nname = "vscode-fix-all-test"\n\n[lint]\ndetect_let_assignment = true\n',
      "utf8",
    );

    const options: Parameters<typeof runTests>[0] = {
      extensionDevelopmentPath,
      extensionTestsPath,
      launchArgs: [
        workspaceRoot,
        "--user-data-dir",
        path.join(temporaryRoot, "user-data"),
        "--disable-workspace-trust",
        "--disable-extensions",
      ],
    };
    const installedVSCode = await installedVSCodeExecutable();
    if (installedVSCode !== undefined) {
      options.vscodeExecutablePath = installedVSCode;
    }

    console.log(`[fix-all] CLI LSP executable: ${goExecutable}`);
    console.log(`[fix-all] VS Code executable: ${installedVSCode ?? "test-electron managed"}`);
    await runTests(options);
  } finally {
    await rm(confirmedTemporaryRoot(temporaryRoot, workspaceRoot), {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    });
  }
}

async function installedVSCodeExecutable(): Promise<string | undefined> {
  const configured = process.env.VSCODE_EXECUTABLE?.trim();
  if (configured) {
    try {
      await access(configured);
      return path.resolve(configured);
    } catch {
      return undefined;
    }
  }
  const localAppData = process.env.LOCALAPPDATA;
  if (localAppData === undefined) return undefined;
  const candidate = path.join(localAppData, "Programs", "Microsoft VS Code", "Code.exe");
  try {
    await access(candidate);
    return candidate;
  } catch {
    return undefined;
  }
}

function confirmedTemporaryRoot(temporaryRoot: string, workspaceRoot: string): string {
  const temporaryBase = path.resolve(os.tmpdir());
  const resolvedTemporaryRoot = path.resolve(temporaryRoot);
  const relativeTemporaryRoot = path.relative(temporaryBase, resolvedTemporaryRoot);
  if (
    !path.basename(resolvedTemporaryRoot).startsWith("xlflow-fix-all-") ||
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
