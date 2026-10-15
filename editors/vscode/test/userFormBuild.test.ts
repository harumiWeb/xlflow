import * as assert from "assert";
import {
  collectFormBuildSources,
  executeFormBuild,
  type FormBuildCommandResult,
  type FormBuildRunner,
  type FormBuildSource,
} from "../src/userFormBuildWorkflow";

const sources: FormBuildSource[] = [
  { name: "Alpha", specPath: "custom/forms/specs/Alpha.yaml" },
  { name: "Beta", specPath: "custom/forms/specs/Beta.json" },
  { name: "Gamma", specPath: "custom/forms/specs/Gamma.yml" },
];
const root = process.cwd();
const workbook = "build/Target.xlsm";
const ok: FormBuildCommandResult = { exitCode: 0, json: { status: "ok" } };
const active: FormBuildCommandResult = {
  exitCode: 0,
  json: { status: "ok", session: { active: true, workbook_path: workbook } },
};
const inactive: FormBuildCommandResult = {
  exitCode: 0,
  json: { status: "ok", session: { active: false, running: false, workbook_open: false } },
};
const failure: FormBuildCommandResult = {
  exitCode: 1,
  json: { status: "failed", error: { code: "form_failed", message: "focused failure" } },
};

function fakeRunner(results: FormBuildCommandResult[]): {
  run: FormBuildRunner;
  calls: string[][];
} {
  const calls: string[][] = [];
  return {
    calls,
    run: async (args) => {
      calls.push(args);
      const result = results.shift();
      assert.ok(result, `Unexpected command: ${args.join(" ")}`);
      return result;
    },
  };
}

async function run(): Promise<void> {
  assert.deepStrictEqual(
    collectFormBuildSources([
      "Zulu.yml",
      "Alpha.yaml",
      "Beta.json",
      "README.md",
      "sub/Nested.yaml",
      "\\Outside.yaml",
    ]),
    {
      sources: [
        { name: "Alpha", specPath: "Alpha.yaml" },
        { name: "Beta", specPath: "Beta.json" },
        { name: "Zulu", specPath: "Zulu.yml" },
      ],
      duplicates: [],
    },
  );
  assert.deepStrictEqual(
    collectFormBuildSources(["Alpha.yaml", "alpha.JSON", "ALPHA.yml", "Beta.json"]),
    {
      sources: [{ name: "Beta", specPath: "Beta.json" }],
      duplicates: [["Alpha.yaml", "alpha.JSON", "ALPHA.yml"]],
    },
  );
  {
    const fake = fakeRunner([inactive, active, ok, ok, ok, ok, ok]);
    const result = await executeFormBuild(sources, workbook, root, fake.run);
    assert.strictEqual(result.phase, "complete");
    assert.strictEqual(result.createdSession, true);
    assert.deepStrictEqual(result.succeeded, ["Alpha", "Beta", "Gamma"]);
    assert.deepStrictEqual(result.pending, []);
    assert.deepStrictEqual(fake.calls, [
      ["--json", "session", "status"],
      ["--json", "session", "start"],
      ...sources.map((source) => [
        "--json",
        "form",
        "build",
        source.specPath,
        "--overwrite",
        "--session",
        "--no-save",
      ]),
      ["--json", "save", "--session"],
      ["--json", "session", "stop"],
    ]);
  }
  {
    const fake = fakeRunner([active, ok, ok]);
    const result = await executeFormBuild(sources.slice(0, 1), workbook, root, fake.run);
    assert.strictEqual(result.phase, "complete");
    assert.strictEqual(result.createdSession, false);
    assert.deepStrictEqual(
      fake.calls.map((args) => args.slice(1, 3)),
      [
        ["session", "status"],
        ["form", "build"],
        ["save", "--session"],
      ],
    );
  }
  for (const existing of [active, inactive]) {
    const fake = fakeRunner(
      existing === active ? [active, ok, failure] : [inactive, active, ok, failure],
    );
    const result = await executeFormBuild(sources, workbook, root, fake.run);
    assert.strictEqual(result.phase, "build");
    assert.deepStrictEqual(result.succeeded, ["Alpha"]);
    assert.strictEqual(result.failed, "Beta");
    assert.deepStrictEqual(result.pending, ["Gamma"]);
    assert.strictEqual(result.error, "focused failure");
    assert.ok(!fake.calls.some((args) => args[1] === "save" || args[2] === "stop"));
  }
  {
    const fake = fakeRunner([inactive, failure]);
    const result = await executeFormBuild(sources, workbook, root, fake.run);
    assert.strictEqual(result.phase, "start");
    assert.deepStrictEqual(result.pending, ["Alpha", "Beta", "Gamma"]);
    assert.strictEqual(result.createdSession, false);
    assert.strictEqual(fake.calls.length, 2);
  }
  for (const result of [
    failure,
    { exitCode: 0 },
    { exitCode: 0, json: { status: "ok" } },
    { exitCode: 0, json: { status: "ok", session: { active: true, workbook_path: "Other.xlsm" } } },
    { exitCode: 0, json: { status: "ok", session: { active: false, running: true } } },
    { exitCode: 0, json: { status: "ok", session: { active: false, workbook_open: true } } },
    {
      exitCode: 0,
      json: { status: "ok", session: { active: false }, coordination: { recovery_required: true } },
    },
    {
      exitCode: 0,
      json: { status: "ok", session: { active: true }, coordination: { busy: true } },
    },
  ] satisfies FormBuildCommandResult[]) {
    const fake = fakeRunner([result]);
    const outcome = await executeFormBuild(sources, workbook, root, fake.run);
    assert.strictEqual(outcome.phase, "status");
    assert.ok(outcome.error);
    assert.strictEqual(fake.calls.length, 1);
  }
  for (const phase of ["save", "stop"] as const) {
    const fake = fakeRunner([inactive, active, ok, ...(phase === "stop" ? [ok] : []), failure]);
    const result = await executeFormBuild(sources.slice(0, 1), workbook, root, fake.run);
    assert.strictEqual(result.phase, phase);
    assert.deepStrictEqual(result.succeeded, ["Alpha"]);
    assert.strictEqual(result.failed, undefined);
    assert.deepStrictEqual(result.pending, []);
    assert.strictEqual(result.createdSession, true);
    if (phase === "save") assert.ok(!fake.calls.some((args) => args[2] === "stop"));
  }
  {
    const fake = fakeRunner([
      inactive,
      {
        exitCode: 0,
        json: { status: "ok", session: { active: true, workbook_path: "Other.xlsm" } },
      },
    ]);
    const result = await executeFormBuild(sources, workbook, root, fake.run);
    assert.strictEqual(result.phase, "start");
    assert.ok(result.error);
    assert.strictEqual(fake.calls.length, 2);
  }
  {
    const result = await executeFormBuild(sources, workbook, root, async () => {
      throw new Error("spawn failed");
    });
    assert.strictEqual(result.phase, "status");
    assert.match(result.error!, /spawn failed/);
  }
  {
    let concurrent = 0;
    let maximum = 0;
    const fake = fakeRunner([active, ok, ok, ok, ok]);
    await executeFormBuild(sources, workbook, root, async (args) => {
      concurrent++;
      maximum = Math.max(maximum, concurrent);
      await new Promise((resolve) => setTimeout(resolve, 1));
      const result = await fake.run(args);
      concurrent--;
      return result;
    });
    assert.strictEqual(maximum, 1, "No CLI command may overlap another");
  }
  console.log("UserForm build workflow tests passed");
}

void run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
