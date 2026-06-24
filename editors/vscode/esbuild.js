const esbuild = require("esbuild");

const production = process.argv.includes("--production");

const common = {
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node18",
  external: ["vscode"],
  sourcemap: !production,
  minify: production,
};

Promise.all([
  esbuild.build({
    ...common,
    entryPoints: ["src/extension.ts"],
    outfile: "dist/extension.js",
  }),
  esbuild.build({
    ...common,
    entryPoints: ["test/runTest.ts"],
    outfile: "dist/test/runTest.js",
  }),
  esbuild.build({
    ...common,
    entryPoints: ["test/suite/extension.test.ts"],
    outfile: "dist/test/suite/extension.test.js",
  }),
]).catch(() => process.exit(1));
