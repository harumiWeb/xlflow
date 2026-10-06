const esbuild = require("esbuild");

const production = process.argv.includes("--production");

const common = {
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node20",
  external: ["vscode"],
  sourcemap: !production,
  minify: production,
};

Promise.all([
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormRendering.test.tsx"],
    outfile: "dist/test/userFormRendering.test.js",
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormProperties.test.tsx"],
    outfile: "dist/test/userFormProperties.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    bundle: true,
    platform: "browser",
    format: "iife",
    target: "es2022",
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["webview/userFormDesigner/App.tsx"],
    outfile: "dist/userFormDesigner/app.js",
    sourcemap: !production,
    minify: production,
  }),
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
