const { copyFileSync } = require("node:fs");
const path = require("node:path");
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
  loader: { ".css": "empty" },
};

Promise.all([
  esbuild.build({
    ...common,
    entryPoints: ["test/userFormBuild.test.ts"],
    outfile: "dist/test/userFormBuild.test.js",
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormPropertyEditors.test.tsx"],
    outfile: "dist/test/userFormPropertyEditors.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    entryPoints: ["test/userFormPictures.test.ts"],
    outfile: "dist/test/userFormPictures.test.js",
  }),
  esbuild.build({
    ...common,
    entryPoints: ["test/userFormPictureReferences.test.ts"],
    outfile: "dist/test/userFormPictureReferences.test.js",
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormDesignerUX.test.tsx"],
    outfile: "dist/test/userFormDesignerUX.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormDesignerIcon.test.tsx"],
    outfile: "dist/test/userFormDesignerIcon.test.js",
  }),
  esbuild.build({
    ...common,
    entryPoints: ["test/userFormDraft.test.ts"],
    outfile: "dist/test/userFormDraft.test.js",
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormCollections.test.tsx"],
    outfile: "dist/test/userFormCollections.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormProductivity.test.tsx"],
    outfile: "dist/test/userFormProductivity.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormParenting.test.tsx"],
    outfile: "dist/test/userFormParenting.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormStructure.test.tsx"],
    outfile: "dist/test/userFormStructure.test.js",
    external: [...common.external, "jsdom"],
  }),
  esbuild.build({
    ...common,
    jsx: "automatic",
    jsxImportSource: "preact",
    entryPoints: ["test/userFormRendering.test.tsx"],
    outfile: "dist/test/userFormRendering.test.js",
    external: [...common.external, "jsdom"],
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
    loader: { ".ttf": "file" },
    assetNames: "[name]",
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
])
  .then(() => {
    copyFileSync(
      require.resolve("@vscode/codicons/LICENSE"),
      path.join(__dirname, "dist", "userFormDesigner", "CODICONS-LICENSE.txt"),
    );
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
