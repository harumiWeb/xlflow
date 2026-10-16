import * as assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { renderToString } from "preact-render-to-string";
import { Icon } from "../webview/userFormDesigner/Icon";

const cssPath = resolve(process.cwd(), "node_modules/@vscode/codicons/dist/codicon.css");
const codiconCss = readFileSync(cssPath, "utf8");

for (const name of ["undo", "redo"]) {
  const markup = renderToString(<Icon name={name} />);
  const codiconName = markup.match(/class="codicon codicon-([^"]+)"/)?.[1];

  assert.ok(codiconName, `${name} renders a Codicon class`);
  assert.ok(
    new RegExp(`\\.codicon-${codiconName}:before\\s*\\{\\s*content:`).test(codiconCss),
    `${name} resolves to a glyph shipped in Codicon CSS`,
  );
}
