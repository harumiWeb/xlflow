const vscode = require("vscode");

const viewType = "xlflow.userFormDraftProviderTest";

function activate() {
  return {
    register(provider) {
      return vscode.window.registerCustomEditorProvider(viewType, provider, {
        supportsMultipleEditorsPerDocument: true,
      });
    },
  };
}

module.exports = { activate };
