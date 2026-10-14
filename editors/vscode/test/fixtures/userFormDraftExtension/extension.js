const vscode = require("vscode");

const viewType = "xlflow.userFormDraftProviderTest";

function activate() {
  return {
    register(provider) {
      return vscode.window.registerCustomEditorProvider(viewType, provider, {
        supportsMultipleEditorsPerDocument: true,
      });
    },
    registerSaveCommand(provider, commandId) {
      return provider.registerSaveCommand(commandId);
    },
  };
}

module.exports = { activate };
