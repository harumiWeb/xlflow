package coordination

func buildDescriptors() []Descriptor {
	sourceRead := Policy{ResourceScope: ResourceNone, OperationKind: OperationRead, ParallelSafe: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryNotApplicable}
	sourceMutate := Policy{ResourceScope: ResourceNone, OperationKind: OperationMutate, ParallelSafe: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryNotApplicable}
	workbookRead := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationRead, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryBlock}
	workbookMutate := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationMutate, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryBlock}
	workbookExecute := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationExecute, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryBlock}
	workbookDesigner := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationDesigner, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryBlock}
	workbookObserver := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationRead, ParallelSafe: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryObserve}
	excelRead := Policy{ResourceScope: ResourceExcelInstance, OperationKind: OperationRead, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryNotApplicable}
	excelMutate := Policy{ResourceScope: ResourceExcelInstance, OperationKind: OperationMutate, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryNotApplicable}
	recoveryAction := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationMutate, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail, RecoveryBehavior: RecoveryRecover}

	return []Descriptor{
		cli("version", "version", sourceRead),
		cli("update", "update", sourceMutate),
		cli("update.check", "update check", sourceRead),
		both("macros", "macros", workbookRead, bridge("macros")),
		cli("backup.list", "backup list", sourceRead),
		cli("backup.prune", "backup prune", sourceMutate),
		cli("backup.delete", "backup delete", sourceMutate),
		cli("rollback", "rollback", workbookMutate),
		cli("form.migrate.sidecar", "form migrate sidecar", workbookDesigner),
		cli("formulas.pull", "formulas pull", workbookRead),
		cli("formulas.inspect", "formulas inspect", sourceRead),
		cli("form.new", "form new", sourceMutate),
		cli("form.snapshot", "form snapshot", workbookDesigner),
		both("form.build", "form build", workbookDesigner, bridgeArgs("form-write", "Action", "build")),
		both("form.apply", "form apply", workbookDesigner, bridgeArgs("form-write", "Action", "apply")),
		both("form.export-image", "form export-image", workbookDesigner, bridge("form-export-image")),
		both("list.forms", "list forms", workbookRead, bridgeArgs("list", "Action", "forms")),
		both("ui.button.add", "ui button add", workbookMutate, bridgeArgs("ui", "Action", "add")),
		both("ui.button.list", "ui button list", workbookRead, bridgeArgs("ui", "Action", "list")),
		both("ui.button.remove", "ui button remove", workbookMutate, bridgeArgs("ui", "Action", "remove")),
		both("new", "new", workbookMutate, bridge("new")),
		cli("init", "init", workbookRead),
		both("doctor", "doctor", excelRead, bridge("doctor")),
		both("attach", "attach", withRecovery(excelMutate, RecoveryBlock), bridge("attach")),
		both("pull", "pull", workbookRead, bridge("pull")),
		cli("pack", "pack", workbookMutate),
		both("push", "push", workbookMutate, bridge("push")),
		cli("generate.test", "generate test", sourceMutate),
		cli("module.new", "module new", sourceMutate),
		cli("module.remove", "module remove", sourceMutate),
		cli("module.rename", "module rename", sourceMutate),
		cli("module.install", "module install", workbookMutate),
		both("session.start", "session start", workbookMutate, bridgeArgs("session", "Action", "start")),
		both("session.status", "session status", workbookObserver, bridgeArgs("session", "Action", "status")),
		both("session.stop", "session stop", recoveryAction, bridgeArgs("session", "Action", "stop")),
		both("session.attach", "session attach", withRecovery(excelMutate, RecoveryBlock), bridgeArgs("session", "Action", "attach")),
		both("save", "save", workbookMutate, bridgeArgs("session", "Action", "save")),
		cli("status", "status", workbookObserver),
		both("runner.install", "runner install", workbookMutate, bridgeArgs("runner", "Action", "install")),
		both("runner.remove", "runner remove", workbookMutate, bridgeArgs("runner", "Action", "remove")),
		both("runner.status", "runner status", workbookRead, bridgeArgs("runner", "Action", "status")),
		both("run", "run", workbookExecute, bridge("run")),
		both("export-image", "export-image", workbookRead, bridge("export-image")),
		both("edit.sheet.add", "edit sheet add", workbookMutate, bridgeArgs("edit", "Action", "sheet_add")),
		both("edit.cell", "edit cell", workbookMutate, bridgeArgs("edit", "Action", "cell")),
		both("edit.formula", "edit formula", workbookMutate, bridgeArgs("edit", "Action", "formula")),
		both("edit.range", "edit range", workbookMutate, bridgeArgs("edit", "Action", "range")),
		both("edit.rows", "edit rows", workbookMutate, bridgeArgs("edit", "Action", "rows")),
		both("edit.columns", "edit columns", workbookMutate, bridgeArgs("edit", "Action", "columns")),
		both("test", "test", workbookExecute, bridge("test")),
		cli("test.list", "test list", sourceRead),
		cli("type.db.status", "type db status", sourceRead),
		both("type.db.init", "type db init", excelRead, bridge("type-db-import")),
		cli("type.db.refresh", "type db refresh", excelRead),
		cli("type.db.clean", "type db clean", sourceMutate),
		both("process.list", "process list", withRecovery(excelRead, RecoveryObserve), bridgeArgs("process", "Action", "list")),
		both("process.cleanup", "process cleanup", withRecovery(excelMutate, RecoveryRecover), bridgeArgs("process", "Action", "cleanup")),
		cli("recovery.clear", "recovery clear", recoveryAction),
		cli("diff", "diff", workbookRead),
		cli("inspect.calls", "inspect calls", sourceRead),
		cli("inspect.symbols", "inspect symbols", sourceRead),
		both("inspect.workbook", "inspect workbook", workbookRead, bridgeArgs("inspect", "Target", "workbook")),
		both("inspect.sheets", "inspect sheets", workbookRead, bridgeArgs("inspect", "Target", "sheets")),
		both("inspect.form", "inspect form", workbookDesigner, bridge("inspect-form")),
		both("inspect.range", "inspect range", workbookRead, bridgeArgs("inspect", "Target", "range")),
		both("inspect.used-range", "inspect used-range", workbookRead, bridgeArgs("inspect", "Target", "used-range")),
		both("inspect.cell", "inspect cell", workbookRead, bridgeArgs("inspect", "Target", "cell")),
		cli("fmt", "fmt", sourceMutate),
		cli("lint", "lint", sourceRead),
		cli("lsp", "lsp", sourceRead),
		cli("analyze", "analyze", sourceRead),
		cli("check", "check", excelRead),
		cli("inspect-gui", "inspect-gui", sourceRead),
		cli("skill.install", "skill install", sourceMutate),
	}
}

func withRecovery(policy Policy, behavior RecoveryBehavior) Policy {
	policy.RecoveryBehavior = behavior
	return policy
}

func cli(id CommandID, path string, policy Policy) Descriptor {
	return Descriptor{ID: id, Policy: policy, CLI: []CLISelector{{Path: path}}}
}

func both(id CommandID, path string, policy Policy, selector BridgeSelector) Descriptor {
	descriptor := cli(id, path, policy)
	descriptor.Bridge = []BridgeSelector{selector}
	return descriptor
}

func bridge(command string) BridgeSelector { return BridgeSelector{Command: command} }

func bridgeArgs(command, key, value string) BridgeSelector {
	return BridgeSelector{Command: command, Args: map[string]string{key: value}}
}
