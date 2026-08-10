package analyze

// DataFlowEndpoint identifies the source or sink occurrence represented by a
// conservative source-to-sink finding.
type DataFlowEndpoint struct {
	Kind  string `json:"kind"`
	Label string `json:"label"`
	Line  int    `json:"line"`
}

// DataFlowStep is one deterministic propagation step between a source and a
// sink. Labels retain the normalized source expression without exposing parser
// or protocol-owned values.
type DataFlowStep struct {
	Kind  string `json:"kind"`
	Label string `json:"label"`
	Line  int    `json:"line"`
}

// DataFlowContext is the additive JSON context attached to source-to-sink
// findings (VBA224 and the command-launch projection VBA236). The human and
// LSP projections also summarize this context in their message and reason
// fields.
type DataFlowContext struct {
	Source DataFlowEndpoint `json:"source"`
	Sink   DataFlowEndpoint `json:"sink"`
	Path   []DataFlowStep   `json:"path,omitempty"`
}

// CommandExecutionContext describes the role and risk classification of a
// process-launch sink. It is deliberately additive to DataFlowContext so
// clients that only understand generic source-to-sink flows can continue to
// consume diagnostics emitted by older analyzer versions.
//
// risk_class is either "injection" (the command text/executable is known to
// contain tainted input) or "process_launch" (the launch is risky but the
// input origin or command role is not sufficiently precise to claim
// injection). risk_kind carries the more specific policy observation such as
// unquoted_executable_path or observability.
type CommandExecutionContext struct {
	RiskClass   string `json:"risk_class"`
	RiskKind    string `json:"risk_kind"`
	Launcher    string `json:"launcher,omitempty"`
	Interpreter string `json:"interpreter,omitempty"`
	CommandRole string `json:"command_role,omitempty"`
	OriginState string `json:"origin_state,omitempty"`
}

// SQLExecutionContext describes the SQL API, input role, and conservative risk
// classification attached to a VBA239 finding.
type SQLExecutionContext struct {
	RiskKind      string `json:"risk_kind"`
	API           string `json:"api,omitempty"`
	InputRole     string `json:"input_role,omitempty"`
	OriginState   string `json:"origin_state,omitempty"`
	Parameterized bool   `json:"parameterized"`
}
