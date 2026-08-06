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

// DataFlowContext is the additive JSON context attached to VBA224 findings.
// The human and LSP projections also summarize this context in their message
// and reason fields.
type DataFlowContext struct {
	Source DataFlowEndpoint `json:"source"`
	Sink   DataFlowEndpoint `json:"sink"`
	Path   []DataFlowStep   `json:"path,omitempty"`
}
