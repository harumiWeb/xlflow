// Package dataflow provides a conservative, procedure-local source-to-sink
// analysis over normalized VBA IR and its control-flow graph.
package dataflow

import (
	"fmt"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// State is the abstract value state used by the analysis. Unknown is kept
// separate from Tainted so callers can distinguish an explicit source from a
// value that crossed an unsupported transformation or recovered syntax.
type State string

const (
	StateClean   State = "clean"
	StateTainted State = "tainted"
	StateUnknown State = "unknown"
)

// SourceKind identifies the initial untrusted-value catalogs.
type SourceKind string

const (
	SourceParameter      SourceKind = "parameter"
	SourceWorksheetCell  SourceKind = "worksheet_cell"
	SourceInputBox       SourceKind = "inputbox"
	SourceFileInput      SourceKind = "file_input"
	SourceEnvironment    SourceKind = "environment_variable"
	SourceCommandLine    SourceKind = "command_line_argument"
	SourceHTTPResponse   SourceKind = "http_response"
	SourceDatabaseResult SourceKind = "database_result"
	SourceUnknown        SourceKind = "unknown"
)

// SinkKind identifies a sensitive API catalog entry.
type SinkKind string

const (
	SinkShell            SinkKind = "shell"
	SinkWScriptShellRun  SinkKind = "wscript_shell_run"
	SinkWScriptShellExec SinkKind = "wscript_shell_exec"
	SinkSQLExecution     SinkKind = "sql_execution"
	SinkDestructiveFile  SinkKind = "destructive_file_operation"
	SinkWorkbooksOpen    SinkKind = "workbooks_open"
	SinkSaveAs           SinkKind = "save_as"
	SinkHTTPURL          SinkKind = "http_request_url"
	SinkHTTPHeader       SinkKind = "http_request_header"
)

// SourceSpec and SinkSpec are the public, protocol-neutral catalog entries.
type SourceSpec struct {
	Kind        SourceKind `json:"kind"`
	Label       string     `json:"label"`
	Description string     `json:"description"`
}

type SinkSpec struct {
	Kind        SinkKind `json:"kind"`
	Label       string   `json:"label"`
	Description string   `json:"description"`
}

// SourceCatalog returns a defensive copy of the initial source catalog.
func SourceCatalog() []SourceSpec {
	return append([]SourceSpec(nil), sourceCatalog...)
}

// SinkCatalog returns a defensive copy of the initial sink catalog.
func SinkCatalog() []SinkSpec {
	return append([]SinkSpec(nil), sinkCatalog...)
}

var sourceCatalog = []SourceSpec{
	{Kind: SourceParameter, Label: "procedure parameter", Description: "A value supplied by a procedure caller."},
	{Kind: SourceWorksheetCell, Label: "worksheet cell value", Description: "A value read from a worksheet cell or range."},
	{Kind: SourceInputBox, Label: "InputBox", Description: "Text supplied through the VBA InputBox function."},
	{Kind: SourceFileInput, Label: "text/CSV file input", Description: "Text read from a file or CSV stream."},
	{Kind: SourceEnvironment, Label: "environment variable", Description: "A process environment value."},
	{Kind: SourceCommandLine, Label: "command-line argument", Description: "The process command line or script arguments."},
	{Kind: SourceHTTPResponse, Label: "HTTP response", Description: "Text or bytes returned by an HTTP client."},
	{Kind: SourceDatabaseResult, Label: "database result", Description: "A value read from a database result or recordset."},
}

var sinkCatalog = []SinkSpec{
	{Kind: SinkShell, Label: "Shell", Description: "Starts a process through the VBA Shell function."},
	{Kind: SinkWScriptShellRun, Label: "WScript.Shell.Run", Description: "Runs a process through WScript.Shell."},
	{Kind: SinkWScriptShellExec, Label: "WScript.Shell.Exec", Description: "Executes a process through WScript.Shell."},
	{Kind: SinkSQLExecution, Label: "SQL execution", Description: "Executes SQL through a recognized database object."},
	{Kind: SinkDestructiveFile, Label: "destructive file operation", Description: "Deletes or removes a file or directory."},
	{Kind: SinkWorkbooksOpen, Label: "Workbooks.Open", Description: "Opens a workbook from a path or URL."},
	{Kind: SinkSaveAs, Label: "SaveAs", Description: "Writes a workbook to a caller-controlled path."},
	{Kind: SinkHTTPURL, Label: "HTTP request URL", Description: "Uses a value as an HTTP request URL."},
	{Kind: SinkHTTPHeader, Label: "HTTP request header", Description: "Uses a value as an HTTP request header name or value."},
}

// Options controls analysis behavior. The zero value enables the conservative
// procedure-local analysis described by this package.
type Options struct {
	// Conservative is retained as an explicit API marker. False does not turn
	// off conservative handling; the first implementation has no unsound mode.
	Conservative bool
}

// Source identifies the origin of a value.
type Source struct {
	Kind        SourceKind   `json:"kind"`
	Label       string       `json:"label"`
	Range       vbaast.Range `json:"range"`
	StatementID int          `json:"statementId,omitempty"`
}

// Sink identifies the sensitive operation receiving a value.
type Sink struct {
	Kind        SinkKind     `json:"kind"`
	Label       string       `json:"label"`
	Range       vbaast.Range `json:"range"`
	StatementID int          `json:"statementId,omitempty"`
	Argument    int          `json:"argument,omitempty"`
}

// PathStep records one deterministic propagation step between a source and a
// sink. The analysis never exposes parser nodes or source byte slices.
type PathStep struct {
	Kind        string       `json:"kind"`
	Label       string       `json:"label"`
	Range       vbaast.Range `json:"range"`
	StatementID int          `json:"statementId,omitempty"`
}

// Finding is one source/sink flow. It is intentionally independent of the
// analyzer and LSP diagnostic protocols.
type Finding struct {
	State   State      `json:"state"`
	Source  Source     `json:"source"`
	Sink    Sink       `json:"sink"`
	Path    []PathStep `json:"path"`
	Message string     `json:"message"`
	Reason  string     `json:"reason"`
}

// Result contains findings and entry states for reachable statement blocks.
// States is useful to protocol adapters and tests without exposing internal
// provenance implementation details.
type Result struct {
	Findings []Finding                        `json:"findings"`
	States   map[cfg.BlockID]map[string]State `json:"states,omitempty"`
}

// Analyze is an alias for AnalyzeProcedure for callers that prefer the short
// package-level verb.
func Analyze(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) Result {
	return AnalyzeProcedure(procedure, graph, options)
}

func sourceKey(source Source) string {
	return fmt.Sprintf("%s\x00%s\x00%d\x00%d\x00%d\x00%d", source.Kind, strings.ToLower(source.Label), source.Range.StartByte, source.Range.EndByte, source.Range.StartLine, source.StatementID)
}

func pathKey(path []PathStep) string {
	var b strings.Builder
	for _, step := range path {
		fmt.Fprintf(&b, "%s\x00%s\x00%d\x00%d\x00%d\x00%d\x01", step.Kind, strings.ToLower(step.Label), step.Range.StartByte, step.Range.EndByte, step.Range.StartLine, step.StatementID)
	}
	return b.String()
}
