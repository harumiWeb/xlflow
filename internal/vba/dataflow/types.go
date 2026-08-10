// Package dataflow provides a conservative, procedure-local source-to-sink
// analysis over normalized VBA IR and its control-flow graph.
package dataflow

import (
	"fmt"
	"strconv"
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
	// SinkShellExecute covers both Shell.Application.ShellExecute and the
	// Win32 ShellExecute/A/W entry points.  The launcher is retained in the
	// command finding metadata so adapters can distinguish the two shapes.
	SinkShellExecute    SinkKind = "shell_execute"
	SinkSQLExecution    SinkKind = "sql_execution"
	SinkDestructiveFile SinkKind = "destructive_file_operation"
	SinkWorkbooksOpen   SinkKind = "workbooks_open"
	SinkSaveAs          SinkKind = "save_as"
	SinkHTTPURL         SinkKind = "http_request_url"
	SinkHTTPHeader      SinkKind = "http_request_header"
)

// SourceSpec and SinkSpec are the public, protocol-neutral catalog entries.
type SourceSpec struct {
	Kind        SourceKind `json:"kind"`
	Label       string     `json:"label"`
	Description string     `json:"description"`
}

// CommandRiskClass is the top-level distinction made by the command
// construction analysis.  Injection is only used when a known tainted source
// reaches a command-bearing role; unknown input remains a general launch
// warning and is never presented as a confirmed vulnerability.
type CommandRiskClass string

const (
	CommandRiskInjection     CommandRiskClass = "injection"
	CommandRiskProcessLaunch CommandRiskClass = "process_launch"
)

// CommandRiskKind identifies the more specific observation behind a command
// finding.  Values are stable wire strings for analyzer/LSP projections.
type CommandRiskKind string

const (
	CommandRiskTaintedCommandText     CommandRiskKind = "tainted_command_text"
	CommandRiskUnknownOrigin          CommandRiskKind = "unknown_origin"
	CommandRiskUnquotedExecutablePath CommandRiskKind = "unquoted_executable_path"
	CommandRiskCredentialExposure     CommandRiskKind = "credential_exposure"
	CommandRiskObservability          CommandRiskKind = "observability"
)

// CommandRole identifies the argument role at a process-launch sink.
type CommandRole string

const (
	CommandRoleExecutable   CommandRole = "executable"
	CommandRoleArguments    CommandRole = "arguments"
	CommandRoleShellCommand CommandRole = "shell_command"
	CommandRoleURL          CommandRole = "url"
	CommandRoleDocument     CommandRole = "document"
	CommandRoleWindowStyle  CommandRole = "window_style"
	CommandRoleWait         CommandRole = "wait"
	CommandRoleUnknown      CommandRole = "unknown"
)

// CommandExecution describes the classified process-launch call.  It is
// deliberately independent from the analyzer protocol and can be projected
// into JSON, LSP, or other clients without exposing parser internals.
type CommandExecution struct {
	Launcher    string       `json:"launcher"`
	Interpreter string       `json:"interpreter,omitempty"`
	Role        CommandRole  `json:"command_role"`
	Argument    int          `json:"argument"`
	Range       vbaast.Range `json:"range"`
}

// CommandFinding is a command-specific source-to-sink observation.  Source
// and Path are populated for known origins; static observations such as an
// unquoted constant executable path have an empty Source and StateClean.
type CommandFinding struct {
	State       State            `json:"state"`
	Source      Source           `json:"source,omitempty"`
	Execution   CommandExecution `json:"command_execution"`
	RiskClass   CommandRiskClass `json:"risk_class"`
	RiskKind    CommandRiskKind  `json:"risk_kind"`
	OriginState State            `json:"origin_state"`
	Path        []PathStep       `json:"path,omitempty"`
	Message     string           `json:"message"`
	Reason      string           `json:"reason"`
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
	{Kind: SinkShellExecute, Label: "ShellExecute", Description: "Launches an executable, document, or URL through ShellExecute."},
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
	// IsKnownConstant reports whether an identifier not bound in the current
	// procedure resolves to a module, project, or host-provided constant.
	// Local declarations still take precedence over this callback.
	IsKnownConstant func(name string) bool
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
	Findings        []Finding                        `json:"findings"`
	CommandFindings []CommandFinding                 `json:"command_findings,omitempty"`
	States          map[cfg.BlockID]map[string]State `json:"states,omitempty"`
}

// Analyze is an alias for AnalyzeProcedure for callers that prefer the short
// package-level verb.
func Analyze(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) Result {
	return AnalyzeProcedure(procedure, graph, options)
}

func sourceKey(source Source) string {
	return fmt.Sprintf("%s\x00%s\x00%d\x00%d\x00%d\x00%d", source.Kind, strings.ToLower(source.Label), source.Range.StartByte, source.Range.EndByte, source.Range.StartLine, source.StatementID)
}

// comparePathKeys compares the byte sequences formerly produced by pathKey
// without materializing those sequences. Labels use the allocation-free ASCII
// case fold on the common path and fall back to strings.ToLower for Unicode.
func comparePathKeys(left, right []PathStep) int {
	leftKey := pathKeyIterator{path: left}
	rightKey := pathKeyIterator{path: right}
	for {
		leftByte, leftOK := leftKey.next()
		rightByte, rightOK := rightKey.next()
		switch {
		case !leftOK && !rightOK:
			return 0
		case !leftOK:
			return -1
		case !rightOK:
			return 1
		case leftByte < rightByte:
			return -1
		case leftByte > rightByte:
			return 1
		}
	}
}

type pathKeyIterator struct {
	path []PathStep

	step  int
	field int
	index int

	labelReady  bool
	labelASCII  bool
	foldedLabel string

	decimalReady  bool
	decimalLength int
	decimal       [32]byte
}

func (it *pathKeyIterator) next() (byte, bool) {
	for it.step < len(it.path) {
		step := &it.path[it.step]
		switch it.field {
		case 0:
			if it.index < len(step.Kind) {
				value := step.Kind[it.index]
				it.index++
				return value, true
			}
			it.advanceField()
		case 1, 3, 5, 7, 9:
			it.advanceField()
			return 0, true
		case 2:
			if !it.labelReady {
				it.prepareLabel(step.Label)
			}
			label := step.Label
			if !it.labelASCII {
				label = it.foldedLabel
			}
			if it.index < len(label) {
				value := label[it.index]
				if it.labelASCII && value >= 'A' && value <= 'Z' {
					value += 'a' - 'A'
				}
				it.index++
				return value, true
			}
			it.advanceField()
		case 4, 6, 8, 10:
			if !it.decimalReady {
				it.prepareDecimal(step)
			}
			if it.index < it.decimalLength {
				value := it.decimal[it.index]
				it.index++
				return value, true
			}
			it.advanceField()
		case 11:
			it.step++
			it.field = 0
			it.index = 0
			return 1, true
		}
	}
	return 0, false
}

func (it *pathKeyIterator) advanceField() {
	it.field++
	it.index = 0
	it.labelReady = false
	it.foldedLabel = ""
	it.decimalReady = false
	it.decimalLength = 0
}

func (it *pathKeyIterator) prepareLabel(label string) {
	it.labelReady = true
	it.labelASCII = true
	for index := 0; index < len(label); index++ {
		if label[index] >= 0x80 {
			it.labelASCII = false
			it.foldedLabel = strings.ToLower(label)
			return
		}
	}
}

func (it *pathKeyIterator) prepareDecimal(step *PathStep) {
	value := 0
	switch it.field {
	case 4:
		value = step.Range.StartByte
	case 6:
		value = step.Range.EndByte
	case 8:
		value = step.Range.StartLine
	case 10:
		value = step.StatementID
	}
	formatted := strconv.AppendInt(it.decimal[:0], int64(value), 10)
	it.decimalLength = len(formatted)
	it.decimalReady = true
}
