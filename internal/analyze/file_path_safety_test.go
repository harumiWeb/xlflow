package analyze

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA245ClassifiesDestructivePathOperations(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Kill "*.*"
    Kill raw & "*"
    Kill ""
    Kill "C:\"
    Kill "C:\safe\..\outside.txt"
    Kill CurDir$ & "\\current.txt"
    Kill ThisWorkbook.Path & "\\" & raw
    Kill UnknownPath()
    RmDir "."
    Name raw As raw
    FileCopy "C:\\safe\\a.txt", "C:\\safe\\a.txt"
    fso.DeleteFile raw
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\\safe.xlsm"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA245")
	if len(got) < 5 {
		t.Fatalf("VBA245 findings = %+v, want destructive path findings", got)
	}
	seen := map[string]bool{}
	for _, finding := range got {
		if finding.FileOperation == nil {
			t.Fatalf("missing file operation context: %+v", finding)
		}
		seen[finding.FileOperation.RiskKind] = true
	}
	for _, risk := range []string{"wildcard_delete", "relative_path", "same_source_destination", "unknown_path", "empty_path", "root_path", "directory_traversal", "untrusted_filename", "current_directory_dependency"} {
		if !seen[risk] {
			t.Fatalf("risk %q missing from %+v", risk, seen)
		}
	}
	encoded, err := json.Marshal(got[0])
	if err != nil || !strings.Contains(string(encoded), `"file_operation"`) {
		t.Fatalf("file operation context JSON = %s, err=%v", encoded, err)
	}
	if len(findingsByCode(findings, "VBA224")) != 0 {
		t.Fatalf("VBA224 duplicated VBA245 file findings: %+v", findings)
	}
}

func TestVBA245AcceptsTrustedAbsoluteAndExistenceChecks(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim p As String
    p = ThisWorkbook.Path & "\\fixed.txt"
    If Dir(p) <> "" Then Debug.Print p
    If CreateObject("Scripting.FileSystemObject").FileExists(p) Then Debug.Print p
    Kill p
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA245"); len(got) != 0 {
		t.Fatalf("trusted anchored path should be clean: %+v", got)
	}
}

func TestVBA245ScopesProcedureLocalConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Private Const TargetPath As String = "C:\safe.txt"

Public Sub First()
    Kill TargetPath
End Sub

Public Sub Second()
    Const TargetPath As String = "..\unsafe.txt"
    Kill TargetPath
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA245")
	if len(got) != 1 || got[0].Procedure != "Second" || got[0].Line != 11 {
		t.Fatalf("procedure-local path constant scope = %+v, want only Second line 11", got)
	}
}

func TestVBA245FallsBackToVBA224WhenDisabled(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Kill raw
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	cfg.Analyze.DetectUnsafeFilePath = false
	cfg.Analyze.DisabledRules = []string{"VBA245"}
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(findings, "VBA245")) != 0 || len(findingsByCode(findings, "VBA224")) != 1 {
		t.Fatalf("disabled VBA245 fallback = %+v", findings)
	}

	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), cfg, []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Kill raw
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(realtime, "VBA245")) != 0 || len(findingsByCode(realtime, "VBA224")) != 1 {
		t.Fatalf("realtime disabled VBA245 fallback = %+v", realtime)
	}
}

func TestVBA245CoversBuildPathBinaryAndFSOValidation(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim fso As Object
    Dim p As String, dest As String, tmp As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = fso.BuildPath(ThisWorkbook.Path, "fixed.txt")
    dest = ThisWorkbook.Path & "\\new.txt"
    If fso.FileExists(dest) Then Exit Sub
    fso.CopyFile Source:=p, Destination:=dest, OverwriteFiles:=False
    fso.OpenTextFile p, 1
    Open "C:relative.bin" For Binary As #1
    Put #1, , "data"
    Close #1
    Open "C:\\input.bin" For Input As #1
    Put #1, , "not a tracked binary handle"
    tmp = fso.GetTempName()
    fso.CreateTextFile tmp, False
    Exit Sub
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA245")
	seen := map[string]bool{}
	for _, finding := range got {
		if finding.FileOperation != nil {
			seen[finding.FileOperation.RiskKind] = true
		}
	}
	for _, risk := range []string{"relative_path", "temporary_cleanup_missing"} {
		if !seen[risk] {
			t.Fatalf("risk %q missing from %+v", risk, seen)
		}
	}
	relativeCount := 0
	for _, finding := range got {
		if finding.FileOperation != nil && finding.FileOperation.RiskKind == "relative_path" {
			relativeCount++
		}
	}
	if relativeCount != 1 {
		t.Fatalf("closed binary handle should not be reused: got %d relative findings (%+v)", relativeCount, got)
	}
	if seen["unchecked_overwrite"] {
		t.Fatalf("explicit overwrite false / existence guard should be safe: %+v", got)
	}
}

func TestVBA245MixedNamedArgumentsAndContextSerialization(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile raw, Destination:=ThisWorkbook.Path & "\\dest.txt", OverwriteFiles:=False
    fso.CopyFile Source:=raw, ThisWorkbook.Path & "\\dest2.txt", OverwriteFiles:=False
    fso.CopyFile raw, ThisWorkbook.Path & "\\dest3.txt"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA245")
	if len(got) != 4 {
		t.Fatalf("mixed named-argument findings = %+v, want two guarded source findings and two implicit-overwrite findings", got)
	}
	falseCount, absentCount := 0, 0
	for _, finding := range got {
		if finding.FileOperation == nil {
			t.Fatalf("missing file operation context: %+v", finding)
		}
		encoded, err := json.Marshal(finding)
		if err != nil {
			t.Fatal(err)
		}
		if finding.FileOperation.Overwrite != nil && !*finding.FileOperation.Overwrite {
			falseCount++
			if !strings.Contains(string(encoded), `"overwrite":false`) {
				t.Fatalf("explicit false overwrite was not serialized: %s", encoded)
			}
		} else if finding.FileOperation.Overwrite == nil {
			absentCount++
			if strings.Contains(string(encoded), `"overwrite"`) {
				t.Fatalf("unset overwrite should be omitted: %s", encoded)
			}
		}
	}
	if falseCount != 2 || absentCount != 2 {
		t.Fatalf("overwrite context counts false=%d absent=%d, want 2/2", falseCount, absentCount)
	}
}

func TestVBA245TemporaryDetectionAvoidsNameSubstrings(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile "C:\\template.xlsx", False
    fso.CreateTextFile "C:\\Templates\\report.xlsm", False
    fso.CreateTextFile "C:\\attempts\\report.xlsm", False
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA245"); len(got) != 0 {
		t.Fatalf("permanent names must not be classified as temporary: %+v", got)
	}
}

func TestVBA245MalformedSaveAsNamedTextDoesNotPanic(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    ThisWorkbook.SaveAs Replace("C:\\safe.xlsx", "name:", "")
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	if _, err := (Analyzer{RootDir: dir, Config: cfg}).Run(); err != nil {
		t.Fatalf("malformed SaveAs named text should not abort analysis: %v", err)
	}
}

func TestVBA245IgnoresUserDefinedSameNameAndNonDestructiveChecks(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim obj As Object
    obj.Kill "relative.txt"
    obj.SaveAs "relative.xlsx"
    Debug.Print Dir("*.txt")
End Sub
Public Sub Kill(ByVal path As String)
    Debug.Print path
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA245"); len(got) != 0 {
		t.Fatalf("user-defined/member operations and existence checks should be ignored: %+v", got)
	}
}
