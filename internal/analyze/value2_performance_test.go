package analyze

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func value2PerformanceTestConfig() config.Config {
	cfg := config.Default()
	cfg.Analyze.DetectValue2PerformanceOpportunities = true
	return cfg
}

func TestVBA243DetectsBulkDynamicLoopAndVariantSignals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet, lastRow As Long)
	Dim values As Variant
	Dim smallValues As String
	Dim smallVariant As Variant
	Dim implicitVariant
	Dim i As Long

	values = ws.Range("A1:J10").Value
	values = ws.Range("A1:J10").Value
	smallValues = ws.Range("A1:J9").Value
	smallVariant = ws.Range("A1:A99").Value
	implicitVariant = ws.Range("B1:B99").Value
    ws.Range("K1:T10").Value = values
    values = ws.Range("A1:A" & lastRow).Value
    For i = 1 To lastRow
        ws.Cells(i, 1).Value = i
    Next i
    Debug.Print ws.Range("A1:J10").Value2
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA243")
	if len(got) != 6 {
		t.Fatalf("VBA243 findings = %+v, want bulk read/write, Variant transfer, dynamic read, and loop access", got)
	}
	if got[0].Severity != "information" || got[1].Severity != "information" || got[2].Severity != "information" || got[3].Severity != "information" || got[4].Severity != "information" {
		t.Fatalf("bulk/dynamic severities = %+v, want information", got)
	}
	if got[5].Severity != "warning" {
		t.Fatalf("loop severity = %+v, want warning", got[5])
	}
	for _, finding := range got {
		if !strings.Contains(finding.Reason, "Date") || !strings.Contains(finding.Reason, "Currency") || !strings.Contains(finding.Suggestion, "Value2") {
			t.Fatalf("VBA243 finding lacks semantic guidance: %+v", finding)
		}
	}
}

func TestVBA243ExcludesDateCurrencyAndWeakSignals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Dim asDate As Date
    Dim asCurrency As Currency
    Dim values As Variant
    Dim weak As String
    Dim uncertain As Variant
    Dim formatted As Variant
    Dim Range As Object

    asDate = ws.Range("A1:A100").Value
    asCurrency = ws.Range("B1:B100").Value
    values = ws.Range("C1:C100").Value
    asDate = CDate(values)
    ws.Range("D1:D100").Value = asCurrency
    ws.Range("I1:I100").Value = Date
    If TypeName(values) = "Date" Then Debug.Print values
    uncertain = ws.Range("J1:J100").Value
    Consume uncertain
    formatted = ws.Range("K1:K100").Value
    Debug.Print Format(formatted, "Currency")
    AcceptDate ws.Range("L1:L100").Value
    weak = ws.Range("E1:E99").Value
    ' ws.Range("F1:F100").Value
    Debug.Print "ws.Range(\"G1:G100\").Value"
    values = Range("H1:H100").Value
End Sub

Private Sub Consume(ByVal value As Variant)
End Sub

Private Sub AcceptDate(ByVal value As Date)
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA243"); len(got) != 0 {
		t.Fatalf("intentional Date/Currency or weak accesses produced VBA243: %+v", got)
	}
}

func TestVBA243OptInSuppressionAndRealtimeParity(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ws As Worksheet)
    Dim values As Variant
    ' xlflow:disable-next-line VBA243
    values = ws.Range("A1:J10").Value
End Sub
`)
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", string(source))
	if findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run(); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(findings, "VBA243"); len(got) != 0 {
		t.Fatalf("disabled-by-default VBA243 = %+v", got)
	}

	cfg := value2PerformanceTestConfig()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(batch, "VBA243"); len(got) != 0 {
		t.Fatalf("inline-suppressed batch VBA243 = %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA243"); len(got) != 0 {
		t.Fatalf("inline-suppressed realtime VBA243 = %+v", got)
	}
}

func TestVBA243RealtimeMatchesBatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ws As Worksheet)
    Dim values As Variant
    values = ws.Range("A1:J10").Value
End Sub
`)
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", string(source))
	cfg := value2PerformanceTestConfig()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	batchFindings := findingsByCode(batch, "VBA243")
	realtimeFindings := findingsByCode(realtime, "VBA243")
	if len(batchFindings) != 1 || len(realtimeFindings) != 1 {
		t.Fatalf("batch/realtime VBA243 = %+v / %+v, want one each", batchFindings, realtimeFindings)
	}
	if batchFindings[0].Severity != realtimeFindings[0].Severity || batchFindings[0].Severity != "information" {
		t.Fatalf("batch/realtime VBA243 severity = %q / %q, want information", batchFindings[0].Severity, realtimeFindings[0].Severity)
	}
}

func TestVBA243RequiresProvenDynamicRangeConstruction(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet, lastRow As Long, address As String)
	Dim values As Variant
	Dim dynamicRange As Range
	Dim dynamicRangeAlias As Range

	values = ws.Range(address).Value
	values = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 10)).Value
	Set dynamicRange = ws.Range("A1:A" & lastRow)
	Set dynamicRangeAlias = dynamicRange
	values = dynamicRangeAlias.Value
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA243")
	if len(got) != 2 {
		t.Fatalf("VBA243 findings = %+v, want Cells-pair and traceable alias only", got)
	}
}

func TestVBA243RecognizesDynamicRangeProperties(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Dim values As Variant
    values = ws.CurrentRegion.Value
    values = ws.UsedRange.Value
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA243"); len(got) != 2 {
		t.Fatalf("VBA243 findings = %+v, want CurrentRegion and UsedRange", got)
	}
}

func TestVBA243DoesNotTreatDateWordsInLabelsOrCommentsAsIntentional(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Debug.Print "Date", ws.Range("J1:J100").Value
    Debug.Print ws.Range("K1:K100").Value ' currency report label
    Debug.Print "VarType(values)=Date", ws.Range("L1:L100").Value
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA243"); len(got) != 3 {
		t.Fatalf("VBA243 findings = %+v, want three non-intentional accesses", got)
	}
}

func TestVBA243HonorsProjectDisabledRules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ws As Worksheet)
    Dim values As Variant
    values = ws.Range("A1:J10").Value
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	cfg := value2PerformanceTestConfig()
	cfg.Analyze.DisabledRules = []string{"VBA243"}
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	loaded, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if loaded.Analyze.DetectValue2PerformanceOpportunities {
		t.Fatal("VBA243 should be disabled through [analyze].disabled_rules")
	}
	findings, err := (Analyzer{RootDir: dir, Config: loaded}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA243"); len(got) != 0 {
		t.Fatalf("disabled VBA243 findings = %+v", got)
	}
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	realtime, err := SourceRealtimeFindings(dir, path, loaded, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA243"); len(got) != 0 {
		t.Fatalf("disabled realtime VBA243 findings = %+v", got)
	}
}

func TestVBA243ClassifiesValueComparisonAsRead(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Debug.Print ws.Range("A1:J100").Value = 1
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: value2PerformanceTestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA243")
	if len(got) != 1 || !strings.Contains(got[0].Message, "read") {
		t.Fatalf("VBA243 comparison findings = %+v, want one read", got)
	}
}
