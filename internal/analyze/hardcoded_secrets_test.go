package analyze

import (
	"encoding/json"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA223DetectsStructuredCredentialsAndRedactsOutput(t *testing.T) {
	dir := t.TempDir()
	apiKey := "g" + "h" + "p_" + strings.Repeat("a", 24)
	awsAccessKey := "A" + "KIA" + strings.Repeat("A", 16)
	source := `Option Explicit
Private Const ApiKey As String = "__API_KEY__"
Public Sub Run()
  Dim password As String
  password = "hardcoded-password-value"
  Dim connectionString As String
  connectionString = "Provider=SQLOLEDB;User ID=appuser;Password=connection-password;"
  Dim authorization As String
  authorization = "Bearer eyJhbGciOiJIUzI1NiJ9.secret-token-value"
  Dim basicAuth As String
  basicAuth = "Basic QWxhZGRpbjpvcGVuIHNlc2FtZQ=="
  Dim endpoint As String
  endpoint = "https://service:url-password@internal.invalid/api"
  Dim privateKey As String
  privateKey = "-----BEGIN RSA PRIVATE KEY-----"
  Dim webhookURL As String
  webhookURL = "https://hooks.slack.com/services/T123/B123/real-webhook-token"
  Dim awsAccessKey As String
  awsAccessKey = "__AWS_ACCESS_KEY__"
  password = Environ$("PASSWORD")
  Debug.Print "Password=your-password"
  Debug.Print "ordinary text"
End Sub
`
	source = strings.ReplaceAll(source, "__API_KEY__", apiKey)
	source = strings.ReplaceAll(source, "__AWS_ACCESS_KEY__", awsAccessKey)
	writeModule(t, dir, "Main.bas", source)

	cfg := config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	batchSecrets := findingsByCode(batch, "VBA223")
	realtimeSecrets := findingsByCode(realtime, "VBA223")
	if len(batchSecrets) != 9 {
		t.Fatalf("VBA223 findings = %+v, want nine structured credentials", batchSecrets)
	}
	if !reflect.DeepEqual(batchSecrets, realtimeSecrets) {
		t.Fatalf("batch/realtime VBA223 findings differ:\nbatch=%+v\nrealtime=%+v", batchSecrets, realtimeSecrets)
	}

	encoded, err := json.Marshal(batchSecrets)
	if err != nil {
		t.Fatal(err)
	}
	output := string(encoded)
	for _, secret := range []string{
		apiKey,
		"hardcoded-password-value",
		"connection-password",
		"eyJhbGciOiJIUzI1NiJ9.secret-token-value",
		"QWxhZGRpbjpvcGVuIHNlc2FtZQ==",
		"url-password",
		"-----BEGIN RSA PRIVATE KEY-----",
		"real-webhook-token",
		awsAccessKey,
	} {
		if strings.Contains(output, secret) {
			t.Fatalf("VBA223 JSON contains secret %q: %s", secret, output)
		}
	}
	if !strings.Contains(output, "[REDACTED]") {
		t.Fatalf("VBA223 JSON does not contain fixed redaction: %s", output)
	}
	for _, finding := range batchSecrets {
		for _, line := range finding.NearbyCode {
			if strings.Contains(line, "hardcoded-password-value") || strings.Contains(line, "connection-password") {
				t.Fatalf("VBA223 nearby code contains a secret: %q", line)
			}
		}
	}
}

func TestVBA223IgnoresPlaceholdersCommentsAndEnvironmentReferences(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
' password = "comment-secret"
Public Sub Run()
  Dim password As String
  password = "your-password"
  password = Environ$("PASSWORD")
  Debug.Print "Password=example"
  Debug.Print "ordinary text"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA223"); len(got) != 0 {
		t.Fatalf("placeholder/comment/environment findings = %+v, want none", got)
	}
}

func TestVBA223HonorsDisabledRulesAndInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim password As String
  ' xlflow:disable-next-line VBA223
  password = "intentional-fixture-secret"
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	path := filepath.Join(dir, "src", "modules", "Main.bas")

	cfg := config.Default()
	cfg.Analyze.DisabledRules = []string{"VBA223"}
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	loaded, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if loaded.Analyze.DetectHardcodedSecrets {
		t.Fatal("VBA223 should be disabled through [analyze].disabled_rules")
	}
	findings, err := (Analyzer{RootDir: dir, Config: loaded}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA223"); len(got) != 0 {
		t.Fatalf("disabled VBA223 findings = %+v", got)
	}

	cfg = config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(batch, "VBA223"); len(got) != 0 {
		t.Fatalf("inline-suppressed batch findings = %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA223"); len(got) != 0 {
		t.Fatalf("inline-suppressed realtime findings = %+v", got)
	}
}
