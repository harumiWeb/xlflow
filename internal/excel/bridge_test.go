package excel

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"os"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/output"
)

type fakeBridgeProvider struct {
	name     string
	response excelbridge.Response
	err      error
}

func (f fakeBridgeProvider) Name() string {
	return f.name
}

func (f fakeBridgeProvider) Supports(string) bool {
	return true
}

func (f fakeBridgeProvider) Info(context.Context) (excelbridge.Info, error) {
	return excelbridge.Info{Name: f.name, Version: "test"}, nil
}

func (f fakeBridgeProvider) Execute(_ context.Context, req excelbridge.Request) (excelbridge.Response, error) {
	_ = req
	return f.response, f.err
}

type trackingBridgeProvider struct {
	name         string
	supports     bool
	supportsFunc func(string) bool
	response     excelbridge.Response
	err          error
	callCount    *int
	requests     *[]excelbridge.Request
}

func (p trackingBridgeProvider) Name() string {
	return p.name
}

func (p trackingBridgeProvider) Supports(command string) bool {
	if p.supportsFunc != nil {
		return p.supportsFunc(command)
	}
	return p.supports
}

func (p trackingBridgeProvider) Info(context.Context) (excelbridge.Info, error) {
	return excelbridge.Info{Name: p.name, Version: "test"}, nil
}

func (p trackingBridgeProvider) Execute(_ context.Context, req excelbridge.Request) (excelbridge.Response, error) {
	if p.callCount != nil {
		*p.callCount = *p.callCount + 1
	}
	if p.requests != nil {
		*p.requests = append(*p.requests, req)
	}
	return p.response, p.err
}

func TestScriptResultAcceptsScalarLogString(t *testing.T) {
	var result ScriptResult
	body := []byte(`{"status":"ok","command":"session","logs":"stopped xlflow Excel session"}`)
	if err := json.Unmarshal(body, &result); err != nil {
		t.Fatal(err)
	}
	if len(result.Logs) != 1 || result.Logs[0] != "stopped xlflow Excel session" {
		t.Fatalf("unexpected logs: %+v", result.Logs)
	}
}

func TestResolveBridgeModePrecedence(t *testing.T) {
	mode, source, err := excelbridge.ResolveMode("", "auto", "dotnet")
	if err != nil {
		t.Fatal(err)
	}
	if mode != excelbridge.ModeAuto || source != "env" {
		t.Fatalf("ResolveMode env precedence = (%q, %q), want (auto, env)", mode, source)
	}

	mode, source, err = excelbridge.ResolveMode("dotnet", "auto", "auto")
	if err != nil {
		t.Fatal(err)
	}
	if mode != excelbridge.ModeDotNet || source != "cli" {
		t.Fatalf("ResolveMode cli precedence = (%q, %q), want (dotnet, cli)", mode, source)
	}
}

func TestResolveBridgeModeRejectsPowerShell(t *testing.T) {
	if _, _, err := excelbridge.ResolveMode("", "powershell", "dotnet"); err == nil {
		t.Fatal("expected powershell bridge mode to be rejected")
	}
}

func TestRunnerRejectsDotNetBridgeWithoutFallback(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"run","logs":[],"error":{"code":"BRIDGE_COMMAND_UNSUPPORTED","message":"Command 'run' is not supported by the .NET bridge.","source":"xlflow-excel-bridge","phase":"bridge.capability"}}`)},
		}
	}

	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	marker := filepath.Join(root, "powershell-fallback-marker.txt")
	script := "$null = New-Item -ItemType File -Path \"" + marker + "\" -Force\n@{ status = \"ok\"; command = \"run\"; logs = @(\"unexpected powershell execution\") } | ConvertTo-Json -Compress"
	if err := os.WriteFile(filepath.Join(scriptsDir, "run.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}

	env, code, err := Runner{RootDir: root, BridgeMode: "dotnet"}.Run(config.Default(), RunOptions{Macro: "Main.Run"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Error == nil || env.Error.Code != "BRIDGE_COMMAND_UNSUPPORTED" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
	if _, err := os.Stat(marker); !os.IsNotExist(err) {
		t.Fatalf("expected no powershell fallback marker, got %v", err)
	}
}

func TestRunnerAutoBridgeUsesDotNetWhenSupported(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls int
	var powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"process","logs":[],"process":[{"pid":1234,"has_workbook":true}]}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[{"pid":9999,"has_workbook":false}]}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	first, ok := processes[0].(map[string]interface{})
	if !ok || first["pid"] != float64(1234) {
		t.Fatalf("unexpected process payload: %+v", processes[0])
	}
}

func TestRunnerAutoBridgeDoesNotFallBackToPowerShellForUnsupportedDotNetInspectTarget(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls int
	var powerShellCalls int
	var dotNetRequests []excelbridge.Request
	var powerShellRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"inspect","logs":[],"error":{"code":"BRIDGE_COMMAND_UNSUPPORTED","message":"Inspect target 'used-range' is not supported by the .NET bridge.","source":"xlflow-excel-bridge","phase":"bridge.capability"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				requests:  &powerShellRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"inspect","logs":["inspected live used range for Sheet1"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","sheet":"Sheet1","description":"Workbook currently open through xlflow session"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"dirty":false,"needs_save":false},"inspect":{"target":"used-range","source":"excel_com","target_info":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"range":{"sheet":"Sheet1","used_range":"A1","row_count":1,"column_count":1,"values":[["value"]],"truncated":false,"max_rows":100,"max_cols":30,"returned_range":"A1"}}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.Inspect(config.Default(), InspectOptions{
		Target:    "used-range",
		Sheet:     "Sheet1",
		Limits:    map[string]int{"max_rows": 100, "max_cols": 30},
		Session:   true,
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 || len(powerShellRequests) != 0 {
		t.Fatalf("unexpected request counts: dotnet=%d powershell=%d", len(dotNetRequests), len(powerShellRequests))
	}
	if got := dotNetRequests[0].Args["Target"]; got != "used-range" {
		t.Fatalf("dotnet target = %q, want used-range", got)
	}
	if env.Command != "inspect" {
		t.Fatalf("Command = %q, want inspect", env.Command)
	}
	if env.Error == nil || env.Error.Code != "BRIDGE_COMMAND_UNSUPPORTED" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDotNetInspectResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"inspect","logs":["inspected live workbook C:\\temp\\Book.xlsm"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","description":"Workbook currently open through xlflow session"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"dirty":false,"needs_save":false},"inspect":{"target":"workbook","source":"excel_com","target_info":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","note":"This command inspected the live workbook currently open in Excel through xlflow session."},"workbook":{"path":"C:\\temp\\Book.xlsm","name":"Book.xlsm","sheets":[{"name":"Sheet1","index":1,"visible":true,"used_range":"$A$1","row_count":1,"column_count":1}]}}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Inspect(config.Default(), InspectOptions{
		Target:    "workbook",
		Session:   true,
		Limits:    map[string]int{"max_rows": 100, "max_cols": 30},
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Inspect == nil {
		t.Fatalf("expected inspect envelope fields to be populated: %+v", env)
	}
	inspectPayload, ok := env.Inspect.(map[string]interface{})
	if !ok || inspectPayload["target"] != "workbook" {
		t.Fatalf("unexpected inspect payload: %+v", env.Inspect)
	}
}

func TestRunnerDotNetInspectFormResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"inspect-form","logs":["inspected both UserForm UserForm1"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","note":"Runtime inspection used a temporary workbook copy."},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":false,"needs_save":false},"forms":{"runtime":{"name":"UserForm1","basis":"runtime","controls":[]},"designer":{"name":"UserForm1","basis":"designer","controls":[]}},"warnings":[{"code":"runtime_form_temp_copy","message":"Runtime inspection executed against a temporary workbook copy so the source workbook and live session are not mutated."}]}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.InspectForm(config.Default(), InspectFormOptions{
		Name:           "UserForm1",
		Basis:          "both",
		StrictDesigner: true,
		Session:        true,
		Keepalive:      CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Forms == nil || env.Warnings == nil {
		t.Fatalf("expected inspect-form envelope fields to be populated: %+v", env)
	}
	formsPayload, ok := env.Forms.(map[string]interface{})
	if !ok {
		t.Fatalf("expected forms payload map, got %#v", env.Forms)
	}
	runtimeForm, ok := formsPayload["runtime"].(map[string]interface{})
	if !ok || runtimeForm["basis"] != "runtime" {
		t.Fatalf("unexpected runtime form payload: %+v", formsPayload["runtime"])
	}
	designerForm, ok := formsPayload["designer"].(map[string]interface{})
	if !ok || designerForm["basis"] != "designer" {
		t.Fatalf("unexpected designer form payload: %+v", formsPayload["designer"])
	}
}

func TestRunnerDotNetFormWriteResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"form-write","logs":["build form UserForm1 from src/forms/specs/UserForm1.yaml"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":true,"dirty":false,"needs_save":false},"forms":{"name":"UserForm1","basis":"designer","action":"build","coordinate_system":"parent-relative","control_count":2,"spec_path":"src/forms/specs/UserForm1.yaml","overwrite":true,"source_synced":true,"source_artifacts":{"form_path":"C:\\temp\\src\\forms\\UserForm1.frm","frx_path":"C:\\temp\\src\\forms\\UserForm1.frx","code_path":"C:\\temp\\src\\forms\\code\\UserForm1.bas"}},"warnings":[],"hints":[{"code":"userform_review_commands","message":"review"}]}`)},
		}
	}

	spec := forms.FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          forms.FormSpecForm{Name: "UserForm1"},
		Controls:      []forms.FormSpecControl{{ID: "label1", Name: "Label1", Type: "label"}},
	}
	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.FormWrite(config.Default(), FormWriteOptions{
		Action:    "build",
		SpecPath:  "src/forms/specs/UserForm1.yaml",
		Spec:      spec,
		Overwrite: true,
		Session:   true,
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Command != "form build" {
		t.Fatalf("Command = %q, want form build", env.Command)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Forms == nil {
		t.Fatalf("expected form-write envelope fields to be populated: %+v", env)
	}
	formsPayload, ok := env.Forms.(map[string]interface{})
	if !ok {
		t.Fatalf("expected forms payload map, got %#v", env.Forms)
	}
	if formsPayload["action"] != "build" || formsPayload["name"] != "UserForm1" {
		t.Fatalf("unexpected forms payload: %+v", formsPayload)
	}
}

func TestRunnerDotNetFormExportImageResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"form-export-image","logs":["exported runtime UserForm UserForm1 to C:\\temp\\UserForm1.png"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","form":"UserForm1","capture_state":"temporary_copy","note":"Runtime export used a temporary workbook copy."},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":false,"needs_save":false},"forms":{"name":"UserForm1","basis":"runtime"},"output":{"path":"C:\\temp\\UserForm1.png","format":"png","width_px":320,"height_px":240},"warnings":[{"code":"runtime_form_temp_copy","message":"Runtime export executed against a temporary workbook copy so the source workbook and live session are not mutated."}]}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.FormExportImage(config.Default(), FormExportImageOptions{
		Name:      "UserForm1",
		OutPath:   `C:\temp\UserForm1.png`,
		Session:   true,
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Forms == nil || env.Output == nil {
		t.Fatalf("expected form-export-image envelope fields to be populated: %+v", env)
	}
	formsPayload, ok := env.Forms.(map[string]interface{})
	if !ok || formsPayload["basis"] != "runtime" {
		t.Fatalf("unexpected forms payload: %+v", env.Forms)
	}
}

func TestRunnerDotNetExportImageResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"export-image","logs":["exported Sheet1!A1:C10 to C:\\temp\\sheet.png"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm","sheet":"Sheet1","range":"A1:C10"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"dirty":false,"needs_save":false},"output":{"path":"C:\\temp\\sheet.png","format":"png","default":true,"width_px":320,"height_px":240},"warnings":[{"code":"clipboard_retry_succeeded","message":"Clipboard image export succeeded after 2 attempts."}]}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.ExportImage(config.Default(), ExportImageOptions{
		Sheet:     "Sheet1",
		Range:     "A1:C10",
		OutPath:   `C:\temp\sheet.png`,
		Session:   true,
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Output == nil {
		t.Fatalf("expected export-image envelope fields to be populated: %+v", env)
	}
	outputPayload, ok := env.Output.(map[string]interface{})
	if !ok || outputPayload["format"] != "png" {
		t.Fatalf("unexpected output payload: %+v", env.Output)
	}
}

func TestRunnerDotNetProcessCleanupResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"process","logs":["terminated 1 Excel process(es)"],"process":{"action":"cleanup","mode":"pid","total":1,"results":[{"pid":1234,"terminated":true,"method":"graceful"}]}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.ProcessCleanup(ProcessCleanupOptions{
		Action: "cleanup",
		PID:    1234,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	processPayload, ok := env.Process.(map[string]interface{})
	if !ok {
		t.Fatalf("expected process cleanup payload map, got %#v", env.Process)
	}
	if processPayload["action"] != "cleanup" || processPayload["mode"] != "pid" {
		t.Fatalf("unexpected process cleanup payload: %+v", processPayload)
	}
}

func TestRunnerDotNetSessionUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"session","logs":["started xlflow Excel session"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":false,"auto_session":false,"saved":true,"dirty":false,"needs_save":false}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Session(config.Default(), "start")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["Action"]; got != "start" {
		t.Fatalf("Action = %q, want start", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil {
		t.Fatalf("expected session envelope fields to be populated: %+v", env)
	}
}

func TestRunnerDotNetAttachUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"attach","logs":["attached active Excel workbook to xlflow session"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":false,"auto_session":false,"saved":false,"dirty":true,"needs_save":true}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Attach(config.Default(), true)
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["Active"]; got != "true" {
		t.Fatalf("Active = %q, want true", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil {
		t.Fatalf("expected attach envelope fields to be populated: %+v", env)
	}
}

func TestRunnerDotNetSessionAttachActiveUsesSessionCommand(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:      string(excelbridge.ModeDotNet),
			supports:  true,
			callCount: &dotNetCalls,
			requests:  &dotNetRequests,
			response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"session","logs":["attached xlflow session to already-open workbook"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"external","owner":"external","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"external","session_requested":true,"auto_session":false,"saved":false,"dirty":false,"needs_save":false}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.SessionAttachActive(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	args := dotNetRequests[0].Args
	if args["Action"] != "attach" || args["Active"] != "true" {
		t.Fatalf("session attach args = %#v", args)
	}
	if args["MetadataPath"] == "" {
		t.Fatalf("MetadataPath must be passed for session attach: %#v", args)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil {
		t.Fatalf("expected session attach envelope fields to be populated: %+v", env)
	}
}

func TestRunnerDotNetListFormsUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"list","logs":["listed 1 form(s)"],"target":{"kind":"saved_workbook","path":"C:\\temp\\Book.xlsm"},"session":{"active":false,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"none","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":false,"session_mode":"none","session_requested":false,"auto_session":false,"saved":false,"dirty":false,"needs_save":false},"forms":{"items":[{"name":"UserForm1","form_path":"C:\\temp\\src\\forms\\UserForm1.frm"}]}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.ListForms(config.Default(), SessionCommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["Action"]; got != "forms" {
		t.Fatalf("Action = %q, want forms", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Forms == nil {
		t.Fatalf("expected list envelope fields to be populated: %+v", env)
	}
}

func TestRunnerDotNetUIButtonAddUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"ui","logs":["added workbook button xlflow.button.run-main"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":false,"needs_save":false},"ui":{"button":{"id":"run-main","name":"xlflow.button.run-main","sheet":"Menu","text":"Run Main","macro":"Module1.RunMain","cell":"B2","left":10,"top":20,"width":160,"height":40,"updated":false}}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.UIButtonAdd(config.Default(), UIButtonAddOptions{
		Sheet:       "Menu",
		Cell:        "B2",
		Text:        "Run Main",
		Macro:       "Module1.RunMain",
		ID:          "run-main",
		Width:       160,
		Height:      40,
		VerifyMacro: true,
		Session:     true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["Action"]; got != "add" {
		t.Fatalf("Action = %q, want add", got)
	}
	if got := dotNetRequests[0].Args["VerifyMacro"]; got != "true" {
		t.Fatalf("VerifyMacro = %q, want true", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.UI == nil {
		t.Fatalf("expected ui envelope fields to be populated: %+v", env)
	}
}

func TestRunnerDotNetEditCellUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"edit","logs":["edited Sheet1!C3 value in the live Excel session","Run ` + "`xlflow save --session`" + ` to persist changes to disk."],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true},"edit":{"kind":"cell","sheet":"Sheet1","cell":"C3","mutation":{"value":{"before":"","after":"42"},"events":{"mode":"off","enable_events_before":true,"enable_events_after":false,"restored":true}},"events":{"mode":"off","enable_events_before":true,"enable_events_after":false,"restored":true}}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	value := "42"
	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.EditCell(config.Default(), EditCellOptions{
		Sheet:     "Sheet1",
		Cell:      "C3",
		Value:     &value,
		Events:    EditEventOff,
		Session:   true,
		Keepalive: CommandOptions{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["Action"]; got != "cell" {
		t.Fatalf("Action = %q, want cell", got)
	}
	if got := dotNetRequests[0].Args["Events"]; got != "off" {
		t.Fatalf("Events = %q, want off", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Edit == nil {
		t.Fatalf("expected edit envelope fields to be populated: %+v", env)
	}
}

func TestRunnerRejectsDotNetProtocolMismatch(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":99,"status":"ok","command":"doctor","logs":[],"error":null,"diagnostics":{"bridge":{"name":"xlflow-excel-bridge"}}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Doctor(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Error == nil || env.Error.Code != "bridge_protocol_mismatch" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDoctorChecksWorkbookOnlyWhenRequested(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var requests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			supports: true,
			requests: &requests,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"doctor","logs":[],"error":null,"diagnostics":{"excel":{"com_activation":true,"vbide_access":true}}}`)},
		}
	}

	root := t.TempDir()
	cfg := config.Default()
	if _, _, err := (Runner{RootDir: root, BridgeMode: "dotnet"}).Doctor(cfg); err != nil {
		t.Fatal(err)
	}
	if _, _, err := (Runner{RootDir: root, BridgeMode: "dotnet"}).DoctorWithOptions(cfg, DoctorOptions{CheckWorkbook: true}); err != nil {
		t.Fatal(err)
	}

	if len(requests) != 2 {
		t.Fatalf("request count = %d, want 2", len(requests))
	}
	if got := requests[0].Args["CheckWorkbook"]; got != "false" {
		t.Fatalf("default CheckWorkbook = %q, want false", got)
	}
	if _, ok := requests[0].Args["WorkbookPath"]; ok {
		t.Fatalf("default doctor unexpectedly sent WorkbookPath: %#v", requests[0].Args)
	}
	if got := requests[1].Args["CheckWorkbook"]; got != "true" {
		t.Fatalf("opt-in CheckWorkbook = %q, want true", got)
	}
	if got := requests[1].Args["WorkbookPath"]; got != filepath.Join(root, cfg.Excel.Path) {
		t.Fatalf("opt-in WorkbookPath = %q, want %q", got, filepath.Join(root, cfg.Excel.Path))
	}
}

func TestRunnerDoctorPreservesDotNetBridgeMetadataAndDiagnostics(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"doctor","logs":[],"error":null,"bridge":{"name":"xlflow-excel-bridge","version":"0.1.0","protocol_version":1,"runtime":".NET 8.0","architecture":"X64"},"diagnostics":{"requested_bridge":"dotnet","selected_bridge":"dotnet","fallback":false,"legacy":false,"protocol_version":1,"runtime":{"os":"Windows 11","process_architecture":"X64","dotnet_runtime":".NET 8.0"},"excel":{"com_activation":true,"version":"16.0","build":"12345","vbide_access":true,"automation_security":1,"trust_vba_access":null}}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Doctor(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}

	bridge, ok := env.Bridge.(map[string]any)
	if !ok {
		t.Fatalf("Bridge = %T, want map[string]any", env.Bridge)
	}
	if bridge["name"] != "xlflow-excel-bridge" {
		t.Fatalf("Bridge.name = %v, want xlflow-excel-bridge", bridge["name"])
	}
	if bridge["protocol_version"] != float64(1) {
		t.Fatalf("Bridge.protocol_version = %v, want 1", bridge["protocol_version"])
	}

	diagnostics, ok := env.Diagnostics.(map[string]any)
	if !ok {
		t.Fatalf("Diagnostics = %T, want map[string]any", env.Diagnostics)
	}
	if diagnostics["selected_bridge"] != "dotnet" {
		t.Fatalf("Diagnostics.selected_bridge = %v, want dotnet", diagnostics["selected_bridge"])
	}
	if diagnostics["requested_bridge"] != "dotnet" {
		t.Fatalf("Diagnostics.requested_bridge = %v, want dotnet", diagnostics["requested_bridge"])
	}
	if diagnostics["fallback"] != false {
		t.Fatalf("Diagnostics.fallback = %v, want false", diagnostics["fallback"])
	}
	if diagnostics["legacy"] != false {
		t.Fatalf("Diagnostics.legacy = %v, want false", diagnostics["legacy"])
	}
	if diagnostics["protocol_version"] != float64(1) {
		t.Fatalf("Diagnostics.protocol_version = %v, want 1", diagnostics["protocol_version"])
	}

	runtime, ok := diagnostics["runtime"].(map[string]any)
	if !ok {
		t.Fatalf("Diagnostics.runtime = %T, want map[string]any", diagnostics["runtime"])
	}
	if runtime["process_architecture"] != "X64" {
		t.Fatalf("Diagnostics.runtime.process_architecture = %v, want X64", runtime["process_architecture"])
	}
	if runtime["dotnet_runtime"] != ".NET 8.0" {
		t.Fatalf("Diagnostics.runtime.dotnet_runtime = %v, want .NET 8.0", runtime["dotnet_runtime"])
	}

	excel, ok := diagnostics["excel"].(map[string]any)
	if !ok {
		t.Fatalf("Diagnostics.excel = %T, want map[string]any", diagnostics["excel"])
	}
	if excel["com_activation"] != true {
		t.Fatalf("Diagnostics.excel.com_activation = %v, want true", excel["com_activation"])
	}
	if excel["version"] != "16.0" {
		t.Fatalf("Diagnostics.excel.version = %v, want 16.0", excel["version"])
	}
	if excel["build"] != "12345" {
		t.Fatalf("Diagnostics.excel.build = %v, want 12345", excel["build"])
	}
	if excel["vbide_access"] != true {
		t.Fatalf("Diagnostics.excel.vbide_access = %v, want true", excel["vbide_access"])
	}
	if excel["automation_security"] != float64(1) {
		t.Fatalf("Diagnostics.excel.automation_security = %v, want 1", excel["automation_security"])
	}
	if excel["trust_vba_access"] != nil {
		t.Fatalf("Diagnostics.excel.trust_vba_access = %v, want nil", excel["trust_vba_access"])
	}
}

func TestRunnerDoctorPreservesDotNetBridgeStructuredErrorMetadata(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"doctor","logs":[],"error":{"code":"excel_com_failure","message":"COM activation failed: HRESULT 0x80040154: Class not registered","source":"xlflow-excel-bridge","number":-2147221164,"phase":"doctor","h_result":"0x80040154","details":{"source":"test","stack_trace":"   at Xlflow.ExcelBridge.Diagnostics.ExcelDiagnostics.Probe()"}},"bridge":{"name":"xlflow-excel-bridge","version":"0.1.0","protocol_version":1,"runtime":".NET 8.0","architecture":"X64"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.Doctor(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Error == nil {
		t.Fatal("Error is nil, want structured error")
	}
	if env.Error.Code != "excel_com_failure" {
		t.Fatalf("Error.Code = %q, want excel_com_failure", env.Error.Code)
	}
	if env.Error.Source != "xlflow-excel-bridge" {
		t.Fatalf("Error.Source = %q, want xlflow-excel-bridge", env.Error.Source)
	}
	if env.Error.Phase != "doctor" {
		t.Fatalf("Error.Phase = %q, want doctor", env.Error.Phase)
	}
	if env.Error.Number != -2147221164 {
		t.Fatalf("Error.Number = %d, want -2147221164", env.Error.Number)
	}
	if env.Error.HResult != "0x80040154" {
		t.Fatalf("Error.HResult = %q, want 0x80040154", env.Error.HResult)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok {
		t.Fatalf("Error.Details = %T, want map[string]any", env.Error.Details)
	}
	if details["source"] != "test" {
		t.Fatalf("Error.Details.source = %v, want test", details["source"])
	}

	bridge, ok := env.Bridge.(map[string]any)
	if !ok {
		t.Fatalf("Bridge = %T, want map[string]any", env.Bridge)
	}
	if bridge["name"] != "xlflow-excel-bridge" {
		t.Fatalf("Bridge.name = %v, want xlflow-excel-bridge", bridge["name"])
	}
}

func TestRunnerDoctorAutoDoesNotFallbackWhenDotNetMissing(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				err: &excelbridge.Error{
					Kind:    excelbridge.ErrorDotNetMissing,
					Message: ".NET bridge missing",
				},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"doctor","logs":[],"bridge":{"host":"powershell.exe","edition":"Desktop","version":"5.1.22621.2506"}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.Doctor(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 || powerShellCalls != 0 {
		t.Fatalf("dotnet calls = %d, powershell calls = %d, want 1 and 0", dotNetCalls, powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "bridge_not_available" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerRejectsExplicitPowerShellBridgeMode(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var calls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:      string(mode),
			supports:  true,
			callCount: &calls,
			response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"doctor","logs":[],"bridge":{"host":"powershell.exe","edition":"Desktop","version":"5.1.22621.2506"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "powershell"}.Doctor(config.Default())
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitConfig {
		t.Fatalf("exit code = %d, want %d", code, output.ExitConfig)
	}
	if calls != 0 {
		t.Fatalf("provider calls = %d, want 0", calls)
	}
	if env.Error == nil || env.Error.Code != "bridge_mode_invalid" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerRejectsEnvPowerShellBridgeMode(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	t.Setenv(excelbridge.EnvBridge, "powershell")

	var calls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:      string(mode),
			supports:  true,
			callCount: &calls,
			response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[]}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir()}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitConfig {
		t.Fatalf("exit code = %d, want %d", code, output.ExitConfig)
	}
	if calls != 0 {
		t.Fatalf("provider calls = %d, want 0", calls)
	}
	if env.Error == nil || env.Error.Code != "bridge_mode_invalid" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerRejectsConfigPowerShellBridgeMode(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var calls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:      string(mode),
			supports:  true,
			callCount: &calls,
			response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[]}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), ConfigBridgeMode: "powershell"}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitConfig {
		t.Fatalf("exit code = %d, want %d", code, output.ExitConfig)
	}
	if calls != 0 {
		t.Fatalf("provider calls = %d, want 0", calls)
	}
	if env.Error == nil || env.Error.Code != "bridge_mode_invalid" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDotNetExplicitModeDoesNotFallBackOnDecodeError(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`not valid json {{{`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[]}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code == output.ExitSuccess {
		t.Fatalf("expected error exit, got ExitSuccess")
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell must not be called in explicit --bridge dotnet mode, got %d calls", powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "invalid_script_json" {
		t.Fatalf("expected invalid_script_json error, got %+v", env.Error)
	}
}

func TestRunnerAutoModeAttemptsDotNetForUnsupportedCommand(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  false,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"run","logs":[],"result":{"stdout":"","stderr":"","exit_code":0,"duration_ms":1}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.Run(config.Default(), RunOptions{
		Macro: "Module1.Main",
	})
	if err != nil {
		t.Fatal(err)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Error == nil || env.Error.Code != "bridge_protocol_mismatch" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerAutoModeDoesNotFallBackToPowerShellOnDecodeError(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`not valid json {{{`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[{"pid":9999,"has_workbook":false}]}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "invalid_script_json" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerAutoModeDoesNotFallBackToPowerShellWhenDotNetProtocolVersionMissing(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[]}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"process","logs":[],"process":[{"pid":7777,"has_workbook":true}]}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "bridge_protocol_mismatch" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerTypeDBImportUsesDotNetAndMapsTypeDBPayload(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var requests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     string(mode),
			supports: mode == excelbridge.ModeDotNet,
			requests: &requests,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"type-db-import","logs":["generated"],"type_db":{"dir":"C:\\typedb","generated_files":["C:\\typedb\\excel.generated.json"]}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.TypeDBImport(TypeDBImportOptions{
		OutputDir:        `C:\typedb`,
		GeneratorVersion: "1.2.3",
		Libraries:        []string{"excel"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want success: %+v", code, env.Error)
	}
	if len(requests) != 1 || requests[0].Command != "type-db-import" {
		t.Fatalf("requests = %+v", requests)
	}
	if requests[0].Args["OutputDir"] != `C:\typedb` || requests[0].Args["GeneratorVersion"] != "1.2.3" || requests[0].Args["Libraries"] != "excel" {
		t.Fatalf("unexpected request args: %+v", requests[0].Args)
	}
	typeDB, ok := env.TypeDB.(map[string]any)
	if !ok || typeDB["dir"] != `C:\typedb` {
		t.Fatalf("type_db not mapped: %+v", env.TypeDB)
	}
}

func TestBuildUIButtonAddScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet:       "Menu",
		Cell:        "B2",
		Text:        "Run",
		Macro:       "Main.Run",
		ID:          "run",
		Width:       160,
		Height:      40,
		CreateSheet: true,
		VerifyMacro: true,
	})
	if args["Action"] != "add" {
		t.Fatalf("action = %q, want add", args["Action"])
	}
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["Sheet"] != "Menu" || args["Cell"] != "B2" || args["Text"] != "Run" || args["Macro"] != "Main.Run" || args["Id"] != "run" {
		t.Fatalf("unexpected args: %+v", args)
	}
	if args["Width"] != "160" || args["Height"] != "40" || args["CreateSheet"] != "true" || args["VerifyMacro"] != "true" {
		t.Fatalf("unexpected numeric/bool args: %+v", args)
	}
}

func TestBuildExportImageScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildExportImageScriptArgs(root, cfg, ExportImageOptions{
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Sheet:        "QR",
		Range:        "A1:AE31",
		OutputDir:    filepath.Join("artifacts", "images"),
		Name:         "qr-demo",
		Format:       "png",
		Overwrite:    true,
		Session:      true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["Sheet"] != "QR" || args["RangeAddress"] != "A1:AE31" {
		t.Fatalf("unexpected sheet/range args: %+v", args)
	}
	if args["OutputPath"] != filepath.Join(root, "artifacts", "images", "qr-demo.png") {
		t.Fatalf("output path = %q", args["OutputPath"])
	}
	if args["ImageFormat"] != "png" || args["Overwrite"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected export args: %+v", args)
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("metadata path = %q", args["MetadataPath"])
	}
}

func TestBuildFormExportImageScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildFormExportImageScriptArgs(root, cfg, FormExportImageOptions{
		Name:        "UserForm1",
		OutPath:     filepath.Join("artifacts", "forms", "UserForm1.png"),
		Initializer: "InitializeForm",
		Overwrite:   true,
		Session:     true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["FormName"] != "UserForm1" || args["Initializer"] != "InitializeForm" {
		t.Fatalf("unexpected form args: %+v", args)
	}
	if args["OutputPath"] != filepath.Join(root, "artifacts", "forms", "UserForm1.png") {
		t.Fatalf("output path = %q", args["OutputPath"])
	}
	if args["Overwrite"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected overwrite/session args: %+v", args)
	}
}

func TestBuildEditCellScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	value := "ABC123"
	args, err := buildEditCellScriptArgs(root, cfg, EditCellOptions{
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Sheet:        "Input",
		Cell:         "B2",
		Value:        &value,
		Events:       EditEventOn,
		Session:      true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Action"] != "cell" || args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("unexpected edit cell args: %+v", args)
	}
	if args["Sheet"] != "Input" || args["Cell"] != "B2" || args["Value"] != "ABC123" || args["Events"] != "on" {
		t.Fatalf("unexpected edit cell payload: %+v", args)
	}
	if args["UseSession"] != "true" || args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("unexpected session args: %+v", args)
	}
}

func TestBuildEditRangeRowsAndColumnsScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	rangeArgs, err := buildEditRangeScriptArgs(root, cfg, EditRangeOptions{
		Sheet:   "QR",
		Range:   "A1:AE31",
		Fill:    "#FFFFFF",
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if rangeArgs["Action"] != "range" || rangeArgs["RangeAddress"] != "A1:AE31" || rangeArgs["Fill"] != "#FFFFFF" {
		t.Fatalf("unexpected edit range args: %+v", rangeArgs)
	}
	rowsArgs := buildEditRowsScriptArgs(root, cfg, EditRowsOptions{
		Sheet:   "QR",
		Rows:    "1:31",
		Height:  12,
		Session: true,
	})
	if rowsArgs["Action"] != "rows" || rowsArgs["Rows"] != "1:31" || rowsArgs["Height"] != "12" {
		t.Fatalf("unexpected edit rows args: %+v", rowsArgs)
	}
	columnArgs := buildEditColumnsScriptArgs(root, cfg, EditColumnsOptions{
		Sheet:   "QR",
		Columns: "A:AE",
		Width:   2.2,
		Session: true,
	})
	if columnArgs["Action"] != "columns" || columnArgs["Columns"] != "A:AE" || columnArgs["Width"] != "2.2" {
		t.Fatalf("unexpected edit columns args: %+v", columnArgs)
	}
}

func TestBuildEditScriptArgsRejectInvalidCombinations(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	value := "ABC123"
	formula := "=A1+B1"
	if _, err := buildEditCellScriptArgs(root, cfg, EditCellOptions{
		Sheet:   "Input",
		Cell:    "B2",
		Value:   &value,
		Formula: &formula,
		Session: true,
	}); err == nil || !strings.Contains(err.Error(), "exactly one") {
		t.Fatalf("expected cell mutation conflict, got %v", err)
	}
	if _, err := buildEditRangeScriptArgs(root, cfg, EditRangeOptions{
		Sheet:   "QR",
		Range:   "A1:B2",
		Fill:    "#FFFFFF",
		Clear:   "all",
		Session: true,
	}); err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected range mutation conflict, got %v", err)
	}
}

func TestResolveExportImageOutputDefaultPath(t *testing.T) {
	root := t.TempDir()
	resolved, err := resolveExportImageOutput(root, filepath.Join("build", "Book.xlsm"), ExportImageOptions{
		Sheet:  "Sheet 1",
		Range:  "A1:AE31",
		Format: "png",
	})
	if err != nil {
		t.Fatal(err)
	}
	if !resolved.Default {
		t.Fatal("expected default output path")
	}
	wantDir := filepath.Join(root, ".xlflow", "artifacts", "images", "Book")
	if filepath.Dir(resolved.Path) != wantDir {
		t.Fatalf("dir = %q, want %q", filepath.Dir(resolved.Path), wantDir)
	}
	base := filepath.Base(resolved.Path)
	if !strings.HasPrefix(base, "Sheet_1_A1-AE31_") || !strings.HasSuffix(base, ".png") {
		t.Fatalf("unexpected default filename %q", base)
	}
}

func TestResolveFormExportImageOutputRejectsUnsupportedExtension(t *testing.T) {
	_, err := resolveFormExportImageOutput(t.TempDir(), FormExportImageOptions{OutPath: "artifacts\\UserForm1.webp"})
	if err == nil || !strings.Contains(err.Error(), "supported formats: png") {
		t.Fatalf("expected unsupported format error, got %v", err)
	}
}

func TestResolveFormExportImageOutputRejectsMissingExtension(t *testing.T) {
	_, err := resolveFormExportImageOutput(t.TempDir(), FormExportImageOptions{OutPath: "artifacts\\UserForm1"})
	if err == nil || !strings.Contains(err.Error(), `unsupported image format ""; supported formats: png`) {
		t.Fatalf("expected missing extension error, got %v", err)
	}
}

func TestNormalizeExportImagePathRejectsUnsupportedExtension(t *testing.T) {
	_, _, err := normalizeExportImagePath(t.TempDir(), "artifacts\\qr.webp", "png")
	if err == nil || !strings.Contains(err.Error(), "supported formats: png") {
		t.Fatalf("expected unsupported format error, got %v", err)
	}
}

func TestRunnerExportImageReturnsValidationFailureForUnsupportedFormat(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.ExportImage(config.Default(), ExportImageOptions{
		Sheet:  "QR",
		Range:  "A1:B2",
		Format: "webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerExportImageReturnsValidationFailureForUnsupportedExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.ExportImage(config.Default(), ExportImageOptions{
		Sheet:   "QR",
		Range:   "A1:B2",
		OutPath: "artifacts\\qr.webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerFormExportImageReturnsValidationFailureForUnsupportedExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.FormExportImage(config.Default(), FormExportImageOptions{
		Name:    "UserForm1",
		OutPath: "artifacts\\UserForm1.webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerFormExportImageReturnsValidationFailureForMissingExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.FormExportImage(config.Default(), FormExportImageOptions{
		Name:    "UserForm1",
		OutPath: "artifacts\\UserForm1",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerInspectFormNormalizesCommandName(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$WorkbookPath,[string]$FormName,[string]$Basis,[string]$Initializer,[string]$Visible,[string]$UseSession,[string]$MetadataPath)
@{ status = "ok"; command = "inspect-form"; logs = @(); error = $null } | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "inspect-form.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.InspectForm(config.Default(), InspectFormOptions{Name: "UserForm1", Basis: "designer"})
	if err != nil {
		t.Fatal(err)
	}
	if env.Command != "inspect" {
		t.Fatalf("command = %q, want inspect", env.Command)
	}
}

func TestOutputFileExistsIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "output_file_exists", Message: "exists"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(output_file_exists) = %d, want %d", got, output.ExitValidation)
	}
}

func TestUnsupportedImageFormatIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "unsupported_image_format", Message: "bad format"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(unsupported_image_format) = %d, want %d", got, output.ExitValidation)
	}
}

func TestFormExportImageFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"runtime_form_load_failed", "form_initializer_failed", "window_not_found", "image_capture_failed"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestUIValidationFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"sheet_not_found", "button_not_found", "ui_button_args_invalid"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestEditValidationFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"session_required", "invalid_color", "invalid_cell_address", "invalid_row_selector", "invalid_column_selector", "vba_event_error"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestRunnerEditReturnsConfigFailureForInvalidMutations(t *testing.T) {
	value := "ABC123"
	formula := "=A1+B1"
	env, code, err := Runner{RootDir: t.TempDir()}.EditCell(config.Default(), EditCellOptions{
		Sheet:   "Input",
		Cell:    "B2",
		Value:   &value,
		Formula: &formula,
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitConfig {
		t.Fatalf("exit code = %d, want %d", code, output.ExitConfig)
	}
	if env.Error == nil || env.Error.Code != "edit_args_invalid" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestTestFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"test_failed", "no_tests_found", "test_not_found", "duplicate_test_name"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestRunnerTestReturnsEnvironmentFailureOnNonWindows(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("non-Windows behavior")
	}
	env, code, err := Runner{RootDir: t.TempDir()}.Test(config.Default(), "")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Command != "test" {
		t.Fatalf("command = %q, want test", env.Command)
	}
}

func TestBuildRunScriptArgsSerializesArgumentsAndSaveAs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:        "Report.Generate",
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Args: []RunArgument{
			{Type: "string", Value: "fixtures\\sample.xlsx"},
			{Type: "int", Value: "3"},
			{Type: "bool", Value: "true"},
		},
		SaveAs: filepath.Join("build", "Result.xlsm"),
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["MacroName"] != "Report.Generate" {
		t.Fatalf("macro name = %q", args["MacroName"])
	}
	if args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["SaveWorkbook"] != "false" {
		t.Fatalf("save flag = %q", args["SaveWorkbook"])
	}
	if args["Direct"] != "false" || args["UseSession"] != "false" {
		t.Fatalf("unexpected direct/session defaults: %+v", args)
	}
	if args["SaveAsPath"] != filepath.Join(root, "build", "Result.xlsm") {
		t.Fatalf("save-as path = %q", args["SaveAsPath"])
	}
	if args["ModulesDir"] != filepath.Join(root, cfg.Src.Modules) ||
		args["ClassesDir"] != filepath.Join(root, cfg.Src.Classes) ||
		args["FormsDir"] != filepath.Join(root, cfg.Src.Forms) ||
		args["WorkbookDir"] != filepath.Join(root, cfg.Src.Workbook) {
		t.Fatalf("source mapping args were not populated: %+v", args)
	}
	if args["CodeSource"] != cfg.UserForm.CodeSource ||
		args["Folders"] != strconv.FormatBool(cfg.VBA.Folders) ||
		args["FolderAnnotation"] != cfg.VBA.FolderAnnotation ||
		args["DefaultComponentFolders"] != strconv.FormatBool(cfg.VBA.DefaultComponentFolders) {
		t.Fatalf("VBA source mapping settings were not populated: %+v", args)
	}
	wantJSON := `[{"type":"string","value":"fixtures\\sample.xlsx"},{"type":"int","value":"3"},{"type":"bool","value":"true"}]`
	wantJSON64 := base64.StdEncoding.EncodeToString([]byte(wantJSON))
	if args["MacroArgsJSON"] != wantJSON64 {
		t.Fatalf("macro args json base64 = %s, want %s", args["MacroArgsJSON"], wantJSON64)
	}
}

func TestBuildRunScriptArgsPassesFastDirectAndSession(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:               "Main.Run",
		Fast:                true,
		Session:             true,
		SuppressModalErrors: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Direct"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected args: %+v", args)
	}
	if args["SuppressModalErrors"] != "true" {
		t.Fatalf("SuppressModalErrors = %q, want true", args["SuppressModalErrors"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("metadata path = %q", args["MetadataPath"])
	}
}

func TestBuildRunScriptArgsDiagnosticDisablesFastDirect(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:      "Main.Run",
		Fast:       true,
		Diagnostic: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Diagnostic"] != "true" {
		t.Fatalf("Diagnostic = %q, want true", args["Diagnostic"])
	}
	if args["Direct"] != "false" {
		t.Fatalf("Direct = %q, want false for diagnostic fast run", args["Direct"])
	}
}

func TestBuildRunScriptArgsPropagatesSuppressModalErrors(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:               "Main.Run",
		SuppressModalErrors: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["SuppressModalErrors"] != "true" {
		t.Fatalf("SuppressModalErrors = %q, want true", args["SuppressModalErrors"])
	}
}

func TestMacroFailureIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_failed", Message: "boom"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_failed) = %d, want %d", got, output.ExitValidation)
	}
}

func TestMacroDisabledIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_disabled", Message: "disabled"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_disabled) = %d, want %d", got, output.ExitValidation)
	}
}

func TestBuildRunScriptArgsNormalizesNilArgsToEmptyArray(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:        "Sheet1.Main",
		WorkbookPath: "Book.xlsm",
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["MacroArgsJSON"] != base64.StdEncoding.EncodeToString([]byte("[]")) {
		t.Fatalf("macro args json = %q, want base64 of []", args["MacroArgsJSON"])
	}
}

func TestBuildRunScriptArgsPassesRuntimeMode(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:         "Main.Run",
		RuntimeMode:   RuntimeModeAgent,
		RuntimeSource: RuntimeSourceEnvironment,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["RuntimeMode"] != RuntimeModeAgent {
		t.Fatalf("RuntimeMode = %q, want %q", args["RuntimeMode"], RuntimeModeAgent)
	}
	if args["RuntimeSource"] != RuntimeSourceEnvironment {
		t.Fatalf("RuntimeSource = %q, want %q", args["RuntimeSource"], RuntimeSourceEnvironment)
	}
}

func TestBuildRunScriptArgsPassesUIResponses(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro: "Main.Run",
		UIResponses: UIResponses{
			MsgBox:     map[string]string{"confirm-save": "yes"},
			Input:      map[string]string{"customer-name": "John"},
			FileDialog: []FileDialogResponse{{Kind: "get-open", DialogID: "source_files", Values: []string{"C:\\tmp\\a.txt", "C:\\tmp\\b.txt"}}},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if got, want := args["MsgBoxResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"confirm-save":"yes"}`)); got != want {
		t.Fatalf("MsgBoxResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["InputResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"customer-name":"John"}`)); got != want {
		t.Fatalf("InputResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["FileDialogResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`[{"kind":"get-open","dialog_id":"source_files","values":["C:\\tmp\\a.txt","C:\\tmp\\b.txt"]}]`)); got != want {
		t.Fatalf("FileDialogResponsesJSON = %q, want %q", got, want)
	}
}

func TestBuildRunScriptArgsPassesUIStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:    "Main.Run",
		UIStream: UIStreamOptions{Enabled: true, RedactInput: true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["UIStreamEnabled"] != "true" {
		t.Fatalf("UIStreamEnabled = %q, want true", args["UIStreamEnabled"])
	}
	if args["UIStreamRedactInput"] != "true" {
		t.Fatalf("UIStreamRedactInput = %q, want true", args["UIStreamRedactInput"])
	}
}

func TestBuildRunScriptArgsPassesDebugStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:       "Main.Run",
		DebugStream: DebugStreamOptions{Enabled: true, PipeName: `\\.\pipe\xlflow-debug-test`},
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["DebugStreamEnabled"] != "true" {
		t.Fatalf("DebugStreamEnabled = %q, want true", args["DebugStreamEnabled"])
	}
	if args["DebugStreamPipeName"] != `\\.\pipe\xlflow-debug-test` {
		t.Fatalf("DebugStreamPipeName = %q, want debug pipe name", args["DebugStreamPipeName"])
	}
}

func TestBuildTestScriptArgsPassesRuntimeMode(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{RuntimeMode: RuntimeModeTest, RuntimeSource: RuntimeSourceCommand})
	if args["RuntimeMode"] != RuntimeModeTest {
		t.Fatalf("RuntimeMode = %q, want %q", args["RuntimeMode"], RuntimeModeTest)
	}
	if args["RuntimeSource"] != RuntimeSourceCommand {
		t.Fatalf("RuntimeSource = %q, want %q", args["RuntimeSource"], RuntimeSourceCommand)
	}
}

func TestBuildTestScriptArgsPassesUIResponses(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{
		UIResponses: UIResponses{
			MsgBox:     map[string]string{"confirm-save": "no"},
			Input:      map[string]string{"customer-name": "Jane"},
			FileDialog: []FileDialogResponse{{Kind: "folder", DialogID: "target_dir", Cancelled: true}},
		},
	})
	if got, want := args["MsgBoxResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"confirm-save":"no"}`)); got != want {
		t.Fatalf("MsgBoxResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["InputResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"customer-name":"Jane"}`)); got != want {
		t.Fatalf("InputResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["FileDialogResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`[{"kind":"folder","dialog_id":"target_dir","cancelled":true}]`)); got != want {
		t.Fatalf("FileDialogResponsesJSON = %q, want %q", got, want)
	}
}

func TestBuildTestScriptArgsPassesUIStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{UIStream: UIStreamOptions{Enabled: true, RedactInput: true}})
	if args["UIStreamEnabled"] != "true" {
		t.Fatalf("UIStreamEnabled = %q, want true", args["UIStreamEnabled"])
	}
	if args["UIStreamRedactInput"] != "true" {
		t.Fatalf("UIStreamRedactInput = %q, want true", args["UIStreamRedactInput"])
	}
}

func TestBuildTestScriptArgsPassesDebugStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{DebugStream: DebugStreamOptions{Enabled: true, PipeName: `\\.\pipe\xlflow-debug-test`}})
	if args["DebugStreamEnabled"] != "true" {
		t.Fatalf("DebugStreamEnabled = %q, want true", args["DebugStreamEnabled"])
	}
	if args["DebugStreamPipeName"] != `\\.\pipe\xlflow-debug-test` {
		t.Fatalf("DebugStreamPipeName = %q, want debug pipe name", args["DebugStreamPipeName"])
	}
}

func TestMergeUIResultAppendsStreamEvents(t *testing.T) {
	existing := map[string]any{"events": []any{map[string]any{"kind": "msgbox", "dialog_id": "existing"}}}
	merged := mergeUIResult(existing, []map[string]any{{"kind": "inputbox", "dialog_id": "customer-name"}})
	mergedMap, ok := merged.(map[string]any)
	if !ok {
		t.Fatalf("merged UI = %#v", merged)
	}
	events, ok := mergedMap["events"].([]any)
	if !ok || len(events) != 2 {
		t.Fatalf("merged events = %#v", mergedMap["events"])
	}
}

func TestMergeDebugResultPreservesExistingEvents(t *testing.T) {
	existing := map[string]any{
		"events":    []any{map[string]any{"message": "existing"}},
		"count":     4,
		"truncated": true,
	}
	streamed := map[string]any{
		"events": []any{map[string]any{"message": "streamed"}},
		"count":  3,
	}

	merged := mergeDebugResult(existing, streamed)
	mergedMap, ok := merged.(map[string]any)
	if !ok {
		t.Fatalf("merged debug = %#v", merged)
	}

	events := debugEventList(mergedMap["events"])
	if len(events) != 2 {
		t.Fatalf("merged debug events = %#v", mergedMap["events"])
	}
	if got := events[0]["message"]; got != "existing" {
		t.Fatalf("first merged debug event = %#v, want existing", got)
	}
	if got := events[1]["message"]; got != "streamed" {
		t.Fatalf("second merged debug event = %#v, want streamed", got)
	}
	if got := mergedMap["count"]; got != 4 {
		t.Fatalf("merged debug count = %#v, want 4", got)
	}
	if got := mergedMap["truncated"]; got != true {
		t.Fatalf("merged debug truncated = %#v, want true", got)
	}
}

func TestVBACompileFailedIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "vba_compile_failed", Message: "compile failed"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(vba_compile_failed) = %d, want %d", got, output.ExitValidation)
	}
}

func TestMacroNotFoundIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_not_found", Message: "missing"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_not_found) = %d, want %d", got, output.ExitValidation)
	}
}

func TestDuplicateModuleNameIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "duplicate_module_name", Message: "duplicate"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(duplicate_module_name) = %d, want %d", got, output.ExitValidation)
	}
}

func TestPullScriptArgsIncludeFolderConfig(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildPullScriptArgs(root, cfg, SessionCommandOptions{})
	if args["Folders"] != "true" || args["FolderAnnotation"] != "update" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder config args: %+v", args)
	}
	if args["CodeSource"] != "sidecar" {
		t.Fatalf("CodeSource = %q", args["CodeSource"])
	}
}

func TestListFormsScriptArgsIncludeFolderAndSessionConfig(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildListFormsScriptArgs(root, cfg, SessionCommandOptions{Session: true})
	if args["Action"] != "forms" {
		t.Fatalf("Action = %q, want forms", args["Action"])
	}
	if args["ProjectRoot"] != root {
		t.Fatalf("ProjectRoot = %q, want %q", args["ProjectRoot"], root)
	}
	if args["FormsDir"] != filepath.Join(root, "src", "forms") {
		t.Fatalf("FormsDir = %q", args["FormsDir"])
	}
	if args["FolderAnnotation"] != "update" || args["Folders"] != "true" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder config args: %+v", args)
	}
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestInspectFormScriptArgsIncludeSessionAndInitializer(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildInspectFormScriptArgs(root, cfg, InspectFormOptions{
		Name:           "UserForm1",
		Basis:          "runtime",
		Initializer:    "InitializeForm",
		StrictDesigner: true,
		Session:        true,
	})
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("WorkbookPath = %q", args["WorkbookPath"])
	}
	if args["FormName"] != "UserForm1" {
		t.Fatalf("FormName = %q", args["FormName"])
	}
	if args["Basis"] != "runtime" {
		t.Fatalf("Basis = %q", args["Basis"])
	}
	if args["Initializer"] != "InitializeForm" {
		t.Fatalf("Initializer = %q", args["Initializer"])
	}
	if args["StrictDesigner"] != "true" {
		t.Fatalf("StrictDesigner = %q", args["StrictDesigner"])
	}
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestFormWriteScriptArgsIncludeSpecPayloadAndSessionFlags(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildFormWriteScriptArgs(root, cfg, FormWriteOptions{
		Action:   "apply",
		SpecPath: "src/forms/UserForm1.form.yaml",
		Spec: forms.FormSpec{
			SchemaVersion: 1,
			Kind:          "xlflow.userform",
			Basis:         "designer",
			Form:          forms.FormSpecForm{Name: "UserForm1"},
			Controls: []forms.FormSpecControl{{
				Name: "txtCustomer",
				Type: "TextBox",
			}},
			Warnings: []forms.FormSpecWarning{},
		},
		Session: true,
		NoSave:  true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Action"] != "apply" {
		t.Fatalf("Action = %q", args["Action"])
	}
	if args["SpecPath"] != "src/forms/UserForm1.form.yaml" {
		t.Fatalf("SpecPath = %q", args["SpecPath"])
	}
	if args["FormsDir"] != filepath.Join(root, "src", "forms") {
		t.Fatalf("FormsDir = %q", args["FormsDir"])
	}
	if args["CodeSource"] != "sidecar" {
		t.Fatalf("CodeSource = %q", args["CodeSource"])
	}
	if args["FolderAnnotation"] != "update" || args["Folders"] != "true" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder args: %+v", args)
	}
	if args["UseSession"] != "true" || args["NoSave"] != "true" {
		t.Fatalf("unexpected session flags: %+v", args)
	}
	if args["SpecJson64"] == "" {
		t.Fatalf("SpecJson64 should be set: %+v", args)
	}
}

func TestBuildUIButtonAddScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet: "Menu",
		Cell:  "B2",
		Text:  "Run",
		Macro: "Main.Run",
	})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonListScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu"})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonRemoveScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run"})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonAddScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet:   "Menu",
		Cell:    "B2",
		Text:    "Run",
		Macro:   "Main.Run",
		Session: true,
	})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonListScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu", Session: true})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonRemoveScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run", Session: true})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonAddScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet: "Menu",
		Cell:  "B2",
		Text:  "Run",
		Macro: "Main.Run",
	})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q (always set for auto-detection)", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonListScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu"})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
}

func TestBuildUIButtonRemoveScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run"})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
}

func TestProcessArgsInvalidIsConfigFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "process_args_invalid", Message: "invalid process args"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitConfig {
		t.Fatalf("exitCodeForScriptResult(process_args_invalid) = %d, want %d", got, output.ExitConfig)
	}
}

func TestProcessNotFoundIsConfigFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "process_not_found", Message: "process not found"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitConfig {
		t.Fatalf("exitCodeForScriptResult(process_not_found) = %d, want %d", got, output.ExitConfig)
	}
}

func TestPullArgsInvalidIsConfigFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "pull_args_invalid", Message: "invalid pull args"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitConfig {
		t.Fatalf("exitCodeForScriptResult(pull_args_invalid) = %d, want %d", got, output.ExitConfig)
	}
}

func TestProcessEnvironmentFailureCodesAreEnvironmentFailures(t *testing.T) {
	for _, code := range []string{"process_enumeration_failed", "process_termination_failed", "process_cleanup_failed"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitEnvironment {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitEnvironment)
			}
		})
	}
}

func TestRunnerDotNetPullResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"pull","logs":["exported 3 VBA component(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true},"source":{"modules_dir":"C:\\temp\\src\\modules","classes_dir":"C:\\temp\\src\\classes","forms_dir":"C:\\temp\\src\\forms","workbook_dir":"C:\\temp\\src\\workbook"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PullWithOptions(config.Default(), SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil {
		t.Fatal("Target is nil")
	}
	if env.Session == nil {
		t.Fatal("Session is nil")
	}
	if env.Workbook == nil {
		t.Fatal("Workbook is nil")
	}
	if env.Source == nil {
		t.Fatal("Source is nil")
	}
	target, ok := env.Target.(map[string]interface{})
	if !ok || target["kind"] != "live_session" {
		t.Fatalf("unexpected Target: %+v", env.Target)
	}
	session, ok := env.Session.(map[string]interface{})
	if !ok || session["active"] != true {
		t.Fatalf("unexpected Session: %+v", env.Session)
	}
	source, ok := env.Source.(map[string]interface{})
	if !ok || source["modules_dir"] == nil {
		t.Fatalf("unexpected Source: %+v", env.Source)
	}
}

func TestRunnerDotNetPushResponsePreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"push","logs":["imported 3 source file(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true},"backup":{"id":"push_20260101T120000","path":"C:\\temp\\.xlflow\\backups\\Book_20260101T120000.xlsm","reason":"before-push","mode":"always"},"source":{"changed_only":false,"changed":true,"state":"C:\\temp\\.xlflow\\state\\push.json"},"push_diagnostic":{"kind":"compile","location":{"source_path":"src/modules/Main.bas","line":6,"text":"  x ="}}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PushWithOptions(config.Default(), PushOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil {
		t.Fatal("Target is nil")
	}
	if env.Session == nil {
		t.Fatal("Session is nil")
	}
	if env.Workbook == nil {
		t.Fatal("Workbook is nil")
	}
	if env.Backup == nil {
		t.Fatal("Backup is nil")
	}
	if env.Source == nil {
		t.Fatal("Source is nil")
	}
	if env.PushDiagnostic == nil {
		t.Fatal("PushDiagnostic is nil")
	}
	target, ok := env.Target.(map[string]interface{})
	if !ok || target["kind"] != "live_session" {
		t.Fatalf("unexpected Target: %+v", env.Target)
	}
	session, ok := env.Session.(map[string]interface{})
	if !ok || session["active"] != true {
		t.Fatalf("unexpected Session: %+v", env.Session)
	}
	backup, ok := env.Backup.(map[string]interface{})
	if !ok || backup["mode"] != "always" {
		t.Fatalf("unexpected Backup: %+v", env.Backup)
	}
	source, ok := env.Source.(map[string]interface{})
	if !ok || source["changed_only"] != false {
		t.Fatalf("unexpected Source: %+v", env.Source)
	}
	diag, ok := env.PushDiagnostic.(map[string]interface{})
	if !ok || diag["kind"] != "compile" {
		t.Fatalf("unexpected PushDiagnostic: %+v", env.PushDiagnostic)
	}
}

func TestRunnerDotNetMacrosUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"macros","logs":["discovered 2 macro(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"dirty":true,"needs_save":true},"default_entry":"Module1.Main","macros":[{"module":"Module1","name":"Main","qualified_name":"Module1.Main","kind":"standard","runnable":true},{"module":"Sheet1","name":"Calculate","qualified_name":"Sheet1.Calculate","kind":"document","runnable":false,"reason_not_runnable":"requires workbook event context"}],"suggestions":[{"code":"save_session","message":"Run xlflow save --session to persist workbook changes."}]}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	cfg := config.Default()
	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.MacrosWithOptions(cfg, MacrosOptions{
		Session:      true,
		Entry:        "Module1.Main",
		RunnableOnly: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["UseSession"]; got != "true" {
		t.Fatalf("UseSession = %q, want true", got)
	}
	if got := dotNetRequests[0].Args["Entry"]; got != "Module1.Main" {
		t.Fatalf("Entry = %q, want Module1.Main", got)
	}
	if got := dotNetRequests[0].Args["RunnableOnly"]; got != "true" {
		t.Fatalf("RunnableOnly = %q, want true", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Macros == nil {
		t.Fatalf("expected macros envelope fields to be populated: %+v", env)
	}
	if env.DefaultEntry != "Module1.Main" {
		t.Fatalf("DefaultEntry = %q, want Module1.Main", env.DefaultEntry)
	}
	macros, ok := env.Macros.([]interface{})
	if !ok || len(macros) != 2 {
		t.Fatalf("Macros = %#v, want 2 entries", env.Macros)
	}
	first, ok := macros[0].(map[string]interface{})
	if !ok || first["qualified_name"] != "Module1.Main" {
		t.Fatalf("unexpected first macro payload: %#v", macros[0])
	}
}

func TestRunnerDotNetRunUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	root := t.TempDir()
	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"run","logs":["ran Module1.Main in 42ms","wrote workbook copy to C:\\temp\\out\\Result.xlsm","SAVE REQUIRED: live workbook is newer than disk; run xlflow save before session stop"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true,"save_as":"C:\\temp\\out\\Result.xlsm"},"macro":{"name":"Module1.Main","duration_ms":42,"arguments":[{"type":"string","value":"hello"},{"type":"int","value":"7"},{"type":"bool","value":"true"}]},"runtime":{"mode":"headless","source":"command","injected":true},"run_diagnostic":{"kind":"runtime","location":{"module":"Module1","procedure":"Main","line":12}},"suggestions":[{"code":"save_session","message":"Run xlflow save --session before session stop."}]}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	cfg := config.Default()
	env, code, err := Runner{RootDir: root, BridgeMode: "dotnet"}.Run(cfg, RunOptions{
		Macro:   "Module1.Main",
		Session: true,
		SaveAs:  filepath.Join("out", "Result.xlsm"),
		Args: []RunArgument{
			{Type: "string", Value: "hello"},
			{Type: "int", Value: "7"},
			{Type: "bool", Value: "true"},
		},
		RuntimeMode:   RuntimeModeHeadless,
		RuntimeSource: RuntimeSourceCommand,
		Timeout:       5 * time.Minute,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["MacroName"]; got != "Module1.Main" {
		t.Fatalf("MacroName = %q, want Module1.Main", got)
	}
	if got := dotNetRequests[0].Args["UseSession"]; got != "true" {
		t.Fatalf("UseSession = %q, want true", got)
	}
	if got := dotNetRequests[0].Args["SaveAsPath"]; got != filepath.Join(root, "out", "Result.xlsm") {
		t.Fatalf("SaveAsPath = %q, want %q", got, filepath.Join(root, "out", "Result.xlsm"))
	}
	if got := dotNetRequests[0].Args["RuntimeMode"]; got != RuntimeModeHeadless {
		t.Fatalf("RuntimeMode = %q, want %q", got, RuntimeModeHeadless)
	}
	if got := dotNetRequests[0].Args["RuntimeSource"]; got != RuntimeSourceCommand {
		t.Fatalf("RuntimeSource = %q, want %q", got, RuntimeSourceCommand)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Macro == nil || env.Runtime == nil || env.RunDiagnostic == nil {
		t.Fatalf("expected run envelope fields to be populated: %+v", env)
	}
	macro, ok := env.Macro.(map[string]interface{})
	if !ok || macro["name"] != "Module1.Main" {
		t.Fatalf("unexpected macro payload: %#v", env.Macro)
	}
	workbook, ok := env.Workbook.(map[string]interface{})
	if !ok || workbook["save_as"] != `C:\temp\out\Result.xlsm` {
		t.Fatalf("unexpected workbook payload: %#v", env.Workbook)
	}
	// Verify save-as session contract: live workbook remains dirty after SaveCopyAs.
	if got := workbook["saved"]; got != false {
		t.Fatalf("workbook.saved = %v, want false (SaveCopyAs does not save original)", got)
	}
	if got := workbook["dirty"]; got != true {
		t.Fatalf("workbook.dirty = %v, want true (live workbook remains dirty after SaveCopyAs)", got)
	}
	if got := workbook["needs_save"]; got != true {
		t.Fatalf("workbook.needs_save = %v, want true (live workbook needs save after SaveCopyAs)", got)
	}
	session, ok := env.Session.(map[string]interface{})
	if !ok {
		t.Fatalf("unexpected session payload: %#v", env.Session)
	}
	if got := session["dirty"]; got != true {
		t.Fatalf("session.dirty = %v, want true", got)
	}
	if got := session["save_required"]; got != true {
		t.Fatalf("session.save_required = %v, want true", got)
	}
	if got := session["source_of_truth"]; got != "live_workbook" {
		t.Fatalf("session.source_of_truth = %v, want live_workbook", got)
	}
}

func TestRunnerDotNetTestUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	root := t.TempDir()
	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"test","logs":["discovered 2 tests","ran 2 tests"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true},"runtime":{"mode":"test","source":"command","injected":true},"tests":{"summary":{"total":2,"passed":2,"failed":0},"items":[{"module":"SpecTests","name":"AddsNumbers","status":"passed"},{"module":"SpecTests","name":"HandlesTags","status":"passed"}]},"ui":{"events":[{"kind":"msgbox","dialog_id":"confirm-save"}]},"debug":{"count":1,"events":[{"message":"Immediate output"}]}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	cfg := config.Default()
	env, code, err := Runner{RootDir: root, BridgeMode: "dotnet"}.TestWithOptions(cfg, "SpecTests", TestOptions{
		Session:       true,
		RuntimeMode:   RuntimeModeTest,
		RuntimeSource: RuntimeSourceCommand,
		ModuleFilter:  "SpecTests",
		TagFilter:     "@smoke",
		UIResponses: UIResponses{
			MsgBox:     map[string]string{"confirm-save": "yes"},
			Input:      map[string]string{"customer-name": "Jane"},
			FileDialog: []FileDialogResponse{{Kind: "folder", DialogID: "target_dir", Cancelled: true}},
		},
		UIStream:    UIStreamOptions{Enabled: true, RedactInput: true},
		DebugStream: DebugStreamOptions{Enabled: true, PipeName: `\\.\pipe\xlflow-debug-test`},
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	req := dotNetRequests[0]
	if got := req.Command; got != "test" {
		t.Fatalf("command = %q, want test", got)
	}
	if got := req.Args["Filter"]; got != "SpecTests" {
		t.Fatalf("Filter = %q, want SpecTests", got)
	}
	if got := req.Args["UseSession"]; got != "true" {
		t.Fatalf("UseSession = %q, want true", got)
	}
	if got := req.Args["RuntimeMode"]; got != RuntimeModeTest {
		t.Fatalf("RuntimeMode = %q, want %q", got, RuntimeModeTest)
	}
	if got := req.Args["RuntimeSource"]; got != RuntimeSourceCommand {
		t.Fatalf("RuntimeSource = %q, want %q", got, RuntimeSourceCommand)
	}
	if got := req.Args["ModuleFilter"]; got != "SpecTests" {
		t.Fatalf("ModuleFilter = %q, want SpecTests", got)
	}
	if got := req.Args["TagFilter"]; got != "@smoke" {
		t.Fatalf("TagFilter = %q, want @smoke", got)
	}
	if got, want := req.Args["MsgBoxResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"confirm-save":"yes"}`)); got != want {
		t.Fatalf("MsgBoxResponsesJSON = %q, want %q", got, want)
	}
	if got, want := req.Args["InputResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"customer-name":"Jane"}`)); got != want {
		t.Fatalf("InputResponsesJSON = %q, want %q", got, want)
	}
	if got, want := req.Args["FileDialogResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`[{"kind":"folder","dialog_id":"target_dir","cancelled":true}]`)); got != want {
		t.Fatalf("FileDialogResponsesJSON = %q, want %q", got, want)
	}
	if got := req.Args["UIStreamEnabled"]; got != "true" {
		t.Fatalf("UIStreamEnabled = %q, want true", got)
	}
	if got := req.Args["UIStreamRedactInput"]; got != "true" {
		t.Fatalf("UIStreamRedactInput = %q, want true", got)
	}
	if got := req.Args["DebugStreamEnabled"]; got != "true" {
		t.Fatalf("DebugStreamEnabled = %q, want true", got)
	}
	if got := req.Args["DebugStreamPipeName"]; got == "" {
		t.Fatal("DebugStreamPipeName = empty, want generated pipe name")
	}
	if got := req.Args["DebugStreamPipeName"]; got == `\\.\pipe\xlflow-debug-test` {
		t.Fatalf("DebugStreamPipeName = %q, want generated session pipe name instead of caller placeholder", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Runtime == nil || env.Tests == nil || env.UI == nil || env.Debug == nil {
		t.Fatalf("expected test envelope fields to be populated: %+v", env)
	}
	testsPayload, ok := env.Tests.(map[string]interface{})
	if !ok {
		t.Fatalf("unexpected tests payload: %#v", env.Tests)
	}
	summary, ok := testsPayload["summary"].(map[string]interface{})
	if !ok || summary["passed"] != float64(2) {
		t.Fatalf("unexpected tests.summary payload: %#v", testsPayload["summary"])
	}
}

func TestRunnerAutoBridgeUsesDotNetForPull(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"pull","logs":["exported 1 VBA component(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true},"source":{"modules_dir":"C:\\temp\\src\\modules","classes_dir":"C:\\temp\\src\\classes","forms_dir":"C:\\temp\\src\\forms","workbook_dir":"C:\\temp\\src\\workbook"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"pull","logs":["exported 1 VBA component(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true},"source":{"modules_dir":"C:\\temp\\src\\modules","classes_dir":"C:\\temp\\src\\classes","forms_dir":"C:\\temp\\src\\forms","workbook_dir":"C:\\temp\\src\\workbook"}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.PullWithOptions(config.Default(), SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Source == nil {
		t.Fatal("Source is nil, want dotnet pull result")
	}
}

func TestRunnerAutoBridgeUsesDotNetForPush(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"push","logs":["imported 2 source file(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true},"backup":{"id":"push_20260101T120000","path":"C:\\temp\\.xlflow\\backups\\Book_20260101T120000.xlsm","reason":"before-push","mode":"always"},"source":{"changed_only":false,"changed":true,"state":"C:\\temp\\.xlflow\\state\\push.json"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"push","logs":["imported 2 source file(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","save_required":true},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true},"backup":{"id":"push_20260101","path":"C:\\temp\\.xlflow\\backups\\Book_20260101.xlsm","reason":"before-push","mode":"always"},"source":{"changed_only":false,"changed":true,"state":"C:\\temp\\.xlflow\\state\\push.json"}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.PushWithOptions(config.Default(), PushOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Backup == nil {
		t.Fatal("Backup is nil, want dotnet push result")
	}
}

func TestRunnerAutoBridgeUsesDotNetForRun(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"run","logs":[],"target":{"kind":"saved_workbook","path":"C:\\temp\\Book.xlsm"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":false,"saved":false,"dirty":false,"needs_save":false},"macro":{"name":"Main.Run","duration_ms":42,"arguments":[]},"runtime":{"mode":"interactive","source":"default","injected":true}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"run","logs":["unexpected powershell execution"]}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.Run(config.Default(), RunOptions{Macro: "Main.Run"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	macro, ok := env.Macro.(map[string]interface{})
	if !ok || macro["name"] != "Main.Run" {
		t.Fatalf("unexpected macro payload: %#v", env.Macro)
	}
}

func TestRunnerAutoBridgeDoesNotFallBackToPowerShellForBridgeCommandUnsupported(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"pull","logs":[],"error":{"code":"BRIDGE_COMMAND_UNSUPPORTED","message":"Command 'pull' is not supported by the .NET bridge.","source":"xlflow-excel-bridge","phase":"bridge.capability"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"pull","logs":["exported 1 VBA component(s)"],"target":{"kind":"saved_workbook","path":"C:\\temp\\Book.xlsm"},"session":{"active":false},"workbook":{"path":"C:\\temp\\Book.xlsm","session":false},"source":{"modules_dir":"C:\\temp\\src\\modules","classes_dir":"C:\\temp\\src\\classes","forms_dir":"C:\\temp\\src\\forms","workbook_dir":"C:\\temp\\src\\workbook"}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.PullWithOptions(config.Default(), SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "BRIDGE_COMMAND_UNSUPPORTED" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerAutoBridgeDoesNotFallBackToPowerShellForDotNetBridgeFileNotOpenable(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"pull","logs":[],"error":{"code":"bridge_file_not_openable","message":"failed to open workbook","source":"xlflow-excel-bridge","phase":"open_workbook"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
				response:  excelbridge.Response{Stdout: []byte(`{"status":"ok","command":"pull","logs":[],"source":{"modules_dir":"C:\\temp\\src\\modules"}}`)},
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "auto"}.PullWithOptions(config.Default(), SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if env.Error == nil || env.Error.Code != "bridge_file_not_openable" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDotNetPullUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"pull","logs":["exported 4 VBA component(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"explicit","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"dirty":false,"needs_save":false},"source":{"modules_dir":"C:\\temp\\src\\modules","classes_dir":"C:\\temp\\src\\classes","forms_dir":"C:\\temp\\src\\forms","workbook_dir":"C:\\temp\\src\\workbook","code_source":"sidecar"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PullWithOptions(cfg, SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["CodeSource"]; got != "sidecar" {
		t.Fatalf("CodeSource = %q, want sidecar", got)
	}
	if got := dotNetRequests[0].Args["UseSession"]; got != "true" {
		t.Fatalf("UseSession = %q, want true", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Source == nil {
		t.Fatalf("expected pull envelope fields to be populated: %+v", env)
	}
	source, ok := env.Source.(map[string]interface{})
	if !ok {
		t.Fatalf("expected source payload map, got %#v", env.Source)
	}
	if source["code_source"] != "sidecar" {
		t.Fatalf("source.code_source = %v, want sidecar", source["code_source"])
	}
}

func TestRunnerDotNetPullSessionRequiredErrorPreserved(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"pull","logs":[],"error":{"code":"session_required","message":"xlflow session is not running","phase":"pull","source":"xlflow"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PullWithOptions(config.Default(), SessionCommandOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "session_required" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDotNetPushSessionRequiredErrorPreserved(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"push","logs":[],"error":{"code":"session_required","message":"xlflow session is not running","phase":"push","source":"xlflow"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PushWithOptions(config.Default(), PushOptions{Session: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "session_required" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerDotNetPushUsesDotNetProviderAndPreservesEnvelopeFields(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })

	var dotNetCalls, powerShellCalls int
	var dotNetRequests []excelbridge.Request
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		switch mode {
		case excelbridge.ModeDotNet:
			return trackingBridgeProvider{
				name:      string(excelbridge.ModeDotNet),
				supports:  true,
				callCount: &dotNetCalls,
				requests:  &dotNetRequests,
				response:  excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"push","logs":["imported 4 source file(s)"],"target":{"kind":"live_session","path":"C:\\temp\\Book.xlsm"},"session":{"active":true,"workbook_path":"C:\\temp\\Book.xlsm","dirty":true,"save_required":true,"live_newer_than_disk":true,"mode":"explicit","source_of_truth":"live_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":true,"session_mode":"explicit","session_requested":true,"auto_session":false,"saved":false,"dirty":true,"needs_save":true},"backup":{"id":"push_20260101T120000","path":"C:\\temp\\.xlflow\\backups\\Book_20260101T120000.xlsm","reason":"before-push","mode":"never"},"source":{"changed_only":true,"changed":true,"state":"C:\\temp\\.xlflow\\state\\push.json"}}`)},
			}
		default:
			return trackingBridgeProvider{
				name:      "powershell",
				supports:  true,
				callCount: &powerShellCalls,
			}
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PushWithOptions(config.Default(), PushOptions{
		Session:     true,
		ChangedOnly: true,
		BackupMode:  "never",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if dotNetCalls != 1 {
		t.Fatalf("dotnet calls = %d, want 1", dotNetCalls)
	}
	if powerShellCalls != 0 {
		t.Fatalf("powershell calls = %d, want 0", powerShellCalls)
	}
	if len(dotNetRequests) != 1 {
		t.Fatalf("dotnet request count = %d, want 1", len(dotNetRequests))
	}
	if got := dotNetRequests[0].Args["ChangedOnly"]; got != "true" {
		t.Fatalf("ChangedOnly = %q, want true", got)
	}
	if got := dotNetRequests[0].Args["BackupMode"]; got != "never" {
		t.Fatalf("BackupMode = %q, want never", got)
	}
	if env.Target == nil || env.Session == nil || env.Workbook == nil || env.Backup == nil || env.Source == nil {
		t.Fatalf("expected push envelope fields to be populated: %+v", env)
	}
	backup, ok := env.Backup.(map[string]interface{})
	if !ok {
		t.Fatalf("expected backup payload map, got %#v", env.Backup)
	}
	if backup["mode"] != "never" {
		t.Fatalf("backup.mode = %v, want never", backup["mode"])
	}
	source, ok := env.Source.(map[string]interface{})
	if !ok {
		t.Fatalf("expected source payload map, got %#v", env.Source)
	}
	if source["changed_only"] != true || source["changed"] != true {
		t.Fatalf("unexpected source payload: %+v", source)
	}
}

func TestRunnerDotNetPushChangedOnlyNoOpKeepsWorkbookFileBacked(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"push","logs":["source state unchanged; skipped workbook import"],"target":{"kind":"file","path":"C:\\temp\\Book.xlsm"},"session":{"active":false,"workbook_path":"C:\\temp\\Book.xlsm","dirty":false,"save_required":false,"live_newer_than_disk":false,"mode":"none","source_of_truth":"saved_workbook"},"workbook":{"path":"C:\\temp\\Book.xlsm","session":false,"session_mode":"none","session_requested":false,"auto_session":false,"saved":false,"dirty":false,"needs_save":false},"backup":{"path":null,"mode":"never"},"source":{"changed_only":true,"changed":false,"state":"C:\\temp\\.xlflow\\state\\push.json"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PushWithOptions(config.Default(), PushOptions{ChangedOnly: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Target == nil {
		t.Fatal("Target is nil")
	}
	target, ok := env.Target.(map[string]interface{})
	if !ok || target["kind"] != "file" {
		t.Fatalf("unexpected Target kind: %+v", target)
	}
	if env.Session == nil {
		t.Fatal("Session is nil")
	}
	session, ok := env.Session.(map[string]interface{})
	if !ok || session["active"] != false {
		t.Fatalf("unexpected Session active: %+v", session)
	}
	if session["mode"] != "none" {
		t.Fatalf("unexpected Session mode: %v, want none", session["mode"])
	}
	if env.Source == nil {
		t.Fatal("Source is nil")
	}
	source, ok := env.Source.(map[string]interface{})
	if !ok || source["changed_only"] != true {
		t.Fatalf("unexpected Source changed_only: %+v", source)
	}
	if source["changed"] != false {
		t.Fatalf("unexpected Source changed: %+v", source)
	}
}

func TestRunnerDotNetPushChangedOnlySessionRequiredFailsFast(t *testing.T) {
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return fakeBridgeProvider{
			name:     string(excelbridge.ModeDotNet),
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"failed","command":"push","logs":[],"error":{"code":"session_required","message":"xlflow session is not running","phase":"push","source":"xlflow"}}`)},
		}
	}

	env, code, err := Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}.PushWithOptions(config.Default(), PushOptions{Session: true, ChangedOnly: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "session_required" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}
