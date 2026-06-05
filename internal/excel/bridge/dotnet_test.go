package bridge

import (
	"context"
	"encoding/json"
	"errors"
	"os/exec"
	"runtime"
	"testing"
	"time"
)

func TestDotNetBridgeCommandFallsBackToInstalledBridgeWhenSDKMissing(t *testing.T) {
	originalLookPath := dotNetLookPath
	originalCandidatesFunc := dotNetBridgeCandidatesFunc
	originalProjectPathFunc := repoLocalDotNetBridgeProjectPathFunc
	t.Cleanup(func() {
		dotNetLookPath = originalLookPath
		dotNetBridgeCandidatesFunc = originalCandidatesFunc
		repoLocalDotNetBridgeProjectPathFunc = originalProjectPathFunc
	})

	dotNetBridgeCandidatesFunc = func() []string { return nil }
	repoLocalDotNetBridgeProjectPathFunc = func() (string, bool) {
		return `C:\dev\go\xlflow\bridge\dotnet\src\Xlflow.ExcelBridge\Xlflow.ExcelBridge.csproj`, true
	}
	dotNetLookPath = func(file string) (string, error) {
		switch file {
		case "dotnet":
			return "", errors.New("not found")
		case dotNetBridgeBinaryName + ".exe":
			return `C:\tools\xlflow-excel-bridge.exe`, nil
		default:
			return "", errors.New("not found")
		}
	}

	command, args, err := DotNetBridgeCommand()
	if err != nil {
		t.Fatalf("DotNetBridgeCommand() error = %v", err)
	}
	if command != `C:\tools\xlflow-excel-bridge.exe` {
		t.Fatalf("command = %q, want installed bridge executable", command)
	}
	if args != nil {
		t.Fatalf("args = %v, want nil for installed bridge executable", args)
	}
}

func TestRepoLocalDotNetBridgeProjectPathExists(t *testing.T) {
	projectPath, ok := repoLocalDotNetBridgeProjectPath()
	if !ok {
		t.Fatal("expected repo-local .NET bridge project path")
	}
	if projectPath == "" {
		t.Fatal("expected non-empty project path")
	}
}

func TestDotNetProviderAdvertisesMajorCommandsForAutoSelection(t *testing.T) {
	provider := DotNetProvider{}

	for _, command := range []string{"doctor", "pull", "push", "run", "macros", "process", "test", "trace"} {
		if !provider.Supports(command) {
			t.Fatalf("Supports(%q) = false, want true; auto mode should prefer .NET for major Windows bridge commands", command)
		}
	}
}

func TestDotNetSupportedCommandsMatchExpectedSet(t *testing.T) {
	expected := []string{
		"attach",
		"doctor",
		"edit",
		"export-image",
		"form-export-image",
		"form-write",
		"inspect",
		"inspect-form",
		"list",
		"macros",
		"new",
		"process",
		"pull",
		"push",
		"run",
		"runner",
		"session",
		"test",
		"trace",
		"ui",
	}

	if len(dotNetSupportedCommands) != len(expected) {
		t.Fatalf("dotNetSupportedCommands has %d entries, want %d. Got: %v", len(dotNetSupportedCommands), len(expected), dotNetSupportedCommands)
	}

	for _, cmd := range expected {
		if _, ok := dotNetSupportedCommands[cmd]; !ok {
			t.Errorf("dotNetSupportedCommands missing %q; must match auto-mode selection set: %v", cmd, expected)
		}
	}
}

func TestDotNetProviderExecuteDoctorWithRealBridge(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows-only repo-local .NET bridge execution")
	}
	if _, err := exec.LookPath("dotnet"); err != nil {
		t.Skip("dotnet is not available")
	}
	if _, ok := repoLocalDotNetBridgeProjectPath(); !ok {
		t.Skip("repo-local .NET bridge project not found")
	}

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	defer cancel()

	response, err := (DotNetProvider{}).Execute(ctx, Request{Command: "doctor", Args: map[string]string{}})
	if err != nil {
		var bridgeErr *Error
		if errors.As(err, &bridgeErr) && bridgeErr.Kind == ErrorDotNetRuntime {
			t.Skip("dotnet runtime is not available for repo-local bridge execution")
		}
		t.Fatalf("Execute(doctor) error = %v", err)
	}

	var payload struct {
		ProtocolVersion int    `json:"protocol_version"`
		Status          string `json:"status"`
		Command         string `json:"command"`
		Error           *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err := json.Unmarshal(response.Stdout, &payload); err != nil {
		t.Fatalf("failed to parse bridge response: %v\nstdout=%s", err, string(response.Stdout))
	}
	if payload.ProtocolVersion != ProtocolVersion {
		t.Fatalf("protocol_version = %d, want %d", payload.ProtocolVersion, ProtocolVersion)
	}
	if payload.Status == "failed" && payload.Error != nil && payload.Error.Code == "excel_com_failure" {
		t.Skipf("Excel COM is not available for repo-local .NET bridge execution: %s", payload.Error.Message)
	}
	if payload.Status != "ok" {
		t.Fatalf("status = %q, want ok", payload.Status)
	}
	if payload.Command != "doctor" {
		t.Fatalf("command = %q, want doctor", payload.Command)
	}
}
