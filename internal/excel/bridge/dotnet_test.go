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

func TestRepoLocalDotNetBridgeProjectPathExists(t *testing.T) {
	projectPath, ok := repoLocalDotNetBridgeProjectPath()
	if !ok {
		t.Fatal("expected repo-local .NET bridge project path")
	}
	if projectPath == "" {
		t.Fatal("expected non-empty project path")
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
	}
	if err := json.Unmarshal(response.Stdout, &payload); err != nil {
		t.Fatalf("failed to parse bridge response: %v\nstdout=%s", err, string(response.Stdout))
	}
	if payload.ProtocolVersion != ProtocolVersion {
		t.Fatalf("protocol_version = %d, want %d", payload.ProtocolVersion, ProtocolVersion)
	}
	if payload.Status != "ok" {
		t.Fatalf("status = %q, want ok", payload.Status)
	}
	if payload.Command != "doctor" {
		t.Fatalf("command = %q, want doctor", payload.Command)
	}
}
