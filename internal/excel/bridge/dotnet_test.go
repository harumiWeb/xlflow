package bridge

import (
	"context"
	"encoding/json"
	"errors"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/coordination"
)

func TestDotNetSupportedCommandsHaveCoordinationSelectors(t *testing.T) {
	covered := make(map[string]bool, len(dotNetSupportedCommands))
	for _, descriptor := range coordination.All() {
		for _, selector := range descriptor.Bridge {
			covered[strings.ToLower(selector.Command)] = true
		}
	}
	for command := range dotNetSupportedCommands {
		if !covered[command] {
			t.Errorf(".NET bridge command %q has no coordination selector", command)
		}
	}
}

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

func TestDotNetBridgeInternalRunFlagConstant(t *testing.T) {
	if dotNetBridgeInternalRunFlag != "--bridge-internal-run" {
		t.Fatalf("dotNetBridgeInternalRunFlag = %q, want %q", dotNetBridgeInternalRunFlag, "--bridge-internal-run")
	}
}

func TestDotNetBridgeRuntimeArgsAppendsInternalRunFlag(t *testing.T) {
	base := []string{"bridge.dll", "--extra"}
	got := dotNetBridgeRuntimeArgs(base)

	if len(got) != 3 {
		t.Fatalf("len(dotNetBridgeRuntimeArgs()) = %d, want 3: %v", len(got), got)
	}
	if got[0] != "bridge.dll" || got[1] != "--extra" {
		t.Fatalf("dotNetBridgeRuntimeArgs() changed base args order: %v", got)
	}
	if got[2] != dotNetBridgeInternalRunFlag {
		t.Fatalf("dotNetBridgeRuntimeArgs() missing internal run flag: %v", got)
	}
	if len(base) != 2 {
		t.Fatalf("dotNetBridgeRuntimeArgs() mutated input slice length: %v", base)
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

func TestDotNetBridgeCommandUsesBuildServerSafeArgsForProjectFallback(t *testing.T) {
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
		if file == "dotnet" {
			return `C:\Program Files\dotnet\dotnet.exe`, nil
		}
		return "", errors.New("not found")
	}

	command, args, err := DotNetBridgeCommand()
	if err != nil {
		t.Fatalf("DotNetBridgeCommand() error = %v", err)
	}
	if command != `C:\Program Files\dotnet\dotnet.exe` {
		t.Fatalf("command = %q, want dotnet executable", command)
	}
	joined := strings.Join(args, " ")
	for _, want := range []string{"run", "--disable-build-servers", "-p:UseSharedCompilation=false", "-p:BuildInParallel=false"} {
		if !strings.Contains(joined, want) {
			t.Fatalf("project fallback args missing %q: %v", want, args)
		}
	}
}

func TestDotNetBridgeCandidatesIncludeResolvedExecutableSiblings(t *testing.T) {
	originalExecutablePath := dotNetExecutablePath
	originalEvalSymlinks := dotNetEvalSymlinks
	t.Cleanup(func() {
		dotNetExecutablePath = originalExecutablePath
		dotNetEvalSymlinks = originalEvalSymlinks
	})

	linkExe := `C:\Users\HARUMI\AppData\Local\Microsoft\WinGet\Links\xlflow.exe`
	packageExe := `C:\Users\HARUMI\AppData\Local\Microsoft\WinGet\Packages\HarumiWeb.Xlflow_Microsoft.Winget.Source_8wekyb3d8bbwe\xlflow.exe`
	dotNetExecutablePath = func() (string, error) { return linkExe, nil }
	dotNetEvalSymlinks = func(path string) (string, error) {
		if path != linkExe {
			t.Fatalf("EvalSymlinks path = %q, want %q", path, linkExe)
		}
		return packageExe, nil
	}

	candidates := dotNetBridgeCandidates()
	wantLinkCandidate := filepath.Join(filepath.Dir(linkExe), dotNetBridgeBinaryName+".exe")
	wantPackageCandidate := filepath.Join(filepath.Dir(packageExe), dotNetBridgeBinaryName+".exe")
	if !containsPath(candidates, wantLinkCandidate) {
		t.Fatalf("candidates missing link sibling %q:\n%v", wantLinkCandidate, candidates)
	}
	if !containsPath(candidates, wantPackageCandidate) {
		t.Fatalf("candidates missing resolved package sibling %q:\n%v", wantPackageCandidate, candidates)
	}
}

func TestDotNetProviderAdvertisesMajorCommandsForAutoSelection(t *testing.T) {
	provider := DotNetProvider{}

	for _, command := range []string{"doctor", "pull", "push", "run", "macros", "process", "test"} {
		if !provider.Supports(command) {
			t.Fatalf("Supports(%q) = false, want true; auto mode should prefer .NET for major Windows bridge commands", command)
		}
	}
}

func containsPath(paths []string, want string) bool {
	for _, path := range paths {
		if samePath(path, want) {
			return true
		}
	}
	return false
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
		"type-db-import",
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

	projectPath, _ := repoLocalDotNetBridgeProjectPath()
	outDir := filepath.Join(t.TempDir(), "publish")
	objDir := filepath.Join(t.TempDir(), "obj")
	buildCmd := exec.Command(
		"dotnet",
		"publish",
		projectPath,
		"-c", "Release",
		"-o", outDir,
		"--disable-build-servers",
		"-p:UseSharedCompilation=false",
		"-p:BuildInParallel=false",
		"-p:BaseIntermediateOutputPath="+objDir+string(os.PathSeparator),
	)
	buildOut, buildErr := buildCmd.CombinedOutput()
	if buildErr != nil {
		t.Skipf("could not build isolated repo-local .NET bridge for execution test: %v\n%s", buildErr, string(buildOut))
	}

	bridgeExe := filepath.Join(outDir, dotNetBridgeBinaryName+".exe")
	if _, err := os.Stat(bridgeExe); err != nil {
		t.Skipf("isolated repo-local .NET bridge executable was not produced at %s: %v", bridgeExe, err)
	}

	originalCandidatesFunc := dotNetBridgeCandidatesFunc
	originalProjectPathFunc := repoLocalDotNetBridgeProjectPathFunc
	t.Cleanup(func() {
		dotNetBridgeCandidatesFunc = originalCandidatesFunc
		repoLocalDotNetBridgeProjectPathFunc = originalProjectPathFunc
	})
	dotNetBridgeCandidatesFunc = func() []string { return []string{bridgeExe} }
	repoLocalDotNetBridgeProjectPathFunc = func() (string, bool) { return "", false }

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

func TestDotNetProviderExecuteNewWithNonASCIIPathKeepsJSONValid(t *testing.T) {
	bridgeExe := publishRepoLocalDotNetBridge(t)
	useIsolatedBridgeExecutable(t, bridgeExe)

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	defer cancel()

	workbookPath := filepath.Join(t.TempDir(), "てすと", "sources", "帳票.xlsm")
	response, err := (DotNetProvider{}).Execute(ctx, Request{
		Command: "new",
		Args: map[string]string{
			"WorkbookPath": workbookPath,
		},
	})
	if err != nil && len(response.Stdout) == 0 {
		var bridgeErr *Error
		if errors.As(err, &bridgeErr) && bridgeErr.Kind == ErrorDotNetRuntime {
			t.Skip("dotnet runtime is not available for repo-local bridge execution")
		}
		t.Fatalf("Execute(new) error = %v", err)
	}

	payload := parseBridgeResponseEnvelope(t, response.Stdout)
	if payload.ProtocolVersion != ProtocolVersion {
		t.Fatalf("protocol_version = %d, want %d", payload.ProtocolVersion, ProtocolVersion)
	}
	if payload.Command != "new" {
		t.Fatalf("command = %q, want new", payload.Command)
	}
	if payload.Error != nil && payload.Error.Code == "BRIDGE_REQUEST_INVALID_JSON" {
		t.Fatalf("transport corrupted non-ASCII path into invalid JSON: %s", payload.Error.Message)
	}
	if payload.Error != nil && payload.Error.Code == "excel_com_failure" {
		t.Skipf("Excel COM is not available for repo-local .NET bridge execution: %s", payload.Error.Message)
	}
}

func TestDotNetProviderExecuteSessionStatusWithNonASCIIPathKeepsJSONValid(t *testing.T) {
	bridgeExe := publishRepoLocalDotNetBridge(t)
	useIsolatedBridgeExecutable(t, bridgeExe)

	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()

	projectRoot := filepath.Join(t.TempDir(), "てすと", "session-source")
	metadataPath := filepath.Join(projectRoot, ".xlflow", "session.json")
	workbookPath := filepath.Join(projectRoot, "帳票.xlsm")
	response, err := (DotNetProvider{}).Execute(ctx, Request{
		Command: "session",
		Args: map[string]string{
			"Action":       "status",
			"WorkbookPath": workbookPath,
			"MetadataPath": metadataPath,
			"UseSession":   "false",
			"Visible":      "false",
		},
	})
	if err != nil && len(response.Stdout) == 0 {
		var bridgeErr *Error
		if errors.As(err, &bridgeErr) && bridgeErr.Kind == ErrorDotNetRuntime {
			t.Skip("dotnet runtime is not available for repo-local bridge execution")
		}
		t.Fatalf("Execute(session status) error = %v", err)
	}

	payload := parseBridgeResponseEnvelope(t, response.Stdout)
	if payload.ProtocolVersion != ProtocolVersion {
		t.Fatalf("protocol_version = %d, want %d", payload.ProtocolVersion, ProtocolVersion)
	}
	if payload.Command != "session" {
		t.Fatalf("command = %q, want session", payload.Command)
	}
	if payload.Error != nil && payload.Error.Code == "BRIDGE_REQUEST_INVALID_JSON" {
		t.Fatalf("transport corrupted non-ASCII session path into invalid JSON: %s", payload.Error.Message)
	}
}

type bridgeResponseEnvelope struct {
	ProtocolVersion int    `json:"protocol_version"`
	Status          string `json:"status"`
	Command         string `json:"command"`
	Error           *struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	} `json:"error"`
}

func parseBridgeResponseEnvelope(t *testing.T, stdout []byte) bridgeResponseEnvelope {
	t.Helper()

	var payload bridgeResponseEnvelope
	if err := json.Unmarshal(stdout, &payload); err != nil {
		t.Fatalf("failed to parse bridge response: %v\nstdout=%s", err, string(stdout))
	}
	return payload
}

func publishRepoLocalDotNetBridge(t *testing.T) string {
	t.Helper()

	if runtime.GOOS != "windows" {
		t.Skip("Windows-only repo-local .NET bridge execution")
	}
	if _, err := exec.LookPath("dotnet"); err != nil {
		t.Skip("dotnet is not available")
	}
	if _, ok := repoLocalDotNetBridgeProjectPath(); !ok {
		t.Skip("repo-local .NET bridge project not found")
	}

	projectPath, _ := repoLocalDotNetBridgeProjectPath()
	outDir := filepath.Join(t.TempDir(), "publish")
	objDir := filepath.Join(t.TempDir(), "obj")
	buildCmd := exec.Command(
		"dotnet",
		"publish",
		projectPath,
		"-c", "Release",
		"-o", outDir,
		"--disable-build-servers",
		"-p:UseSharedCompilation=false",
		"-p:BuildInParallel=false",
		"-p:BaseIntermediateOutputPath="+objDir+string(os.PathSeparator),
	)
	buildOut, buildErr := buildCmd.CombinedOutput()
	if buildErr != nil {
		t.Skipf("could not build isolated repo-local .NET bridge for execution test: %v\n%s", buildErr, string(buildOut))
	}

	bridgeExe := filepath.Join(outDir, dotNetBridgeBinaryName+".exe")
	if _, err := os.Stat(bridgeExe); err != nil {
		t.Skipf("isolated repo-local .NET bridge executable was not produced at %s: %v", bridgeExe, err)
	}

	return bridgeExe
}

func useIsolatedBridgeExecutable(t *testing.T, bridgeExe string) {
	t.Helper()

	originalCandidatesFunc := dotNetBridgeCandidatesFunc
	originalProjectPathFunc := repoLocalDotNetBridgeProjectPathFunc
	t.Cleanup(func() {
		dotNetBridgeCandidatesFunc = originalCandidatesFunc
		repoLocalDotNetBridgeProjectPathFunc = originalProjectPathFunc
	})
	dotNetBridgeCandidatesFunc = func() []string { return []string{bridgeExe} }
	repoLocalDotNetBridgeProjectPathFunc = func() (string, bool) { return "", false }
}
