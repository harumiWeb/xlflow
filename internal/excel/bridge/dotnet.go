package bridge

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

const dotNetBridgeBinaryName = "xlflow-excel-bridge"
const dotNetBridgeProjectRelativePath = "bridge/dotnet/src/Xlflow.ExcelBridge/Xlflow.ExcelBridge.csproj"
const dotNetBridgeInternalRunFlag = "--bridge-internal-run"

var dotNetLookPath = exec.LookPath
var dotNetBridgeCandidatesFunc = dotNetBridgeCandidates
var repoLocalDotNetBridgeProjectPathFunc = repoLocalDotNetBridgeProjectPath

// dotNetSupportedCommands is the Go-side source of truth for Windows auto-mode
// provider support.
var dotNetSupportedCommands = map[string]struct{}{
	"attach":            {},
	"doctor":            {},
	"edit":              {},
	"export-image":      {},
	"form-export-image": {},
	"form-write":        {},
	"inspect":           {},
	"inspect-form":      {},
	"list":              {},
	"macros":            {},
	"new":               {},
	"process":           {},
	"pull":              {},
	"push":              {},
	"run":               {},
	"runner":            {},
	"session":           {},
	"test":              {},
	"type-db-import":    {},
	"ui":                {},
}

type DotNetProvider struct {
	RootDir string
}

type dotNetBridgeInfo struct {
	Name            string `json:"name"`
	Version         string `json:"version"`
	ProtocolVersion int    `json:"protocol_version"`
}

type dotNetBridgeRequest struct {
	ProtocolVersion int               `json:"protocol_version"`
	RequestID       string            `json:"request_id"`
	Command         string            `json:"command"`
	TimeoutMS       int64             `json:"timeout_ms,omitempty"`
	Payload         map[string]string `json:"payload"`
}

func (p DotNetProvider) Name() string {
	return string(ModeDotNet)
}

func (p DotNetProvider) Supports(command string) bool {
	_, ok := dotNetSupportedCommands[strings.ToLower(strings.TrimSpace(command))]
	return ok
}

func (p DotNetProvider) Info(ctx context.Context) (Info, error) {
	command, args, err := DotNetBridgeCommand()
	if err != nil {
		return Info{}, err
	}

	cmd := exec.CommandContext(ctx, command, append(args, "--version-json")...)
	var stdout bytes.Buffer
	var stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		if stderr.Len() > 0 {
			return Info{}, fmt.Errorf("%s", strings.TrimSpace(stderr.String()))
		}
		return Info{}, err
	}

	var info dotNetBridgeInfo
	if err := json.Unmarshal(stdout.Bytes(), &info); err != nil {
		return Info{}, err
	}
	return Info{Name: info.Name, Version: info.Version}, nil
}

func (p DotNetProvider) Execute(ctx context.Context, req Request) (Response, error) {
	command, args, err := DotNetBridgeCommand()
	if err != nil {
		return Response{}, err
	}

	request := dotNetBridgeRequest{
		ProtocolVersion: ProtocolVersion,
		RequestID:       newBridgeRequestID(req.Command),
		Command:         req.Command,
		Payload:         req.Args,
	}
	if deadline, ok := ctx.Deadline(); ok {
		request.TimeoutMS = max(1, time.Until(deadline).Milliseconds())
	}

	body, err := json.Marshal(request)
	if err != nil {
		return Response{}, err
	}

	execArgs := dotNetBridgeRuntimeArgs(args)
	cmd := exec.CommandContext(ctx, command, execArgs...)
	cmd.Stdin = bytes.NewReader(body)
	var stdout bytes.Buffer
	var stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	err = cmd.Run()

	response := Response{
		Stdout:   stdout.Bytes(),
		Stderr:   stderr.Bytes(),
		TimedOut: ctx.Err() == context.DeadlineExceeded,
	}
	if err != nil {
		if response.TimedOut {
			return response, err
		}
		if len(response.Stdout) > 0 {
			return response, nil
		}
		if len(response.Stderr) > 0 {
			return response, fmt.Errorf("%s", strings.TrimSpace(stderr.String()))
		}
		return response, err
	}
	return response, nil
}

func dotNetBridgeRuntimeArgs(args []string) []string {
	runtimeArgs := append([]string{}, args...)
	runtimeArgs = append(runtimeArgs, dotNetBridgeInternalRunFlag)
	return runtimeArgs
}

func DotNetBridgeCommand() (string, []string, error) {
	var deferredErr error

	for _, candidate := range dotNetBridgeCandidatesFunc() {
		if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
			switch strings.ToLower(filepath.Ext(candidate)) {
			case ".dll":
				dotnetExe, err := dotNetLookPath("dotnet")
				if err != nil {
					if deferredErr == nil {
						deferredErr = &Error{
							Kind:    ErrorDotNetRuntime,
							Message: "dotnet executable not found while resolving the .NET Excel bridge runtime",
							Err:     err,
						}
					}
					continue
				}
				return dotnetExe, []string{candidate}, nil
			default:
				return candidate, nil, nil
			}
		}
	}

	if projectPath, ok := repoLocalDotNetBridgeProjectPathFunc(); ok {
		dotnetExe, err := dotNetLookPath("dotnet")
		if err == nil {
			return dotnetExe, []string{
				"run",
				"--project", projectPath,
				"--configuration", "Release",
				"--disable-build-servers",
				"-p:UseSharedCompilation=false",
				"-p:BuildInParallel=false",
				"--",
			}, nil
		}
		if deferredErr == nil {
			deferredErr = &Error{
				Kind:    ErrorDotNetRuntime,
				Message: "dotnet executable not found while resolving the repo-local .NET Excel bridge project",
				Err:     err,
			}
		}
	}

	for _, candidate := range []string{dotNetBridgeBinaryName + ".exe", dotNetBridgeBinaryName} {
		if path, err := dotNetLookPath(candidate); err == nil {
			return path, nil, nil
		}
	}

	if deferredErr != nil {
		return "", nil, deferredErr
	}

	return "", nil, &Error{
		Kind:    ErrorDotNetMissing,
		Message: ".NET Excel bridge executable was not found; build bridge/dotnet/Xlflow.ExcelBridge.sln or install xlflow-excel-bridge beside xlflow",
	}
}

func dotNetBridgeCandidates() []string {
	candidates := []string{}
	if _, file, _, ok := runtime.Caller(0); ok {
		repoRoot := repoRootFromBridgeFile(file)
		candidates = append(candidates,
			filepath.Join(repoRoot, "bridge", "dotnet", "artifacts", "publish", "win-x64", dotNetBridgeBinaryName+".exe"),
			filepath.Join(repoRoot, "bridge", "dotnet", "artifacts", "publish", "win-x64", dotNetBridgeBinaryName+".dll"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Release", "net8.0", "win-x64", dotNetBridgeBinaryName+".exe"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Release", "net8.0", "win-x64", dotNetBridgeBinaryName+".dll"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Release", "net8.0", dotNetBridgeBinaryName+".exe"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Release", "net8.0", dotNetBridgeBinaryName+".dll"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Debug", "net8.0", dotNetBridgeBinaryName+".exe"),
			filepath.Join(repoRoot, "bridge", "dotnet", "src", "Xlflow.ExcelBridge", "bin", "Debug", "net8.0", dotNetBridgeBinaryName+".dll"),
		)
	}
	if exe, err := os.Executable(); err == nil {
		dir := filepath.Dir(exe)
		candidates = append(candidates,
			filepath.Join(dir, dotNetBridgeBinaryName+".exe"),
			filepath.Join(dir, dotNetBridgeBinaryName),
			filepath.Join(dir, dotNetBridgeBinaryName+".dll"),
		)
	}
	return candidates
}

func repoLocalDotNetBridgeProjectPath() (string, bool) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		return "", false
	}
	projectPath := filepath.Join(repoRootFromBridgeFile(file), filepath.FromSlash(dotNetBridgeProjectRelativePath))
	if info, err := os.Stat(projectPath); err == nil && !info.IsDir() {
		return projectPath, true
	}
	return "", false
}

func repoRootFromBridgeFile(file string) string {
	return filepath.Dir(filepath.Dir(filepath.Dir(filepath.Dir(file))))
}

func newBridgeRequestID(command string) string {
	trimmed := strings.ToLower(strings.TrimSpace(command))
	if trimmed == "" {
		trimmed = "unknown"
	}
	return fmt.Sprintf("%s-%d", trimmed, time.Now().UTC().UnixNano())
}
