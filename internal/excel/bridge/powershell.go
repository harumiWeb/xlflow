package bridge

import (
	"bytes"
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"

	bundledscripts "github.com/harumiWeb/xlflow/internal/excel/scripts"
)

type PowerShellProvider struct {
	RootDir string
}

func (p PowerShellProvider) Name() string {
	return string(ModePowerShell)
}

func (p PowerShellProvider) Supports(string) bool {
	return true
}

func (p PowerShellProvider) Info(context.Context) (Info, error) {
	return Info{Name: p.Name()}, nil
}

func (p PowerShellProvider) Execute(ctx context.Context, req Request) (Response, error) {
	if !ScriptExecutionSupported(p.RootDir, req.Command) {
		return Response{}, fmt.Errorf("excel automation is only supported on Windows in the MVP unless a script override is provided at scripts/%s.ps1", req.Command)
	}

	script, cleanup, err := ScriptPath(p.RootDir, req.Command)
	if err != nil {
		return Response{}, err
	}
	if cleanup != nil {
		defer cleanup()
	}

	powershellExe, err := PowerShellExecutable()
	if err != nil {
		return Response{}, err
	}

	cmdArgs := []string{"-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script}
	for k, v := range req.Args {
		cmdArgs = append(cmdArgs, "-"+k, v)
	}

	cmd := exec.CommandContext(ctx, powershellExe, cmdArgs...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	err = cmd.Start()
	if err == nil {
		err = cmd.Wait()
	}
	return Response{
		Stdout:   stdout.Bytes(),
		Stderr:   stderr.Bytes(),
		TimedOut: ctx.Err() == context.DeadlineExceeded,
	}, err
}

func ScriptPath(root, commandName string) (string, func(), error) {
	if path, ok := ExternalScriptPath(root, commandName); ok {
		return path, nil, nil
	}
	return MaterializeBundledScript(commandName)
}

func MaterializeBundledScript(commandName string) (string, func(), error) {
	name := commandName + ".ps1"
	path, cleanup, err := bundledscripts.Materialize(commandName)
	if err != nil {
		return "", nil, fmt.Errorf("script %s was not available from on-disk script locations or embedded runtime assets: %w", name, err)
	}
	return path, cleanup, nil
}

func ExternalScriptPath(root, commandName string) (string, bool) {
	name := commandName + ".ps1"
	candidates := []string{}
	if path, ok := RootScriptOverridePath(root, commandName); ok {
		candidates = append(candidates, path)
	}
	if _, file, _, ok := runtime.Caller(0); ok {
		candidates = append(candidates, filepath.Join(filepath.Dir(filepath.Dir(file)), "scripts", name))
	}
	if exe, err := os.Executable(); err == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(exe), "scripts", name))
	}
	for _, candidate := range candidates {
		clean := filepath.Clean(candidate)
		if _, err := os.Stat(clean); err == nil {
			return clean, true
		}
	}
	return "", false
}

func HasExternalScriptOverride(root, commandName string) bool {
	_, ok := RootScriptOverridePath(root, commandName)
	return ok
}

func ScriptExecutionSupported(root, commandName string) bool {
	return runtime.GOOS == "windows" || HasExternalScriptOverride(root, commandName)
}

func PowerShellExecutable() (string, error) {
	return PowerShellExecutableFor(runtime.GOOS, exec.LookPath)
}

func PowerShellExecutableFor(goos string, lookPath func(file string) (string, error)) (string, error) {
	var candidates []string
	if goos == "windows" {
		candidates = []string{"powershell"}
	} else {
		candidates = []string{"pwsh", "powershell"}
	}
	for _, candidate := range candidates {
		if _, err := lookPath(candidate); err == nil {
			return candidate, nil
		}
	}
	return "", fmt.Errorf("PowerShell executable not found on %s (searched: %s)", goos, strings.Join(candidates, ", "))
}

func RootScriptOverridePath(root, commandName string) (string, bool) {
	if root == "" {
		return "", false
	}
	candidate := filepath.Clean(filepath.Join(root, "scripts", commandName+".ps1"))
	if _, err := os.Stat(candidate); err == nil {
		return candidate, true
	}
	return "", false
}
