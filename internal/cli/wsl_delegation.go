package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/wsl"
)

var (
	isWSL                    = wsl.IsWSL
	wslDistroName            = wsl.DistroName
	resolveWindowsExecutable = wsl.ResolveWindowsExecutable
	translateWSLPath         = wsl.ToWindowsPath
	translateWSLArgs         = wsl.TranslateArgs
	newDelegatedCommand      = exec.CommandContext
)

var delegatedTopLevelCommands = map[string]struct{}{
	"attach":       {},
	"check":        {},
	"doctor":       {},
	"edit":         {},
	"export-image": {},
	"form":         {},
	"init":         {},
	"inspect":      {},
	"list":         {},
	"macros":       {},
	"new":          {},
	"process":      {},
	"pull":         {},
	"push":         {},
	"rollback":     {},
	"run":          {},
	"runner":       {},
	"save":         {},
	"session":      {},
	"status":       {},
	"test":         {},
	"ui":           {},
}

var errWSLDelegated = errors.New("wsl delegated command completed")

func (a *app) delegateWSLCommand(cmd *cobra.Command) error {
	if len(a.rawArgs) == 0 || !isWSL() || strings.TrimSpace(os.Getenv(wsl.EnvDelegated)) != "" {
		return nil
	}
	topLevel := topLevelCommandName(cmd)
	if !shouldDelegateTopLevelCommand(topLevel) {
		return nil
	}

	ctx := cmd.Context()
	if ctx == nil {
		ctx = context.Background()
	}
	windowsCWD, err := translateWSLPath(ctx, a.cwd)
	if err != nil {
		if topLevel == "doctor" {
			return a.writeWSLDoctorFailure(err, "", "", false)
		}
		return a.writeWSLFailure(topLevel, err)
	}
	executable, windowsExecutable, err := resolveWindowsExecutable(ctx)
	if err != nil {
		if topLevel == "doctor" {
			return a.writeWSLDoctorFailure(err, windowsCWD, "", false)
		}
		return a.writeWSLFailure(topLevel, err)
	}
	args, err := translateWSLArgs(ctx, a.rawArgs)
	if err != nil {
		if topLevel == "doctor" {
			return a.writeWSLDoctorFailure(err, windowsCWD, windowsExecutable, true)
		}
		return a.writeWSLFailure(topLevel, err)
	}

	if topLevel == "doctor" {
		return a.runDelegatedDoctor(ctx, executable, windowsExecutable, windowsCWD, args)
	}
	return a.runDelegatedCommand(ctx, topLevel, executable, args)
}

func shouldDelegateTopLevelCommand(name string) bool {
	_, ok := delegatedTopLevelCommands[name]
	return ok
}

func topLevelCommandName(cmd *cobra.Command) string {
	if cmd == nil {
		return ""
	}
	current := cmd
	for current.Parent() != nil && current.Parent().Parent() != nil {
		current = current.Parent()
	}
	return current.Name()
}

func (a *app) runDelegatedCommand(ctx context.Context, commandName string, executable string, args []string) error {
	child := newDelegatedCommand(ctx, executable, args...)
	child.Dir = a.cwd
	child.Stdin = os.Stdin
	child.Stdout = a.stdoutWriter()
	child.Stderr = a.stderrWriter()
	child.Env = delegatedEnvironment()
	err := child.Run()
	if err == nil {
		return output.WithExitCode(output.ExitSuccess, errWSLDelegated)
	}
	var exitErr *exec.ExitError
	if errors.As(err, &exitErr) {
		return output.WithExitCode(exitErr.ExitCode(), err)
	}
	return a.writeFailure(commandName, output.ExitEnvironment, "windows_xlflow_execution_failed", err)
}

func (a *app) runDelegatedDoctor(ctx context.Context, executable string, windowsExecutable string, windowsCWD string, args []string) error {
	doctorArgs := ensureJSONFlag(args)
	var stdout bytes.Buffer
	child := newDelegatedCommand(ctx, executable, doctorArgs...)
	child.Dir = a.cwd
	child.Stdin = os.Stdin
	child.Stdout = &stdout
	child.Stderr = a.stderrWriter()
	child.Env = delegatedEnvironment()
	runErr := child.Run()
	exitCode := output.ExitSuccess
	if runErr != nil {
		var exitErr *exec.ExitError
		if errors.As(runErr, &exitErr) {
			exitCode = exitErr.ExitCode()
		} else {
			return a.writeWSLDoctorFailure(
				&wsl.Error{Code: "windows_xlflow_execution_failed", Message: runErr.Error(), Err: runErr},
				windowsCWD,
				windowsExecutable,
				true,
			)
		}
	}

	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		return a.writeWSLDoctorFailure(
			&wsl.Error{
				Code:    "windows_xlflow_execution_failed",
				Message: fmt.Sprintf("Windows xlflow doctor returned invalid JSON: %v", err),
				Err:     err,
			},
			windowsCWD,
			windowsExecutable,
			true,
		)
	}
	diagnostics := mapValue(env.Diagnostics)
	diagnostics["host"] = map[string]any{
		"os":      "linux",
		"is_wsl":  true,
		"distro":  wslDistroName(),
		"version": a.buildInfo.Version,
	}
	diagnostics["windows"] = map[string]any{
		"xlflow_found":    true,
		"xlflow_path":     windowsExecutable,
		"xlflow_version":  a.windowsVersion(ctx, executable),
		"bridge_found":    delegatedDoctorBridgeFound(env, diagnostics),
		"excel_available": delegatedDoctorExcelAvailable(env, diagnostics),
	}
	diagnostics["path_translation"] = map[string]any{
		"supported":    true,
		"wsl_path":     a.cwd,
		"windows_path": windowsCWD,
	}
	env.Diagnostics = diagnostics
	a.addDelegatedVersionWarning(&env, diagnostics)
	if err := a.write(env, exitCode); err != nil {
		return err
	}
	return output.WithExitCode(output.ExitSuccess, errWSLDelegated)
}

func (a *app) windowsVersion(ctx context.Context, executable string) string {
	var stdout bytes.Buffer
	child := newDelegatedCommand(ctx, executable, "--json", "version")
	child.Dir = a.cwd
	child.Stdout = &stdout
	child.Stderr = a.stderrWriter()
	child.Env = delegatedEnvironment()
	if err := child.Run(); err != nil {
		return ""
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		return ""
	}
	return stringValueFromMap(mapValue(env.Version), "version")
}

func (a *app) addDelegatedVersionWarning(env *output.Envelope, diagnostics map[string]any) {
	windows := mapValue(diagnostics["windows"])
	windowsVersion := stringValueFromMap(windows, "xlflow_version")
	localVersion := strings.TrimSpace(a.buildInfo.Version)
	if windowsVersion == "" || localVersion == "" || windowsVersion == "dev" || localVersion == "dev" || windowsVersion == localVersion {
		return
	}
	warnings := anySlice(env.Warnings)
	warnings = append(warnings, map[string]any{
		"code":    "wsl_windows_version_mismatch",
		"message": fmt.Sprintf("WSL xlflow version %s differs from Windows xlflow version %s.", localVersion, windowsVersion),
	})
	env.Warnings = warnings
}

func delegatedDoctorBridgeFound(env output.Envelope, diagnostics map[string]any) bool {
	if env.Bridge != nil || stringValueFromMap(diagnostics, "selected_bridge") != "" {
		return true
	}
	if env.Error == nil {
		return false
	}
	return env.Error.Code != "bridge_not_available" &&
		env.Error.Code != "dotnet_missing" &&
		env.Error.Code != "dotnet_runtime_missing" &&
		env.Error.Code != "powershell_missing"
}

func delegatedDoctorExcelAvailable(env output.Envelope, diagnostics map[string]any) bool {
	excel := mapValue(diagnostics["excel"])
	if value, ok := excel["com_activation"].(bool); ok {
		return value
	}
	if value, ok := diagnostics["excel_installed"].(bool); ok {
		return value
	}
	return false
}

func ensureJSONFlag(args []string) []string {
	for _, arg := range args {
		if arg == "--json" {
			return append([]string{}, args...)
		}
	}
	result := append([]string{}, args...)
	return append(result, "--json")
}

func delegatedEnvironment() []string {
	env := append([]string{}, os.Environ()...)
	env = setEnvironmentValue(env, wsl.EnvDelegated, "1")
	env = setEnvironmentValue(env, "WSLENV", mergeWSLEnv(os.Getenv("WSLENV"),
		wsl.EnvDelegated,
		"XLFLOW_EXCEL_BRIDGE",
		"XLFLOW_MODE",
		"XLFLOW_NO_UPDATE_CHECK",
	))
	return env
}

func mergeWSLEnv(current string, names ...string) string {
	entries := strings.Split(current, ":")
	seen := make(map[string]struct{}, len(entries)+len(names))
	result := make([]string, 0, len(entries)+len(names))
	for _, entry := range entries {
		entry = strings.TrimSpace(entry)
		if entry == "" {
			continue
		}
		name := entry
		if before, _, ok := strings.Cut(entry, "/"); ok {
			name = before
		}
		seen[name] = struct{}{}
		result = append(result, entry)
	}
	for _, name := range names {
		if _, ok := seen[name]; ok {
			continue
		}
		seen[name] = struct{}{}
		result = append(result, name+"/w")
	}
	return strings.Join(result, ":")
}

func setEnvironmentValue(env []string, name string, value string) []string {
	prefix := name + "="
	for i, entry := range env {
		if strings.HasPrefix(entry, prefix) {
			env[i] = prefix + value
			return env
		}
	}
	return append(env, prefix+value)
}

func (a *app) writeWSLFailure(commandName string, err error) error {
	var wslErr *wsl.Error
	if errors.As(err, &wslErr) {
		return a.writeFailure(commandName, output.ExitEnvironment, wslErr.Code, wslErr)
	}
	return a.writeFailure(commandName, output.ExitEnvironment, "windows_xlflow_execution_failed", err)
}

func (a *app) writeWSLDoctorFailure(err error, windowsCWD string, windowsExecutable string, executableFound bool) error {
	code := "windows_xlflow_execution_failed"
	message := err.Error()
	var wslErr *wsl.Error
	if errors.As(err, &wslErr) {
		code = wslErr.Code
		message = wslErr.Message
	}
	env := output.Failure("doctor", output.Error{
		Code:    code,
		Message: message,
		Source:  "xlflow",
		Phase:   "wsl.delegate",
	})
	env.Diagnostics = map[string]any{
		"host": map[string]any{
			"os":      "linux",
			"is_wsl":  true,
			"distro":  wslDistroName(),
			"version": a.buildInfo.Version,
		},
		"windows": map[string]any{
			"xlflow_found":    executableFound,
			"xlflow_path":     windowsExecutable,
			"xlflow_version":  "",
			"bridge_found":    false,
			"excel_available": false,
		},
		"path_translation": map[string]any{
			"supported":    windowsCWD != "",
			"wsl_path":     a.cwd,
			"windows_path": windowsCWD,
		},
	}
	return a.write(env, output.ExitEnvironment)
}

func mapValue(value any) map[string]any {
	if value == nil {
		return map[string]any{}
	}
	if typed, ok := value.(map[string]any); ok {
		return typed
	}
	data, err := json.Marshal(value)
	if err != nil {
		return map[string]any{}
	}
	var result map[string]any
	if err := json.Unmarshal(data, &result); err != nil || result == nil {
		return map[string]any{}
	}
	return result
}

func stringValueFromMap(value map[string]any, key string) string {
	raw, ok := value[key]
	if !ok {
		return ""
	}
	text, _ := raw.(string)
	return strings.TrimSpace(text)
}

func anySlice(value any) []any {
	if value == nil {
		return nil
	}
	if typed, ok := value.([]any); ok {
		return typed
	}
	data, err := json.Marshal(value)
	if err != nil {
		return nil
	}
	var result []any
	if err := json.Unmarshal(data, &result); err != nil {
		return nil
	}
	return result
}
