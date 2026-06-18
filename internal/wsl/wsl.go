package wsl

import (
	"context"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
)

const (
	EnvWindowsExecutable = "XLFLOW_WINDOWS_EXE"
	EnvDelegated         = "XLFLOW_WINDOWS_DELEGATED"
)

type Error struct {
	Code    string
	Message string
	Err     error
}

func (e *Error) Error() string {
	if e == nil {
		return ""
	}
	return e.Message
}

func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Err
}

type CommandRunner func(context.Context, string, ...string) ([]byte, error)

type LookPathFunc func(string) (string, error)

type StatFunc func(string) (os.FileInfo, error)

func IsWSL() bool {
	return IsWSLFor(runtime.GOOS, os.Getenv, os.ReadFile)
}

func IsWSLFor(goos string, getenv func(string) string, readFile func(string) ([]byte, error)) bool {
	if goos != "linux" {
		return false
	}
	if strings.TrimSpace(getenv("WSL_INTEROP")) != "" || strings.TrimSpace(getenv("WSL_DISTRO_NAME")) != "" {
		return true
	}
	data, err := readFile("/proc/version")
	if err != nil {
		return false
	}
	version := strings.ToLower(string(data))
	return strings.Contains(version, "microsoft") || strings.Contains(version, "wsl")
}

func DistroName() string {
	return strings.TrimSpace(os.Getenv("WSL_DISTRO_NAME"))
}

func IsWindowsDriveMount(path string) bool {
	clean := filepath.ToSlash(filepath.Clean(path))
	if len(clean) < len("/mnt/c") || !strings.HasPrefix(clean, "/mnt/") {
		return false
	}
	rest := strings.TrimPrefix(clean, "/mnt/")
	if len(rest) < 1 || !isASCIIAlpha(rest[0]) {
		return false
	}
	return len(rest) == 1 || rest[1] == '/'
}

func ToWindowsPath(ctx context.Context, path string) (string, error) {
	return ToWindowsPathWith(ctx, path, runCommand)
}

func ToWindowsPathWith(ctx context.Context, path string, run CommandRunner) (string, error) {
	if !IsWindowsDriveMount(path) {
		return "", &Error{
			Code:    "wsl_project_path_unsupported",
			Message: fmt.Sprintf("path %q is not under a Windows-mounted drive such as /mnt/c", path),
		}
	}
	return runWSLPath(ctx, run, "-w", path)
}

func ToUnixPath(ctx context.Context, path string) (string, error) {
	return ToUnixPathWith(ctx, path, runCommand)
}

func ToUnixPathWith(ctx context.Context, path string, run CommandRunner) (string, error) {
	return runWSLPath(ctx, run, "-u", path)
}

func ResolveWindowsExecutable(ctx context.Context) (string, string, error) {
	return ResolveWindowsExecutableWith(
		ctx,
		os.Getenv,
		exec.LookPath,
		os.Stat,
		func(ctx context.Context, path string) (string, error) {
			return ToUnixPath(ctx, path)
		},
		func(ctx context.Context, path string) (string, error) {
			return ToWindowsPath(ctx, path)
		},
	)
}

func ResolveWindowsExecutableWith(
	ctx context.Context,
	getenv func(string) string,
	lookPath LookPathFunc,
	stat StatFunc,
	toUnix func(context.Context, string) (string, error),
	toWindows func(context.Context, string) (string, error),
) (string, string, error) {
	candidate := strings.TrimSpace(getenv(EnvWindowsExecutable))
	if candidate != "" {
		resolved := candidate
		if isWindowsAbsolutePath(candidate) {
			var err error
			resolved, err = toUnix(ctx, candidate)
			if err != nil {
				return "", "", &Error{
					Code:    "wsl_path_translation_failed",
					Message: fmt.Sprintf("failed to translate %s to a WSL path", EnvWindowsExecutable),
					Err:     err,
				}
			}
		}
		if err := validateExecutable(resolved, stat); err != nil {
			return "", "", err
		}
		windowsPath := candidate
		if !isWindowsAbsolutePath(candidate) {
			var err error
			windowsPath, err = toWindows(ctx, resolved)
			if err != nil {
				return "", "", &Error{
					Code:    "wsl_path_translation_failed",
					Message: fmt.Sprintf("failed to translate %s to a Windows path", EnvWindowsExecutable),
					Err:     err,
				}
			}
		}
		return resolved, windowsPath, nil
	}

	resolved, err := lookPath("xlflow.exe")
	if err != nil {
		return "", "", &Error{
			Code:    "windows_xlflow_not_found",
			Message: "Windows xlflow.exe was not found; install xlflow on Windows or set XLFLOW_WINDOWS_EXE",
			Err:     err,
		}
	}
	if err := validateExecutable(resolved, stat); err != nil {
		return "", "", err
	}
	windowsPath, err := toWindows(ctx, resolved)
	if err != nil {
		return "", "", &Error{
			Code:    "wsl_path_translation_failed",
			Message: "failed to translate the resolved Windows xlflow executable path",
			Err:     err,
		}
	}
	return resolved, windowsPath, nil
}

func TranslateArgs(ctx context.Context, args []string) ([]string, error) {
	return TranslateArgsWith(ctx, args, func(ctx context.Context, path string) (string, error) {
		return ToWindowsPath(ctx, path)
	})
}

func TranslateArgsWith(ctx context.Context, args []string, translate func(context.Context, string) (string, error)) ([]string, error) {
	translated := make([]string, len(args))
	for i := 0; i < len(args); i++ {
		arg := args[i]
		value, err := translateArgument(ctx, arg, translate)
		if err != nil {
			return nil, err
		}
		translated[i] = value
		if flagName, inlineValue, hasInlineValue := splitLongFlag(arg); flagName != "" {
			switch {
			case hasInlineValue && isPathFlag(flagName):
				translatedValue, err := translatePathValue(ctx, inlineValue, translate)
				if err != nil {
					return nil, err
				}
				translated[i] = flagName + "=" + translatedValue
			case hasInlineValue && flagName == "--filedialog":
				translatedValue, err := translateFileDialogValue(ctx, inlineValue, translate)
				if err != nil {
					return nil, err
				}
				translated[i] = flagName + "=" + translatedValue
			case !hasInlineValue && flagTakesValue(flagName) && i+1 < len(args):
				i++
				next := args[i]
				switch {
				case isPathFlag(flagName):
					translated[i], err = translatePathValue(ctx, next, translate)
				case flagName == "--filedialog":
					translated[i], err = translateFileDialogValue(ctx, next, translate)
				default:
					translated[i] = next
				}
				if err != nil {
					return nil, err
				}
			}
		}
	}
	return translated, nil
}

func translateArgument(ctx context.Context, arg string, translate func(context.Context, string) (string, error)) (string, error) {
	if isWSLAbsolutePath(arg) {
		return translateAbsoluteArgument(ctx, arg, translate)
	}
	return arg, nil
}

func translatePathValue(ctx context.Context, value string, translate func(context.Context, string) (string, error)) (string, error) {
	if isWSLAbsolutePath(value) {
		return translateAbsoluteArgument(ctx, value, translate)
	}
	return value, nil
}

func translateFileDialogValue(ctx context.Context, value string, translate func(context.Context, string) (string, error)) (string, error) {
	if prefix, path, ok := strings.Cut(value, "=/"); ok {
		translated, err := translateAbsoluteArgument(ctx, "/"+path, translate)
		if err != nil {
			return "", err
		}
		return prefix + "=" + translated, nil
	}
	return value, nil
}

func splitLongFlag(arg string) (string, string, bool) {
	if !strings.HasPrefix(arg, "--") {
		return "", "", false
	}
	name, value, ok := strings.Cut(arg, "=")
	return name, value, ok
}

func isPathFlag(name string) bool {
	switch name {
	case "--input", "--out", "--output-dir", "--path", "--save-as":
		return true
	default:
		return false
	}
}

func flagTakesValue(name string) bool {
	switch name {
	case "--address", "--agent", "--arg", "--backup", "--bridge", "--cell", "--clear",
		"--columns", "--events", "--filedialog", "--fill", "--filter", "--format", "--from",
		"--formula", "--id", "--initializer", "--input", "--inputbox", "--line-numbers", "--macro",
		"--module", "--msgbox", "--name", "--out", "--output-dir", "--range", "--rows",
		"--path", "--save-as", "--sheet", "--tag", "--target", "--text", "--timeout", "--to", "--value",
		"--vba-after", "--vba-before":
		return true
	default:
		return false
	}
}

func translateAbsoluteArgument(ctx context.Context, path string, translate func(context.Context, string) (string, error)) (string, error) {
	if !IsWindowsDriveMount(path) {
		return "", &Error{
			Code:    "wsl_project_path_unsupported",
			Message: fmt.Sprintf("absolute WSL path %q is not visible to Windows; use a project path under /mnt/<drive>", path),
		}
	}
	translated, err := translate(ctx, path)
	if err != nil {
		var wslErr *Error
		if errors.As(err, &wslErr) {
			return "", err
		}
		return "", &Error{
			Code:    "wsl_path_translation_failed",
			Message: fmt.Sprintf("failed to translate WSL path %q", path),
			Err:     err,
		}
	}
	return strings.TrimSpace(translated), nil
}

func runWSLPath(ctx context.Context, run CommandRunner, mode string, path string) (string, error) {
	output, err := run(ctx, "wslpath", mode, path)
	if err != nil {
		return "", &Error{
			Code:    "wsl_path_translation_failed",
			Message: fmt.Sprintf("wslpath %s failed for %q", mode, path),
			Err:     err,
		}
	}
	translated := strings.TrimSpace(string(output))
	if translated == "" {
		return "", &Error{
			Code:    "wsl_path_translation_failed",
			Message: fmt.Sprintf("wslpath %s returned an empty path for %q", mode, path),
		}
	}
	return translated, nil
}

func runCommand(ctx context.Context, name string, args ...string) ([]byte, error) {
	return exec.CommandContext(ctx, name, args...).Output()
}

func validateExecutable(path string, stat StatFunc) error {
	if !strings.EqualFold(filepath.Ext(path), ".exe") {
		return &Error{
			Code:    "windows_xlflow_not_found",
			Message: fmt.Sprintf("Windows xlflow path %q must reference an .exe file", path),
		}
	}
	info, err := stat(path)
	if err != nil {
		return &Error{
			Code:    "windows_xlflow_not_found",
			Message: fmt.Sprintf("Windows xlflow executable %q was not found", path),
			Err:     err,
		}
	}
	if info.IsDir() {
		return &Error{
			Code:    "windows_xlflow_not_found",
			Message: fmt.Sprintf("Windows xlflow path %q is a directory", path),
		}
	}
	return nil
}

func isWindowsAbsolutePath(path string) bool {
	return len(path) >= 3 && isASCIIAlpha(path[0]) && path[1] == ':' && (path[2] == '\\' || path[2] == '/')
}

func isWSLAbsolutePath(path string) bool {
	return strings.HasPrefix(path, "/")
}

func isASCIIAlpha(value byte) bool {
	return value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z'
}
