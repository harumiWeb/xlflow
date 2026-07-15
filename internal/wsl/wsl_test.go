package wsl

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

func TestIsWSLFor(t *testing.T) {
	tests := []struct {
		name     string
		goos     string
		env      map[string]string
		version  string
		readErr  error
		expected bool
	}{
		{name: "non linux", goos: "windows", env: map[string]string{"WSL_INTEROP": "/run/WSL/1"}},
		{name: "interop", goos: "linux", env: map[string]string{"WSL_INTEROP": "/run/WSL/1"}, expected: true},
		{name: "distro", goos: "linux", env: map[string]string{"WSL_DISTRO_NAME": "Ubuntu"}, expected: true},
		{name: "proc version", goos: "linux", version: "Linux version 5.15.90.1-microsoft-standard-WSL2", expected: true},
		{name: "plain linux", goos: "linux", version: "Linux version 6.8.0"},
		{name: "proc unavailable", goos: "linux", readErr: os.ErrNotExist},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			got := IsWSLFor(
				test.goos,
				func(key string) string { return test.env[key] },
				func(string) ([]byte, error) { return []byte(test.version), test.readErr },
			)
			if got != test.expected {
				t.Fatalf("IsWSLFor() = %v, want %v", got, test.expected)
			}
		})
	}
}

func TestIsWindowsDriveMount(t *testing.T) {
	tests := map[string]bool{
		"/mnt/c":                   true,
		"/mnt/d/dev/project":       true,
		"/mnt/C/日本語 project":       true,
		"/home/user/project":       false,
		"/mnt":                     false,
		"/mnt/shared/project":      false,
		`C:\dev\project`:           false,
		"/mnt/candidate/not-drive": false,
	}
	for path, expected := range tests {
		if got := IsWindowsDriveMount(path); got != expected {
			t.Errorf("IsWindowsDriveMount(%q) = %v, want %v", path, got, expected)
		}
	}
}

func TestToWindowsPathWith(t *testing.T) {
	got, err := ToWindowsPathWith(context.Background(), "/mnt/c/dev/日本語 project", func(_ context.Context, name string, args ...string) ([]byte, error) {
		if name != "wslpath" || !reflect.DeepEqual(args, []string{"-w", "/mnt/c/dev/日本語 project"}) {
			t.Fatalf("command = %s %v", name, args)
		}
		return []byte("C:\\dev\\日本語 project\r\n"), nil
	})
	if err != nil {
		t.Fatalf("ToWindowsPathWith() error = %v", err)
	}
	if got != `C:\dev\日本語 project` {
		t.Fatalf("ToWindowsPathWith() = %q", got)
	}
}

func TestToWindowsPathWithRejectsWSLOnlyPath(t *testing.T) {
	_, err := ToWindowsPathWith(context.Background(), "/home/user/project", nil)
	assertWSLErrorCode(t, err, "wsl_project_path_unsupported")
}

func TestTranslateArgsWith(t *testing.T) {
	translate := func(_ context.Context, path string) (string, error) {
		return "WIN[" + path + "]", nil
	}
	args := []string{
		"--json",
		"run",
		"/mnt/c/dev/Book.xlsm",
		"--input=/mnt/d/input.xlsm",
		"--save-as",
		"/mnt/c/out/Result.xlsm",
		"--src=/mnt/c/source/Book.xlsm",
		"--template",
		"/mnt/c/templates/Base.xlsm",
		"--vba-before=/mnt/c/vba/before",
		"--vba-after",
		"/mnt/c/vba/after",
		"--filedialog",
		"get-open:source=/mnt/c/data/input.csv",
		"--arg",
		"string:/home/user/not-a-path-argument",
		"--value",
		"/home/user/literal-cell-value",
		"--msgbox=confirm=/home/user/literal-response",
	}
	got, err := TranslateArgsWith(context.Background(), args, translate)
	if err != nil {
		t.Fatalf("TranslateArgsWith() error = %v", err)
	}
	want := []string{
		"--json",
		"run",
		"WIN[/mnt/c/dev/Book.xlsm]",
		"--input=WIN[/mnt/d/input.xlsm]",
		"--save-as",
		"WIN[/mnt/c/out/Result.xlsm]",
		"--src=WIN[/mnt/c/source/Book.xlsm]",
		"--template",
		"WIN[/mnt/c/templates/Base.xlsm]",
		"--vba-before=WIN[/mnt/c/vba/before]",
		"--vba-after",
		"WIN[/mnt/c/vba/after]",
		"--filedialog",
		"get-open:source=WIN[/mnt/c/data/input.csv]",
		"--arg",
		"string:/home/user/not-a-path-argument",
		"--value",
		"/home/user/literal-cell-value",
		"--msgbox=confirm=/home/user/literal-response",
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("TranslateArgsWith() = %#v, want %#v", got, want)
	}
}

func TestTranslateArgsWithRejectsWSLOnlyAbsolutePath(t *testing.T) {
	_, err := TranslateArgsWith(context.Background(), []string{"run", "--input", "/home/user/Book.xlsm"}, func(context.Context, string) (string, error) {
		return "", nil
	})
	assertWSLErrorCode(t, err, "wsl_project_path_unsupported")
}

func TestResolveWindowsExecutableWithPrefersEnvironment(t *testing.T) {
	dir := t.TempDir()
	executable := filepath.Join(dir, "xlflow.exe")
	if err := os.WriteFile(executable, []byte("test"), 0o755); err != nil {
		t.Fatal(err)
	}
	got, windowsPath, err := ResolveWindowsExecutableWith(
		context.Background(),
		func(key string) string {
			if key == EnvWindowsExecutable {
				return `C:\tools\xlflow.exe`
			}
			return ""
		},
		func(string) (string, error) {
			return "", errors.New("PATH must not be used")
		},
		os.Stat,
		func(context.Context, string) (string, error) { return executable, nil },
		func(context.Context, string) (string, error) { return "", errors.New("not needed") },
	)
	if err != nil {
		t.Fatalf("ResolveWindowsExecutableWith() error = %v", err)
	}
	if got != executable || windowsPath != `C:\tools\xlflow.exe` {
		t.Fatalf("ResolveWindowsExecutableWith() = (%q, %q)", got, windowsPath)
	}
}

func TestResolveWindowsExecutableWithUsesPath(t *testing.T) {
	dir := t.TempDir()
	executable := filepath.Join(dir, "xlflow.exe")
	if err := os.WriteFile(executable, []byte("test"), 0o755); err != nil {
		t.Fatal(err)
	}
	got, windowsPath, err := ResolveWindowsExecutableWith(
		context.Background(),
		func(string) string { return "" },
		func(name string) (string, error) {
			if name != "xlflow.exe" {
				t.Fatalf("LookPath(%q)", name)
			}
			return executable, nil
		},
		os.Stat,
		nil,
		func(_ context.Context, path string) (string, error) {
			return `C:\tools\xlflow.exe`, nil
		},
	)
	if err != nil {
		t.Fatalf("ResolveWindowsExecutableWith() error = %v", err)
	}
	if got != executable || windowsPath != `C:\tools\xlflow.exe` {
		t.Fatalf("ResolveWindowsExecutableWith() = (%q, %q)", got, windowsPath)
	}
}

func TestResolveWindowsExecutableWithRejectsMissingAndNonExe(t *testing.T) {
	t.Run("missing", func(t *testing.T) {
		_, _, err := ResolveWindowsExecutableWith(
			context.Background(),
			func(string) string { return "" },
			func(string) (string, error) { return "", os.ErrNotExist },
			os.Stat,
			nil,
			nil,
		)
		assertWSLErrorCode(t, err, "windows_xlflow_not_found")
	})

	t.Run("non exe", func(t *testing.T) {
		path := "/mnt/c/tools/xlflow"
		_, _, err := ResolveWindowsExecutableWith(
			context.Background(),
			func(key string) string {
				if key == EnvWindowsExecutable {
					return path
				}
				return ""
			},
			nil,
			func(candidate string) (os.FileInfo, error) {
				t.Fatalf("stat should not be called for non-exe candidate %q", candidate)
				return nil, nil
			},
			nil,
			nil,
		)
		assertWSLErrorCode(t, err, "windows_xlflow_not_found")
	})
}

func assertWSLErrorCode(t *testing.T, err error, code string) {
	t.Helper()
	var wslErr *Error
	if !errors.As(err, &wslErr) {
		t.Fatalf("error = %v, want *Error", err)
	}
	if wslErr.Code != code {
		t.Fatalf("error code = %q, want %q", wslErr.Code, code)
	}
}
