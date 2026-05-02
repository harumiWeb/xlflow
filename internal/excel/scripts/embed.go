package scripts

import (
	"embed"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
)

//go:embed *.ps1
var bundled embed.FS

// Materialize writes a full embedded runtime-script bundle into a per-call temp
// directory. The directory must stay unique to the caller because cleanup
// removes the whole tree, and concurrent xlflow commands must never share it.
func Materialize(commandName string) (string, func(), error) {
	name := commandName + ".ps1"

	scriptNames, err := fs.Glob(bundled, "*.ps1")
	if err != nil {
		return "", nil, fmt.Errorf("failed to enumerate embedded PowerShell scripts: %w", err)
	}
	sort.Strings(scriptNames)
	if len(scriptNames) == 0 {
		return "", nil, fmt.Errorf("no embedded xlflow PowerShell scripts were found")
	}

	found := false
	for _, scriptName := range scriptNames {
		if scriptName == name {
			found = true
			break
		}
	}
	if !found {
		return "", nil, fmt.Errorf("embedded script %s was not found", name)
	}

	dir, err := os.MkdirTemp("", "xlflow-scripts-*")
	if err != nil {
		return "", nil, fmt.Errorf("failed to create temporary script directory: %w", err)
	}
	cleanup := func() {
		_ = os.RemoveAll(dir)
	}

	for _, scriptName := range scriptNames {
		body, err := bundled.ReadFile(scriptName)
		if err != nil {
			cleanup()
			return "", nil, fmt.Errorf("failed to read embedded script %s: %w", scriptName, err)
		}
		if err := os.WriteFile(filepath.Join(dir, scriptName), body, 0o600); err != nil {
			cleanup()
			return "", nil, fmt.Errorf("failed to write embedded script %s: %w", scriptName, err)
		}
	}

	return filepath.Join(dir, name), cleanup, nil
}
