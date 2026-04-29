package scripts_test

import (
	"os/exec"
	"path/filepath"
	"testing"
)

func TestPowerShellScriptsParse(t *testing.T) {
	scripts := []string{"common.ps1", "doctor.ps1", "pull.ps1", "push.ps1", "run.ps1"}
	for _, script := range scripts {
		script := script
		t.Run(script, func(t *testing.T) {
			path := filepath.Join(".", script)
			cmd := exec.Command("pwsh", "-NoProfile", "-Command", "try { [scriptblock]::Create((Get-Content -Raw -LiteralPath '"+path+"')) | Out-Null } catch { Write-Error $_; exit 1 }")
			out, err := cmd.CombinedOutput()
			if err != nil {
				t.Fatalf("script parse failed: %v\n%s", err, out)
			}
		})
	}
}
