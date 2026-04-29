package excel

import "testing"

func TestScriptPathFindsRepositoryScripts(t *testing.T) {
	path, err := scriptPath(t.TempDir(), "run")
	if err != nil {
		t.Fatal(err)
	}
	if path == "" {
		t.Fatal("expected script path")
	}
}
