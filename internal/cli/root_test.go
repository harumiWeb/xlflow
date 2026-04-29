package cli

import "testing"

func TestRootCommandIncludesTestCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"test"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "test" {
		t.Fatalf("expected test command, got %#v", cmd)
	}
	if flag := cmd.Flags().Lookup("filter"); flag == nil {
		t.Fatal("expected test command to define --filter")
	}
}
