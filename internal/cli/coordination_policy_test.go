package cli

import (
	"bytes"
	"encoding/json"
	"errors"
	"testing"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestExecutableCommandsHaveCoordinationPolicies(t *testing.T) {
	root := (&app{}).rootCommand()
	var visit func(*cobra.Command)
	visit = func(cmd *cobra.Command) {
		if (cmd.Run != nil || cmd.RunE != nil) && !isGeneratedCobraCommand(cmd) {
			if _, err := coordination.LookupCLI(cmd.CommandPath()); err != nil {
				t.Errorf("%s: %v", cmd.CommandPath(), err)
			}
		}
		for _, child := range cmd.Commands() {
			visit(child)
		}
	}
	visit(root)
}

func TestMissingCoordinationPolicyRendersStructuredFailure(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{stdout: &stdout, stderr: &bytes.Buffer{}, json: true}
	root := &cobra.Command{Use: "xlflow"}
	unknown := &cobra.Command{Use: "future", RunE: func(*cobra.Command, []string) error { return nil }}
	root.AddCommand(unknown)

	err := a.requireCoordinationPolicy(unknown)
	if output.ExitCode(err) != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", output.ExitCode(err), output.ExitEnvironment)
	}
	var missing *coordination.MissingPolicyError
	if !errors.As(err, &missing) {
		t.Fatalf("error = %v, want MissingPolicyError", err)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("decode failure envelope: %v\n%s", err, stdout.String())
	}
	if env.Error == nil || env.Error.Code != coordination.MissingPolicyCode {
		t.Fatalf("error envelope = %#v", env.Error)
	}
}
