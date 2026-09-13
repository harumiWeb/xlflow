package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestCapabilitiesCommandWritesV1JSONEnvelope(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "capabilities"})

	if err := root.Execute(); err != nil {
		t.Fatalf("capabilities command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Status       string                    `json:"status"`
		Command      string                    `json:"command"`
		Capabilities coordination.Capabilities `json:"capabilities"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("decode capabilities JSON: %v\n%s", err, stdout.String())
	}
	if got.Status != output.StatusOK || got.Command != "capabilities" {
		t.Fatalf("envelope = %#v", got)
	}
	if got.Capabilities.CapabilityVersion != coordination.CapabilityVersion {
		t.Fatalf("capability version = %d, want %d", got.Capabilities.CapabilityVersion, coordination.CapabilityVersion)
	}
	push, ok := got.Capabilities.Commands["push"]
	if !ok {
		t.Fatal("push capability missing")
	}
	if len(push.CLIPaths) != 1 || push.CLIPaths[0] != "push" || push.ResourceScope != coordination.ResourceWorkbook || push.OperationKind != coordination.OperationMutate || push.ParallelSafe || !push.RetryableWhenBusy || push.DefaultWaitPolicy != coordination.WaitFail || push.RecoveryBehavior != coordination.RecoveryBlock || !push.RequiresExcel {
		t.Fatalf("push capability = %#v", push)
	}
	check, ok := got.Capabilities.Commands["encoding.check"]
	if !ok || len(check.CLIPaths) != 1 || check.CLIPaths[0] != "encoding check" || check.ResourceScope != coordination.ResourceNone || check.OperationKind != coordination.OperationRead || !check.ParallelSafe || check.RequiresExcel {
		t.Fatalf("encoding.check capability = %#v", check)
	}
	convert, ok := got.Capabilities.Commands["encoding.convert"]
	if !ok || len(convert.CLIPaths) != 1 || convert.CLIPaths[0] != "encoding convert" || convert.ResourceScope != coordination.ResourceWorkbook || convert.OperationKind != coordination.OperationMutate || convert.ParallelSafe || !convert.RetryableWhenBusy || convert.DefaultWaitPolicy != coordination.WaitFail || convert.RecoveryBehavior != coordination.RecoveryNotApplicable || convert.RequiresExcel {
		t.Fatalf("encoding.convert capability = %#v", convert)
	}
}

func TestCapabilitiesCommandDoesNotRequireAWorkbook(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "capabilities"})

	if err := root.Execute(); err != nil {
		t.Fatalf("capabilities command without project error = %v", err)
	}
}
