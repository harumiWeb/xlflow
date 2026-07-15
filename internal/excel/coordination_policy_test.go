package excel

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestBridgeInvocationWithoutCoordinationPolicyFailsBeforeExecution(t *testing.T) {
	env, code, err := (Runner{}).run("future-bridge-command", map[string]string{"Action": "future"})
	if err != nil {
		t.Fatalf("run() error = %v", err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Error == nil || env.Error.Code != coordination.MissingPolicyCode || env.Error.Phase != "coordination.policy" {
		t.Fatalf("error envelope = %#v", env.Error)
	}
}
