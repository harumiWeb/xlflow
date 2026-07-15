package coordination

import (
	"errors"
	"reflect"
	"sort"
	"strings"
	"testing"
)

func TestPolicyValidation(t *testing.T) {
	valid := Policy{ResourceScope: ResourceWorkbook, OperationKind: OperationExecute, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail}
	if err := valid.Validate(); err != nil {
		t.Fatalf("valid policy: %v", err)
	}

	tests := []Policy{
		{ResourceScope: "invalid", OperationKind: OperationRead, DefaultWaitPolicy: WaitFail},
		{ResourceScope: ResourceNone, OperationKind: "invalid", DefaultWaitPolicy: WaitFail},
		{ResourceScope: ResourceNone, OperationKind: OperationRead, DefaultWaitPolicy: "invalid"},
		{ResourceScope: ResourceNone, OperationKind: OperationRead, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail},
		{ResourceScope: ResourceWorkbook, OperationKind: OperationRead, ParallelSafe: true, RetryableWhenBusy: true, DefaultWaitPolicy: WaitFail},
	}
	for _, policy := range tests {
		if err := policy.Validate(); err == nil {
			t.Errorf("Policy.Validate() unexpectedly accepted %+v", policy)
		}
	}
}

func TestRegistryIsValid(t *testing.T) {
	if err := ValidateRegistry(); err != nil {
		t.Fatal(err)
	}
}

func TestLookupByIDAndCLI(t *testing.T) {
	byID, err := Lookup("run")
	if err != nil {
		t.Fatal(err)
	}
	byCLI, err := LookupCLI("  xlflow   RUN ")
	if err != nil {
		t.Fatal(err)
	}
	if byID.ID != byCLI.ID || byCLI.Policy.OperationKind != OperationExecute || byCLI.Policy.ParallelSafe {
		t.Fatalf("unexpected run policy: %+v", byCLI)
	}

	status, err := LookupCLI("status")
	if err != nil {
		t.Fatal(err)
	}
	if !status.Policy.ParallelSafe || status.Policy.ResourceScope != ResourceWorkbook {
		t.Fatalf("status should be a workbook observer: %+v", status.Policy)
	}

	lint, err := LookupCLI("lint")
	if err != nil {
		t.Fatal(err)
	}
	if lint.Policy.ResourceScope != ResourceNone {
		t.Fatalf("lint should be source-only: %+v", lint.Policy)
	}
}

func TestLookupBridgeUsesPayloadSelectors(t *testing.T) {
	tests := []struct {
		command string
		args    map[string]string
		want    CommandID
	}{
		{"run", nil, "run"},
		{"FORM-WRITE", map[string]string{"action": "BUILD", "SpecPath": "form.yaml"}, "form.build"},
		{"edit", map[string]string{"Action": "formula"}, "edit.formula"},
		{"inspect", map[string]string{"Target": "used-range"}, "inspect.used-range"},
		{"session", map[string]string{"Action": "status"}, "session.status"},
	}
	for _, test := range tests {
		got, err := LookupBridge(test.command, test.args)
		if err != nil {
			t.Errorf("LookupBridge(%q): %v", test.command, err)
			continue
		}
		if got.ID != test.want {
			t.Errorf("LookupBridge(%q) ID = %q, want %q", test.command, got.ID, test.want)
		}
	}
}

func TestUnknownSelectorsFailClosed(t *testing.T) {
	tests := []func() error{
		func() error { _, err := Lookup("future.command"); return err },
		func() error { _, err := LookupCLI("future command"); return err },
		func() error { _, err := LookupBridge("edit", map[string]string{"Action": "future"}); return err },
		func() error { _, err := LookupBridge("session", nil); return err },
	}
	for _, lookup := range tests {
		err := lookup()
		var missing *MissingPolicyError
		if !errors.As(err, &missing) {
			t.Fatalf("error = %v, want MissingPolicyError", err)
		}
		if missing.Code() != MissingPolicyCode || !strings.Contains(missing.Error(), MissingPolicyCode) {
			t.Fatalf("unexpected missing policy error: %v", missing)
		}
	}
}

func TestAllIsStableAndDefensive(t *testing.T) {
	first := All()
	second := All()
	if !reflect.DeepEqual(first, second) {
		t.Fatal("All() is not stable")
	}
	if !sort.SliceIsSorted(first, func(i, j int) bool { return first[i].ID < first[j].ID }) {
		t.Fatal("All() is not sorted by command ID")
	}

	index := -1
	for i := range first {
		if len(first[i].Bridge) > 0 && len(first[i].Bridge[0].Args) > 0 {
			index = i
			break
		}
	}
	if index < 0 {
		t.Fatal("test requires a descriptor with bridge args")
	}
	first[index].CLI[0].Path = "changed"
	for key := range first[index].Bridge[0].Args {
		first[index].Bridge[0].Args[key] = "changed"
	}
	mutatedID := first[index].ID
	fresh, err := Lookup(mutatedID)
	if err != nil {
		t.Fatal(err)
	}
	if fresh.CLI[0].Path == "changed" {
		t.Fatal("CLI selector mutation escaped defensive copy")
	}
	for _, value := range fresh.Bridge[0].Args {
		if value == "changed" {
			t.Fatal("bridge selector mutation escaped defensive copy")
		}
	}
}

func TestImportantCommandsHaveConservativePolicies(t *testing.T) {
	tests := []struct {
		path string
		kind OperationKind
	}{
		{"push", OperationMutate},
		{"run", OperationExecute},
		{"test", OperationExecute},
		{"form build", OperationDesigner},
		{"form apply", OperationDesigner},
		{"form snapshot", OperationDesigner},
		{"inspect form", OperationDesigner},
	}
	for _, test := range tests {
		descriptor, err := LookupCLI(test.path)
		if err != nil {
			t.Fatal(err)
		}
		if descriptor.Policy.ResourceScope != ResourceWorkbook || descriptor.Policy.OperationKind != test.kind || descriptor.Policy.ParallelSafe {
			t.Errorf("%s policy = %+v", test.path, descriptor.Policy)
		}
	}
}
