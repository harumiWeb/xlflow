package procedureir

import (
	"strings"
	"testing"
)

func TestSafeProbeResultInspectionRecognizesOnlyIsArrayOfProbeTarget(t *testing.T) {
	tests := []struct {
		name        string
		value       string
		probeTarget string
		want        bool
	}{
		{name: "matching target", value: "IsArray(testVal)", probeTarget: "testVal", want: true},
		{name: "case insensitive", value: "isarray( testVal )", probeTarget: "TESTVAL", want: true},
		{name: "different target", value: "IsArray(other)", probeTarget: "testVal", want: false},
		{name: "different intrinsic", value: "IsObject(testVal)", probeTarget: "testVal", want: false},
		{name: "compound expression", value: "IsArray(testVal) And True", probeTarget: "testVal", want: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := SafeProbeResultInspection(test.value, test.probeTarget); got != test.want {
				t.Fatalf("SafeProbeResultInspection(%q, %q) = %v, want %v", test.value, test.probeTarget, got, test.want)
			}
		})
	}
}

func TestSafeBooleanComparisonRequiresProvenNonNullableOperands(t *testing.T) {
	safeOperands := map[string]bool{"left": true, "right": true}
	operandIsSafe := func(name string) bool { return safeOperands[name] }
	tests := []struct {
		name  string
		value string
		check func(string) bool
		want  bool
	}{
		{name: "known operands", value: "left = right", check: operandIsSafe, want: true},
		{name: "parenthesized known operands", value: "(left = right)", check: operandIsSafe, want: true},
		{name: "unknown variant operand", value: "left = nullable", check: operandIsSafe, want: false},
		{name: "Null literal", value: "left = Null", check: operandIsSafe, want: false},
		{name: "member access", value: "obj.Value = right", check: operandIsSafe, want: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := SafeBooleanComparison(test.value, test.check); got != test.want {
				t.Fatalf("SafeBooleanComparison(%q) = %v, want %v", test.value, got, test.want)
			}
		})
	}
}

func TestSafeBooleanComparisonOperandTypeRejectsNullableAndUnknownValues(t *testing.T) {
	tests := []struct {
		name     string
		typeName string
		array    bool
		object   bool
		want     bool
	}{
		{name: "String", typeName: "String", want: true},
		{name: "Long", typeName: "Long", want: true},
		{name: "LongPtr", typeName: "LongPtr", want: true},
		{name: "Variant", typeName: "Variant", want: false},
		{name: "Object", typeName: "Object", object: true, want: false},
		{name: "array", typeName: "String", array: true, want: false},
		{name: "user type", typeName: "CustomValue", want: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := SafeBooleanComparisonOperandType(test.typeName, test.array, test.object); got != test.want {
				t.Fatalf("SafeBooleanComparisonOperandType(%q, %v, %v) = %v, want %v", test.typeName, test.array, test.object, got, test.want)
			}
		})
	}
}

func TestSafeBooleanCoercionProbeRecognizesCheckedDictionaryConversions(t *testing.T) {
	objectTargets := func(name string) bool { return strings.EqualFold(name, "envelope") }
	tests := []struct {
		name  string
		value string
		want  bool
	}{
		{name: "string key", value: `CBool(envelope("defer"))`, want: true},
		{name: "identifier key", value: "CBool(envelope(key))", want: false},
		{name: "direct value is not keyed probe", value: "CBool(value)", want: false},
		{name: "member chain", value: "CBool(envelope.Item(\"defer\"))", want: false},
		{name: "helper call", value: "CBool(ReadValue(envelope))", want: false},
		{name: "compound expression", value: `CBool(envelope("defer")) Or True`, want: false},
		{name: "unresolved helper with string key", value: `CBool(ReadValue("defer"))`, want: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := SafeBooleanCoercionProbe(test.value, objectTargets); got != test.want {
				t.Fatalf("SafeBooleanCoercionProbe(%q) = %v, want %v", test.value, got, test.want)
			}
		})
	}
	if SafeBooleanCoercionProbe(`CBool(envelope("defer"))`, nil) {
		t.Fatal("SafeBooleanCoercionProbe accepted a probe without an Object declaration proof")
	}
}

func TestSafeArrayBoundsProbeAssignmentRecognizesOnlyTwoStepShape(t *testing.T) {
	tests := []struct {
		name            string
		target          string
		value           string
		wantArray       string
		wantLowerTarget string
		wantKind        SafeArrayBoundsProbeKind
		wantOK          bool
	}{
		{
			name:            "lower bound",
			target:          "lowerBound",
			value:           "LBound(values)",
			wantArray:       "values",
			wantLowerTarget: "lowerbound",
			wantKind:        SafeArrayBoundsProbeLowerBound,
			wantOK:          true,
		},
		{
			name:            "length from lower bound",
			target:          "itemCount",
			value:           "UBound(values) - lowerBound + 1",
			wantArray:       "values",
			wantLowerTarget: "lowerbound",
			wantKind:        SafeArrayBoundsProbeLength,
			wantOK:          true,
		},
		{
			name:     "compound value",
			target:   "itemCount",
			value:    "UBound(values) - lowerBound + 1 + 0",
			wantKind: SafeArrayBoundsProbeNone,
			wantOK:   false,
		},
		{
			name:     "different lower target syntax",
			target:   "itemCount",
			value:    "UBound(values) - LBound(values) + 1",
			wantKind: SafeArrayBoundsProbeNone,
			wantOK:   false,
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			array, lowerTarget, kind, ok := SafeArrayBoundsProbeAssignment(test.target, test.value)
			if array != test.wantArray || lowerTarget != test.wantLowerTarget || kind != test.wantKind || ok != test.wantOK {
				t.Fatalf("SafeArrayBoundsProbeAssignment(%q, %q) = (%q, %q, %v, %v), want (%q, %q, %v, %v)", test.target, test.value, array, lowerTarget, kind, ok, test.wantArray, test.wantLowerTarget, test.wantKind, test.wantOK)
			}
		})
	}
}

func TestErrorNumberThenBranchIsFailureRecognizesPolarityAndMasksStrings(t *testing.T) {
	tests := []struct {
		name        string
		condition   string
		thenFailure bool
		known       bool
	}{
		{name: "nonzero then", condition: "Err.Number <> 0", thenFailure: true, known: true},
		{name: "zero then", condition: "Err.Number = 0", thenFailure: false, known: true},
		{name: "reversed nonzero then", condition: "0 <> Err.Number", thenFailure: true, known: true},
		{name: "composite nonzero then", condition: "Err.Number <> 0 Or itemCount < 0", thenFailure: true, known: true},
		{name: "and compound", condition: "Err.Number <> 0 And itemCount < 0", thenFailure: false, known: false},
		{name: "not comparison", condition: "Not (Err.Number <> 0)", thenFailure: false, known: false},
		{name: "outer comparison", condition: "(Err.Number <> 0) = False", thenFailure: false, known: false},
		{name: "string literal", condition: `value = "Err.Number <> 0"`, thenFailure: false, known: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			thenFailure, known := ErrorNumberThenBranchIsFailure(test.condition)
			if thenFailure != test.thenFailure || known != test.known {
				t.Fatalf("ErrorNumberThenBranchIsFailure(%q) = (%v, %v), want (%v, %v)", test.condition, thenFailure, known, test.thenFailure, test.known)
			}
		})
	}
}

func TestSafeLiteralAssignmentRequiresCompatibleCompleteLiterals(t *testing.T) {
	tests := []struct {
		name       string
		value      string
		targetType string
		want       bool
	}{
		{name: "long zero", value: "0", targetType: "Long", want: true},
		{name: "variant null", value: "Null", targetType: "Variant", want: true},
		{name: "object nothing", value: "Nothing", targetType: "Object", want: true},
		{name: "string literal", value: `"x"`, targetType: "String", want: true},
		{name: "numeric string to long", value: `"1"`, targetType: "Long", want: false},
		{name: "null to long", value: "Null", targetType: "Long", want: false},
		{name: "out of range long", value: "1E100", targetType: "Long", want: false},
		{name: "out of range variant", value: "1E309", targetType: "Variant", want: false},
		{name: "out of range string", value: "1E309", targetType: "String", want: false},
		{name: "out of range boolean", value: "1E309", targetType: "Boolean", want: false},
		{name: "nothing to known object", value: "Nothing", targetType: "Worksheet", want: true},
		{name: "nothing to unknown type", value: "Nothing", targetType: "MyEnum", want: false},
		{name: "nothing to array", value: "Nothing", targetType: "Object()", want: false},
		{name: "rounded byte", value: "255.1", targetType: "Byte", want: true},
		{name: "rounded byte overflow", value: "255.5", targetType: "Byte", want: false},
		{name: "true overflows byte", value: "True", targetType: "Byte", want: false},
		{name: "false fits byte", value: "False", targetType: "Byte", want: true},
		{name: "rounded long", value: "2147483647.1", targetType: "Long", want: true},
		{name: "exact longlong overflow", value: "9223372036854775808", targetType: "LongLong", want: false},
		{name: "longptr max 32 bit", value: "2147483647", targetType: "LongPtr", want: true},
		{name: "longptr 64 bit value", value: "2147483648", targetType: "LongPtr", want: false},
		{name: "string expression", value: `"" & Risky() & ""`, targetType: "String", want: false},
		{name: "nothing to long", value: "Nothing", targetType: "Long", want: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := SafeLiteralAssignment(test.value, test.targetType); got != test.want {
				t.Fatalf("SafeLiteralAssignment(%q, %q) = %v, want %v", test.value, test.targetType, got, test.want)
			}
		})
	}
}
