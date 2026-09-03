package procedureir

import "testing"

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
