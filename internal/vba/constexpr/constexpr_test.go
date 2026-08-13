package constexpr

import "testing"

func TestEvaluateIntegerClassifiesKnownUnknownAndInvalid(t *testing.T) {
	constants := map[string]int{"limit": 3, "offset": -1}
	for _, test := range []struct {
		name  string
		expr  string
		kind  ResultKind
		value int
	}{
		{name: "literal arithmetic", expr: "(Limit + 2) * 2", kind: Known, value: 10},
		{name: "unary", expr: "-Offset", kind: Known, value: 1},
		{name: "unresolved", expr: "runtimeValue", kind: Unknown},
		{name: "call", expr: "GetBound()", kind: Unknown},
		{name: "trailing call", expr: "Limit + GetBound()", kind: Unknown},
		{name: "integer division operator", expr: "10 \\ 2", kind: Unknown},
		{name: "mod operator", expr: "10 Mod 3", kind: Unknown},
		{name: "exponent operator", expr: "10 ^ 2", kind: Unknown},
		{name: "non-integral division", expr: "10 / 3", kind: Unknown},
		{name: "divide zero", expr: "4 / 0", kind: Invalid},
		{name: "syntax", expr: "1 To 2", kind: Invalid},
	} {
		t.Run(test.name, func(t *testing.T) {
			result := EvaluateInteger(test.expr, constants)
			if result.Kind != test.kind || result.Kind == Known && result.Value != test.value {
				t.Fatalf("EvaluateInteger(%q) = %#v, want kind=%q value=%d", test.expr, result, test.kind, test.value)
			}
		})
	}
}
