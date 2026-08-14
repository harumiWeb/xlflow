package constexpr

import "testing"

func TestValuesResolveIsCaseInsensitiveAndDeterministic(t *testing.T) {
	values := Values{
		"LIMIT": {Kind: ValueLong, Integer: 1},
		"limit": {Kind: ValueLong, Integer: 2},
	}
	for i := 0; i < 20; i++ {
		value, ok := values.Resolve("Limit")
		if !ok || value.Integer != 1 {
			t.Fatalf("Resolve(Limit) = %#v, %v; want the lexicographically first spelling", value, ok)
		}
	}
}

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
		{name: "integer division operator", expr: "10 \\ 2", kind: Known, value: 5},
		{name: "mod operator", expr: "10 Mod 3", kind: Known, value: 1},
		{name: "exponent operator", expr: "10 ^ 2", kind: Known, value: 100},
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

func TestEvaluateTypedValuesAndOperators(t *testing.T) {
	values := Values{
		"limit": {Kind: ValueLong, Integer: 3},
		"label": {Kind: ValueString, String: "ok"},
	}
	tests := []struct {
		name string
		expr string
		kind ResultKind
		want Value
	}{
		{name: "string literal", expr: `"a"`, kind: Known, want: Value{Kind: ValueString, String: "a"}},
		{name: "escaped string", expr: `"a""b"`, kind: Known, want: Value{Kind: ValueString, String: `a"b`}},
		{name: "string concat", expr: `"a" & label`, kind: Known, want: Value{Kind: ValueString, String: "aok"}},
		{name: "boolean comparison", expr: "True And Not False", kind: Known, want: Value{Kind: ValueBoolean, Boolean: true}},
		{name: "numeric comparison", expr: "limit >= 3", kind: Known, want: Value{Kind: ValueBoolean, Boolean: true}},
		{name: "float", expr: "1.5! + 2.5#", kind: Known, want: Value{Kind: ValueDouble, Float: 4}},
		{name: "currency", expr: "1.25@ + 0.75@", kind: Known, want: Value{Kind: ValueCurrency, Currency: 20000}},
		{name: "hex", expr: "&H10", kind: Known, want: Value{Kind: ValueLong, Integer: 16}},
		{name: "octal", expr: "&O10", kind: Known, want: Value{Kind: ValueLong, Integer: 8}},
		{name: "runtime call", expr: "Chr$(65)", kind: Unknown},
		{name: "unresolved", expr: "missing + 1", kind: Unknown},
		{name: "null coercion", expr: "Null + 1", kind: Unknown},
		{name: "divide zero", expr: "1 / 0", kind: Invalid},
		{name: "integer overflow", expr: "9223372036854775807 + 1", kind: Unknown},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := Evaluate(test.expr, values)
			if result.Kind != test.kind {
				t.Fatalf("Evaluate(%q) kind = %q, want %q (%#v)", test.expr, result.Kind, test.kind, result)
			}
			if result.Kind == Known && result.Typed != test.want {
				t.Fatalf("Evaluate(%q) value = %#v, want %#v", test.expr, result.Typed, test.want)
			}
		})
	}
}

func TestEvaluateDoesNotExecuteCalls(t *testing.T) {
	values := Values{}
	if result := Evaluate("DangerousFunction()", values); result.Kind != Unknown {
		t.Fatalf("call result = %#v, want Unknown", result)
	}
}

func TestEvaluateNumericBoundaryFixtures(t *testing.T) {
	for _, test := range []struct {
		expr  string
		kind  ResultKind
		value int64
	}{
		{expr: "32767%", kind: Known, value: 32767},
		{expr: "32768%", kind: Unknown},
		{expr: "2147483647&", kind: Known, value: 2147483647},
		{expr: "2147483648&", kind: Unknown},
		{expr: "2147483648", kind: Unknown},
		{expr: "9223372036854775807^", kind: Known, value: 9223372036854775807},
		{expr: "9223372036854775808^", kind: Unknown},
		{expr: "1e309", kind: Unknown},
		{expr: "1e39!", kind: Unknown},
	} {
		t.Run(test.expr, func(t *testing.T) {
			result := Evaluate(test.expr, nil)
			if result.Kind != test.kind {
				t.Fatalf("Evaluate(%q) kind = %q, want %q (%#v)", test.expr, result.Kind, test.kind, result)
			}
			if test.kind == Known && result.Typed.Integer != test.value {
				t.Fatalf("Evaluate(%q) value = %#v, want %d", test.expr, result.Typed, test.value)
			}
		})
	}
}
