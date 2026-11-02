package constexpr

import "testing"

func TestEvaluateConditionalCurrencyPrecision(t *testing.T) {
	for _, tc := range []struct {
		expr string
		want int64
	}{
		{"900719925474.0987@", 9007199254740987},
		{"900719925474.0988@", 9007199254740988},
		{"1@", 10000},
		{"0.0001@", 1},
		{"922337203685477.5807@", 9223372036854775807},
	} {
		got := EvaluateConditional(tc.expr, nil)
		if got.Kind != Known || got.Typed.Kind != ValueCurrency || got.Typed.Currency != tc.want {
			t.Errorf("%s = %+v, want exact Currency %d", tc.expr, got, tc.want)
		}
	}
	for _, expr := range []string{
		"900719925474.0987@ < 900719925474.0988@",
		"900719925474.0987@ <> 900719925474.0988@",
		"900719925474.0988@ > 900719925474.0987@",
		"1@ = 1.0000@",
		"922337203685477.5806@ < 922337203685477.5807@",
		"-900719925474.0988@ < -900719925474.0987@",
	} {
		got := EvaluateConditional(expr, nil)
		if got.Kind != Known || got.Typed.Kind != ValueBoolean || !got.Typed.Boolean {
			t.Errorf("%s = %+v, want True", expr, got)
		}
	}
	for _, expr := range []string{"0.00015@", "1E2@", "922337203685477.5808@", "1@ + 1@", "1@ ^ 2", "1@ = 1#"} {
		if got := EvaluateConditional(expr, nil); got.Kind != Unknown {
			t.Errorf("%s = %+v, want conservative Unknown", expr, got)
		}
	}
}

func TestEvaluateConditionalCoercionsAndPrecedence(t *testing.T) {
	for _, expr := range []string{"(16777216! + 1!) = 16777216!", "2! ^ 3!", "&HFFFF < 0", "1.00000001# = 1!"} {
		if got := EvaluateConditional(expr, nil); got.Kind != Unknown {
			t.Errorf("%s = %+v, want Unknown", expr, got)
		}
	}
	for _, tc := range []struct {
		expr string
		want int64
	}{
		{"Not Empty", -1}, {"Not 0", -1}, {"Not 1", -2},
		{"3 And 1", 1}, {"2 Or 1", 3}, {"3 Xor 1", 2}, {"0 Eqv 0", -1}, {"1 Imp 0", -2},
		{"True + 1", 0}, {"Empty + 1", 1},
	} {
		got := EvaluateConditional(tc.expr, nil)
		if got.Kind != Known || got.Typed.Integer != tc.want {
			t.Errorf("%s = %+v, want %d", tc.expr, got, tc.want)
		}
	}
	for _, expr := range []string{"Not 1 = 2", "Not 1 = 1 Or True", "Empty = 0", "True = -1", "(3 And 1) = 1", "Not False And True"} {
		got := EvaluateConditional(expr, nil)
		if got.Kind != Known || got.Typed.Kind != ValueBoolean || !got.Typed.Boolean {
			t.Errorf("%s = %+v, want True", expr, got)
		}
	}
	for _, expr := range []string{"Not 1 = 1", "Not 2 > 1", "True = 1"} {
		got := EvaluateConditional(expr, nil)
		if got.Kind != Known || got.Typed.Kind != ValueBoolean || got.Typed.Boolean {
			t.Errorf("%s = %+v, want False", expr, got)
		}
	}
	if got := EvaluateConditional("9007199254740992^ = 9007199254740993^", nil); got.Kind == Known {
		t.Errorf("wide integer comparison = %+v, want conservative Unknown", got)
	}
	if got := EvaluateConditional("2 ^ True", nil); got.Kind != Known || got.Typed.Float != 0.5 {
		t.Errorf("Boolean exponent = %+v", got)
	}
	for _, expr := range []string{"Not 0", "3 And 1", "True = -1", "Empty = 0"} {
		if got := Evaluate(expr, nil); got.Kind == Known {
			t.Errorf("ordinary evaluator unexpectedly widened for %s: %+v", expr, got)
		}
	}
}
