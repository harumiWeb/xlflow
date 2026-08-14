package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
)

func TestConstantValuesFromSourceResolvesConstAndEnumReferences(t *testing.T) {
	source := `Public Const Limit As Long = 2
Public Const Expanded As Long = Limit + 3
Enum State
    Ready = Expanded
    Done
End Enum
Public Const RuntimeValue As Long = Missing()
#If WINDOWS Then
Public Const ConditionalValue As Long = 99
#End If
`
	values := ConstantValuesFromSource(source, nil, nil)
	for name, want := range map[string]int64{
		"limit":       2,
		"expanded":    5,
		"ready":       5,
		"state.ready": 5,
		"done":        6,
		"state.done":  6,
	} {
		value, ok := values[name]
		if !ok || value.Kind != constexpr.ValueLong || value.Integer != want {
			t.Fatalf("value[%q] = %#v, %v; want Long(%d)", name, value, ok, want)
		}
	}
	if _, ok := values["runtimevalue"]; ok {
		t.Fatal("runtime-dependent Const must not be projected as a known value")
	}
	if _, ok := values["conditionalvalue"]; ok {
		t.Fatal("conditional Const must remain unknown")
	}
}
