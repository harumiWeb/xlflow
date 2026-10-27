package symbols

import "testing"

func TestIsPredeclaredClass(t *testing.T) {
	t.Parallel()
	for _, tc := range []struct {
		name, kind, source string
		want               bool
	}{
		{name: "true", kind: "class", source: "Attribute VB_PredeclaredId = True\n", want: true},
		{name: "case insensitive", kind: "class", source: "Attribute vb_predeclaredid = true\n", want: true},
		{name: "false", kind: "class", source: "Attribute VB_PredeclaredId = False\n"},
		{name: "omitted", kind: "class"},
		{name: "comment", kind: "class", source: "' Attribute VB_PredeclaredId = True\n"},
		{name: "not a class", kind: "standard", source: "Attribute VB_PredeclaredId = True\n"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			file, err := InspectSourceContext(t.Context(), SourceOptions{Path: "Widget.bas", ModuleKind: tc.kind}, []byte("Attribute VB_Name = \"Widget\"\n"+tc.source+"Option Explicit\n"))
			if err != nil {
				t.Fatal(err)
			}
			if len(file.Symbols) != 1 {
				t.Fatalf("symbols = %+v, want module only", file.Symbols)
			}
			if got := IsPredeclaredClass(file.Symbols[0]); got != tc.want {
				t.Fatalf("IsPredeclaredClass(%+v) = %v, want %v", file.Symbols[0], got, tc.want)
			}
		})
	}
}
