package conditional

import (
	"bytes"
	"strings"
	"testing"
)

func TestProjectSelectsOnlyProvenInactiveSource(t *testing.T) {
	for _, tc := range []struct {
		name, source    string
		constants       map[string]string
		present, absent []string
	}{
		{name: "issue966", source: "#If NotDefinedConst Then\nThis is plain prose, not VBA code.\n## Heading\n1. item\n#End If\nPublic Function Answer() As Long\nAnswer = 42\nEnd Function\n", present: []string{"Answer = 42"}, absent: []string{"plain prose", "## Heading", "1. item", "#If", "#End If"}},
		{name: "false", source: "#If False Then\ninactive prose\n#Else\nactive source\n#End If", present: []string{"active source"}, absent: []string{"inactive prose", "#Else"}},
		{name: "shared function body", source: "#If Feature Then\nPublic Function Value() As Long\n#Else\nPublic Function Value() As String\n#End If\nValue = \"x\"\nEnd Function", present: []string{"Public Function Value() As String", "Value = \"x\"", "End Function"}, absent: []string{"As Long", "#If", "#Else", "#End If"}},
		{name: "not undefined", source: "#If Not MissingFlag Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "not comparison precedence", source: "#If Not 1 = 2 Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "true", source: "#If True Then\nactive prose\n#Else\ninactive source\n#End If", present: []string{"active prose"}, absent: []string{"inactive source"}},
		{name: "nested", source: "#If False Then\n#If True Then\ninactive nested\n#End If\n#Else\n#If 1 Then\nactive nested\n#End If\n#End If", present: []string{"active nested"}, absent: []string{"inactive nested"}},
		{name: "elseif", source: "#If False Then\nfirst branch\n#ElseIf 2 > 1 Then\nsecond branch\n#ElseIf True Then\nthird branch\n#Else\nlast branch\n#End If", present: []string{"second branch"}, absent: []string{"first branch", "third branch", "last branch"}},
		{name: "module constant", source: "#Const Debug = 2 + 3\n#If Debug = 5 Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source", "#Const Debug"}, absent: []string{"inactive prose"}},
		{name: "continued const remains uncertain", source: "#Const Flag = False Or _\n    OtherFlag\n#If Flag Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible", "#If Flag"}},
		{name: "custom project argument", constants: map[string]string{"DEBUG": "True"}, source: "#If debug Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "module overrides project", constants: map[string]string{"DEBUG": "True"}, source: "#Const Debug = False\n#If debug Then\ninactive prose\n#Else\nactive source\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "inactive const", source: "#If False Then\n#Const Debug = True\n#End If\n#If Debug Then\ninactive prose\n#End If", absent: []string{"#Const Debug", "inactive prose"}},
		{name: "unknown builtins retain candidates", source: "#If Win64 Then\nfirst branch\n#ElseIf Mac Then\nsecond branch\n#Else\nlast branch\n#End If", present: []string{"first branch", "second branch", "last branch"}},
		{name: "explicit builtin", constants: map[string]string{"Win64": "True"}, source: "#If Win64 Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "unknown assignment merge", source: "#If Win64 Then\n#Const Flag = True\n#End If\n#If Flag Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
		{name: "equal assignment merge", source: "#If Win64 Then\n#Const Flag = True\n#Else\n#Const Flag = True\n#End If\n#If Flag Then\nactive source\n#Else\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "unknown excluded after true", source: "#If True Then\nactive source\n#ElseIf Win64 Then\ninactive prose\n#End If", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "known false nested in uncertain", source: "#If Win64 Then\n#If False Then\ninactive prose\n#End If\n#Else\npossible source\n#End If", present: []string{"possible source"}, absent: []string{"inactive prose"}},
		{name: "string const quote comment remains uncertain", source: "#Const Flag = \"it's \"\"quoted\"\"\" ' comment\n#If Flag = \"it's \"\"quoted\"\"\" Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
		{name: "signed hexadecimal remains uncertain", source: "#If &HFFFF < 0 Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
		{name: "single arithmetic rounding remains uncertain", source: "#If (16777216! + 1!) = 16777216! Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
		{name: "mixed precision comparison remains uncertain", source: "#If 1.00000001# = 1! Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
		{name: "tabs casing endif", source: "#iF\tFalse\tThen ' comment\ninactive prose\n#eLsE\nactive source\n#EnD\tIf", present: []string{"active source"}, absent: []string{"inactive prose"}},
		{name: "unsupported condition retains candidates", source: "#If CBool(1) Then\npossible source\n#Else\nalso possible\n#End If", present: []string{"possible source", "also possible"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := []byte(tc.source)
			before := bytes.Clone(source)
			projected := Project(source, tc.constants)
			if !bytes.Equal(source, before) {
				t.Fatal("input source mutated")
			}
			if len(projected) != len(source) {
				t.Fatalf("byte length changed: %d -> %d", len(source), len(projected))
			}
			for _, text := range tc.present {
				if !bytes.Contains(projected, []byte(text)) {
					t.Errorf("missing %q in %s", text, projected)
				}
			}
			for _, text := range tc.absent {
				if bytes.Contains(projected, []byte(text)) {
					t.Errorf("inactive %q retained in %s", text, projected)
				}
			}
			if again := Project(projected, tc.constants); !bytes.Equal(again, projected) {
				t.Fatalf("projection not idempotent:\n%s\n%s", projected, again)
			}
		})
	}
}

func TestProjectPreservesMalformedStructureAndUnsupportedContinuations(t *testing.T) {
	for _, source := range []string{
		"#If False Then\nprose\n", "#Else\nprose\n", "#End If\n",
		"#If False Then\nprose\n#Else\n#Else\n#End If",
		"#If False Then\nprose\n#ElseIf True Then\n#Else\n#ElseIf True Then\n#End If",
		"#If False Then\nprose\n#End If junk",
		"#If False Then\nprose\n#EndIf junk",
		"#If False\nprose\n#End If",
		"#If False Then\nprose\n#Else junk\n#End If",
		"#Const Bad =\n#If Bad Then\nprose\n#End If",
		"#If False _\n Then\nprose\n#End If",
	} {
		if got := Project([]byte(source), nil); string(got) != source {
			t.Errorf("structure hidden:\n%s\nresult:\n%s", source, got)
		}
	}
}

func TestProjectPreservesPositionsAndModuleScope(t *testing.T) {
	source := []byte("#If False Then\r\n日本語😀 prose\r\n#End If\r\nPublic Sub Run()\r\nEnd Sub\r\n")
	got := Project(source, nil)
	for i, b := range source {
		if b == '\r' || b == '\n' {
			if got[i] != b {
				t.Fatalf("newline at %d moved", i)
			}
		}
	}
	if bytes.Index(got, []byte("Public Sub Run")) != bytes.Index(source, []byte("Public Sub Run")) {
		t.Fatal("active code byte offset moved")
	}
	Project([]byte("#Const Debug = True\n"), nil)
	if got := Project([]byte("#If Debug Then\nprose\n#End If"), nil); bytes.Contains(got, []byte("prose")) {
		t.Fatal("module #Const leaked into next module")
	}
}

func TestProjectDoesNotHideInvalidSourceEncoding(t *testing.T) {
	source := append([]byte("#If False Then\n"), 0xff)
	source = append(source, []byte("\n#End If")...)
	if !bytes.Equal(Project(source, nil), source) {
		t.Fatal("invalid UTF-8 was hidden by projection")
	}
}

func TestValidateConstants(t *testing.T) {
	for _, values := range []map[string]string{
		nil, {"DEBUG": "True", "Win64": "False", "Count": "1 + 2", "Title": `"it's ok"`}, {"EmptyFlag": "Empty"},
	} {
		if err := ValidateConstants(values); err != nil {
			t.Fatalf("valid table %v: %v", values, err)
		}
	}
	for _, values := range []map[string]string{
		{"Debug": "True", "DEBUG": "False"}, {"bad name": "1"}, {"Flag": "Unresolved"}, {"Flag": "1 / 0"}, {"Flag": "Null"}, {"True": "1"}, {"Flag": "Nothing"},
	} {
		if err := ValidateConstants(values); err == nil {
			t.Fatalf("invalid table accepted: %v", values)
		}
	}
}

func BenchmarkProject(b *testing.B) {
	source := []byte(strings.Repeat("Public Sub Run()\nEnd Sub\n", 1000) + "#If False Then\nignored prose\n#End If\n")
	b.ReportAllocs()
	b.SetBytes(int64(len(source)))
	for b.Loop() {
		Project(source, nil)
	}
}
