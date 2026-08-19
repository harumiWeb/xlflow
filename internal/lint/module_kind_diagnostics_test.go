package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestModuleKindDiagnosticsMatrix(t *testing.T) {
	cases := []struct {
		name string
		kind string
		src  string
		code string
		want int
	}{
		{"event standard", "standard", "Public Event Changed()\n", "VB050", 1},
		{"event class", "class", "Public Event Changed()\n", "VB050", 0},
		{"friend standard", "standard", "Friend Sub Run()\nEnd Sub\n", "VB050", 1},
		{"friend document", "document", "Friend Sub Run()\nEnd Sub\n", "VB050", 0},
		{"friend class", "class", "Friend Sub Run()\nEnd Sub\n", "VB050", 0},
		{"implements standard", "standard", "Implements IFoo\n", "VB050", 1},
		{"implements class", "class", "Implements IFoo\n", "VB050", 0},
		{"withevents standard", "standard", "Private WithEvents App As Excel.Application\n", "VB050", 1},
		{"withevents object", "class", "Private WithEvents App As Object\n", "VB050", 1},
		{"withevents scalar", "class", "Private WithEvents App As Long\n", "VB050", 1},
		{"withevents array", "class", "Private WithEvents App() As Excel.Application\n", "VB050", 1},
		{"withevents as new", "class", "Private WithEvents App As New Excel.Application\n", "VB050", 1},
		{"withevents as new spaced", "class", "Private WithEvents App As  New Excel.Application\n", "VB050", 1},
		{"withevents qualified", "class", "Private WithEvents App As Excel.Application\n", "VB050", 0},
		{"withevents unresolved", "class", "Private WithEvents App As ExternalSink\n", "VB050", 0},
		{"withevents procedure local", "class", "Sub Run()\nPrivate WithEvents App As Excel.Application\nEnd Sub\n", "VB050", 1},
		{"public const object module", "class", "Public Const Flag = True\n", "VB050", 1},
		{"public const form module", "form", "Public Const Flag = True\n", "VB050", 1},
		{"public array object module", "class", "Public Values() As Long\n", "VB050", 1},
		{"public fixed string object module", "class", "Public Name As String * 10\n", "VB050", 1},
		{"public type object module", "class", "Public Type State\n Value As Long\nEnd Type\n", "VB050", 1},
		{"public declare object module", "class", "Public Declare Function GetValue Lib \"x\" () As Long\n", "VB050", 1},
		{"public scalar object module", "class", "Public Value As Long\n", "VB050", 0},
		{"me standard", "standard", "Sub Run()\n Me\nEnd Sub\n", "VB051", 1},
		{"me class", "class", "Sub Run()\n Me\nEnd Sub\n", "VB051", 0},
	}
	objectTypesByCase := map[string]map[string]int{
		// Keep the qualified accepted case on the resolved-object path rather
		// than letting it pass only because unknown types are fail-open.
		"withevents qualified": {"excel.application": 1},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{
				Config:                 config.Default(),
				ModuleKind:             tc.kind,
				ObjectTypeDeclarations: objectTypesByCase[tc.name],
			}).LintSource("Module.cls", []byte(tc.src))
			if err != nil {
				t.Fatal(err)
			}
			got := 0
			for _, issue := range issues {
				if issue.Code == tc.code {
					got++
					if issue.Kind == "" {
						t.Fatalf("%s issue has no kind: %+v", tc.code, issue)
					}
				}
			}
			if got != tc.want {
				t.Fatalf("%s diagnostics = %d, want %d: %+v", tc.code, got, tc.want, issues)
			}
		})
	}
}

func TestModuleKindDiagnosticsProcedureLocalWithEventsKind(t *testing.T) {
	source := "Sub Run()\nPrivate WithEvents App As Excel.Application\nEnd Sub\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "class"}).LintSource("Module.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB050" {
			if issue.Kind != "invalid_withevents_module" {
				t.Fatalf("WithEvents issue kind = %q, want invalid_withevents_module: %+v", issue.Kind, issue)
			}
			return
		}
	}
	t.Fatalf("missing procedure-local WithEvents diagnostic: %+v", issues)
}

func TestModuleKindDiagnosticsUseTightKeywordAndMeRanges(t *testing.T) {
	checks := []struct {
		name string
		kind string
		src  string
		code string
		line int
		col  int
	}{
		{name: "event name", kind: "standard", src: "Public Event Changed()\n", code: "VB050", line: 1, col: 14},
		{name: "me token", kind: "standard", src: "Sub Run()\n Me\nEnd Sub\n", code: "VB051", line: 2, col: 2},
	}
	for _, tc := range checks {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default(), ModuleKind: tc.kind}).LintSource("Module.bas", []byte(tc.src))
			if err != nil {
				t.Fatal(err)
			}
			var found *Issue
			for i := range issues {
				if issues[i].Code == tc.code {
					found = &issues[i]
					break
				}
			}
			if found == nil {
				t.Fatalf("missing %s diagnostic: %+v", tc.code, issues)
			}
			if found.Line != tc.line || found.Column != tc.col {
				t.Fatalf("%s range = %d:%d, want %d:%d", tc.code, found.Line, found.Column, tc.line, tc.col)
			}
		})
	}
}

func TestModuleKindDiagnosticsRespectConstantConditionalBranches(t *testing.T) {
	source := "#If False Then\nPublic Event Ignored()\n#Else\nSub Run()\nEnd Sub\n#End If\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "standard"}).LintSource("Module.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB050" {
			t.Fatalf("inactive module declaration produced VB050: %+v", issue)
		}
	}
}

func TestModuleKindDiagnosticsSkipKnownFalseElseIfBranches(t *testing.T) {
	source := "#If False Then\nPublic Event Ignored()\n#ElseIf False Then\nPublic Event AlsoIgnored()\n#Else\nSub Run()\nEnd Sub\n#End If\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "standard"}).LintSource("Module.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB050" {
			t.Fatalf("known-false conditional branch produced VB050: %+v", issue)
		}
	}
}

func TestModuleKindDiagnosticsIgnoreCommentsAndStringsForMe(t *testing.T) {
	source := "Sub Run()\n  ' Me in a comment\n  Debug.Print \"Me\"\n  obj.Me\nEnd Sub\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "standard"}).LintSource("Module.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB051" {
			t.Fatalf("comment/string Me produced VB051: %+v", issue)
		}
	}
}
