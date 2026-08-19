package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestProcedureStartPreservesPropertyAccessor(t *testing.T) {
	tests := []struct {
		name     string
		line     string
		kind     string
		accessor string
		nameWant string
	}{
		{name: "sub", line: "Public Sub Run()", kind: "Sub", nameWant: "Run"},
		{name: "function", line: "Private Function Value() As String", kind: "Function", nameWant: "Value"},
		{name: "property get", line: "Friend Property Get Caption() As String", kind: "Property", accessor: "Get", nameWant: "Caption"},
		{name: "property let", line: "Public Property Let Caption(ByVal value As String)", kind: "Property", accessor: "Let", nameWant: "Caption"},
		{name: "property set", line: "Public Property Set Item(ByVal value As Object)", kind: "Property", accessor: "Set", nameWant: "Item"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			frame, ok := procedureStart(tt.line, 3)
			if !ok {
				t.Fatalf("procedureStart(%q) did not recognize a procedure", tt.line)
			}
			if frame.Kind != tt.kind || frame.Accessor != tt.accessor || frame.Name != tt.nameWant {
				t.Fatalf("procedureStart(%q) = %+v, want kind=%q accessor=%q name=%q", tt.line, frame, tt.kind, tt.accessor, tt.nameWant)
			}
		})
	}
}

func TestProcedureBoundaryClassifierReceivesVBEAuditContext(t *testing.T) {
	var got []ProcedureBoundaryContext
	tracker := newProcedureBoundaryTracker(Linter{
		ModuleKind: "class",
		ProcedureBoundaryClassifier: func(context ProcedureBoundaryContext) ProcedureBoundaryDecision {
			got = append(got, context)
			return ProcedureBoundaryDecision{Outcome: ProcedureBoundaryAccepted}
		},
	}, "Widget.cls")
	tracker.processStatement(7, "Property Get Caption() As String")
	tracker.processStatement(9, "End Function")
	if issues := tracker.finish(); len(issues) != 0 {
		t.Fatalf("accepted mismatch still emitted boundary issues: %+v", issues)
	}
	if len(got) != 1 {
		t.Fatalf("classifier calls = %d, want 1", len(got))
	}
	want := ProcedureBoundaryContext{
		ModuleKind:     "class",
		OpenerKind:     "Property",
		Accessor:       "Get",
		TerminatorKind: "Function",
		ProcedureName:  "Caption",
		OpeningLine:    7,
		ClosingLine:    9,
	}
	if got[0] != want {
		t.Fatalf("classifier context = %+v, want %+v", got[0], want)
	}
}

func TestProcedureBoundaryClassifierSeparatesAcceptedStyleAndCompileInvalid(t *testing.T) {
	tests := []struct {
		name       string
		decision   ProcedureBoundaryDecision
		wantCode   string
		wantSev    string
		wantIssues int
	}{
		{
			name:       "accepted",
			decision:   ProcedureBoundaryDecision{Outcome: ProcedureBoundaryAccepted},
			wantIssues: 0,
		},
		{
			name:       "style without promoted code",
			decision:   ProcedureBoundaryDecision{Outcome: ProcedureBoundaryStyleMismatch},
			wantIssues: 0,
		},
		{
			name: "style custom code",
			decision: ProcedureBoundaryDecision{
				Outcome:  ProcedureBoundaryStyleMismatch,
				Code:     "TEST066",
				Severity: "information",
			},
			wantCode:   "TEST066",
			wantSev:    "information",
			wantIssues: 1,
		},
		{
			name:       "compile invalid",
			decision:   ProcedureBoundaryDecision{Outcome: ProcedureBoundaryCompileInvalid},
			wantCode:   "VB012",
			wantSev:    "error",
			wantIssues: 1,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tracker := newProcedureBoundaryTracker(Linter{
				ProcedureBoundaryClassifier: func(ProcedureBoundaryContext) ProcedureBoundaryDecision {
					return tt.decision
				},
			}, "Main.bas")
			tracker.processStatement(2, "Sub Run()")
			tracker.processStatement(4, "End Function")
			issues := tracker.finish()
			if len(issues) != tt.wantIssues {
				t.Fatalf("issues = %+v, want %d issue(s)", issues, tt.wantIssues)
			}
			if tt.wantIssues == 0 {
				return
			}
			if issues[0].Code != tt.wantCode || issues[0].Severity != tt.wantSev {
				t.Fatalf("issue = %+v, want code=%q severity=%q", issues[0], tt.wantCode, tt.wantSev)
			}
		})
	}
}

func TestProcedureBoundaryClassifierNilUsesVB012ForRejectedCombinations(t *testing.T) {
	tracker := newProcedureBoundaryTracker(Linter{}, "Main.bas")
	tracker.processStatement(2, "Sub Run()")
	tracker.processStatement(4, "End Function")
	issues := tracker.finish()
	if len(issues) != 1 || issues[0].Code != "VB012" || issues[0].Severity != "error" {
		t.Fatalf("default mismatch classification = %+v, want VB012/error", issues)
	}
}

func TestDefaultProcedureBoundaryClassifierMatrix(t *testing.T) {
	moduleKinds := []string{"standard", "class", "document-workbook", "document-worksheet"}
	tests := []struct {
		name       string
		opener     string
		accessor   string
		terminator string
		wantCode   string
		wantSev    string
	}{
		{name: "sub end function", opener: "Sub", terminator: "Function", wantCode: "VB012", wantSev: "error"},
		{name: "sub end property", opener: "Sub", terminator: "Property", wantCode: "VB012", wantSev: "error"},
		{name: "function end sub", opener: "Function", terminator: "Sub", wantCode: "VB012", wantSev: "error"},
		{name: "function end property", opener: "Function", terminator: "Property", wantCode: "VB012", wantSev: "error"},
		{name: "property get end sub", opener: "Property", accessor: "Get", terminator: "Sub", wantCode: "VB066", wantSev: "warning"},
		{name: "property get end function", opener: "Property", accessor: "Get", terminator: "Function", wantCode: "VB066", wantSev: "warning"},
		{name: "property let end sub", opener: "Property", accessor: "Let", terminator: "Sub", wantCode: "VB012", wantSev: "error"},
		{name: "property let end function", opener: "Property", accessor: "Let", terminator: "Function", wantCode: "VB012", wantSev: "error"},
		{name: "property set end sub", opener: "Property", accessor: "Set", terminator: "Sub", wantCode: "VB012", wantSev: "error"},
		{name: "property set end function", opener: "Property", accessor: "Set", terminator: "Function", wantCode: "VB012", wantSev: "error"},
	}
	for _, moduleKind := range moduleKinds {
		for _, tt := range tests {
			t.Run(moduleKind+"/"+tt.name, func(t *testing.T) {
				tracker := newProcedureBoundaryTracker(Linter{ModuleKind: moduleKind}, "Probe.bas")
				declaration := tt.opener + " Probe()"
				if tt.opener == "Property" {
					declaration = "Property " + tt.accessor + " Probe()"
				}
				tracker.processStatement(2, declaration)
				tracker.processStatement(4, "End "+tt.terminator)
				issues := tracker.finish()
				if len(issues) != 1 || issues[0].Code != tt.wantCode || issues[0].Severity != tt.wantSev {
					t.Fatalf("issues = %+v, want %s/%s", issues, tt.wantCode, tt.wantSev)
				}
				if tt.wantCode == "VB066" && len(PushBlockingIssues(issues)) != 0 {
					t.Fatalf("style mismatch became preflight-blocking: %+v", issues)
				}
			})
		}
	}
}

func TestDefaultProcedureBoundaryClassifierLeavesUnauditedUserFormsCompileInvalid(t *testing.T) {
	decision := DefaultProcedureBoundaryClassifier(ProcedureBoundaryContext{
		ModuleKind:     "form",
		OpenerKind:     "Property",
		Accessor:       "Get",
		TerminatorKind: "Sub",
	})
	if decision.Outcome != ProcedureBoundaryCompileInvalid {
		t.Fatalf("UserForm compatibility decision = %+v, want compile-invalid outside the audited scope", decision)
	}
}

func TestProcedureBoundaryCompatibilityFullLintAndPreflight(t *testing.T) {
	tests := []struct {
		name       string
		moduleKind string
		source     string
		wantCode   string
		wantSev    string
		blocking   bool
	}{
		{name: "accepted property get end sub", moduleKind: "standard", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Get Value() As Long\nEnd Sub\n", wantCode: "VB066", wantSev: "warning"},
		{name: "accepted property get end function document", moduleKind: "document", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Get Value() As Long\nEnd Function\n", wantCode: "VB066", wantSev: "warning"},
		{name: "accepted style warning can be suppressed", moduleKind: "standard", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Get Value() As Long\nEnd Sub ' xlflow:disable-line VB066\n"},
		{name: "rejected function end sub", moduleKind: "class", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Function Value() As Long\nEnd Sub\n", wantCode: "VB012", wantSev: "error", blocking: true},
		{name: "rejected property let end function", moduleKind: "document", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Let Value(ByVal value As Long)\nEnd Function\n", wantCode: "VB012", wantSev: "error", blocking: true},
		{name: "matching property set", moduleKind: "document", source: "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Set Value(ByVal value As Object)\nEnd Property\n"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default(), ModuleKind: tt.moduleKind}).LintSource("Probe.bas", []byte(tt.source))
			if err != nil {
				t.Fatal(err)
			}
			got := issuesByCode(issues, tt.wantCode)
			if tt.wantCode == "" {
				if got := issuesByCode(issues, "VB012"); len(got) != 0 {
					t.Fatalf("matching source emitted VB012: %+v", got)
				}
				if got := issuesByCode(issues, "VB066"); len(got) != 0 {
					t.Fatalf("matching source emitted VB066: %+v", got)
				}
				return
			}
			if len(got) != 1 || got[0].Severity != tt.wantSev {
				t.Fatalf("%s = %+v, want one %s/%s; all issues=%+v", tt.wantCode, got, tt.wantCode, tt.wantSev, issues)
			}
			if hasVB014 := len(issuesByCode(issues, "VB014")); hasVB014 != 0 && tt.wantCode == "VB066" {
				t.Fatalf("accepted VBE source retained VB014 recovery: %+v", issues)
			}
			blocking := len(PushBlockingIssues(issues)) > 0
			if blocking != tt.blocking {
				t.Fatalf("blocking = %v, want %v; issues=%+v", blocking, tt.blocking, issues)
			}
		})
	}
}
