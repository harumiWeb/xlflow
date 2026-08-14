package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

// These source-shape cases currently have only parser-recovery evidence. Keep
// the generic VB014 fallback until a host-backed rule can distinguish syntax
// categories without guessing from tree-sitter recovery nodes.
func TestLinterMalformedExpressionAndHeaderRemainVB014Fallback(t *testing.T) {
	t.Parallel()
	cases := []struct {
		name        string
		source      string
		parserNode  string
		parserToken string
		line        int
	}{
		{
			name:        "unbalanced parenthesis",
			source:      "Sub Main()\n  value = (left + right\nEnd Sub\n",
			parserNode:  "MISSING",
			parserToken: ")",
			line:        2,
		},
		{
			name:        "extra closing parenthesis",
			source:      "Sub Main()\n  value = left + right)\nEnd Sub\n",
			parserNode:  "ERROR",
			parserToken: ")",
			line:        2,
		},
		{
			name:        "operator without right operand",
			source:      "Sub Main()\n  value = left +\nEnd Sub\n",
			parserNode:  "ERROR",
			parserToken: "+",
			line:        2,
		},
		{
			name:        "procedure header missing close",
			source:      "Sub Main(\n  value = 1\nEnd Sub\n",
			parserNode:  "ERROR",
			parserToken: "(",
			line:        1,
		},
		{
			name:        "procedure header missing name",
			source:      "Sub ()\nEnd Sub\n",
			parserNode:  "ERROR",
			parserToken: "Sub ()",
			line:        1,
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(tc.source))
			if err != nil {
				t.Fatal(err)
			}
			got := issuesByCode(issues, "VB014")
			if len(got) != 1 {
				t.Fatalf("VB014 = %+v, want one generic recovery issue", got)
			}
			issue := got[0]
			if issue.Kind != "parser_recovery" || issue.ParserNode != tc.parserNode || issue.ParserToken != tc.parserToken || issue.Line != tc.line {
				t.Fatalf("VB014 = %+v, want node=%q token=%q line=%d", issue, tc.parserNode, tc.parserToken, tc.line)
			}
			if issue.Message != parserRecoveryMessage {
				t.Fatalf("VB014 message = %q, want neutral parser-recovery guidance", issue.Message)
			}
		})
	}
}

func TestLinterTopLevelExecutableStatementRemainsVB014Fallback(t *testing.T) {
	t.Parallel()
	source := "Debug.Print 1\nSub Main()\nEnd Sub\n"
	issues, err := (Linter{Config: config.Default()}).LintSource("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB014")
	if len(got) != 1 || got[0].Kind != "parser_recovery" || got[0].Line != 1 {
		t.Fatalf("VB014 = %+v, want generic recovery at top-level statement", got)
	}
}
