package filepush

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// The push-parity corpus is shared with the .NET bridge test project
// (testdata/push-parity/cases.json). Both sides normalize to LF before
// comparing because the bridge joins lines with Environment.NewLine.
type pushParityFixture struct {
	AnnotationCases []struct {
		ID         string `json:"id"`
		Mode       string `json:"mode"`
		Annotation string `json:"annotation"`
		Source     string `json:"source"`
		Expected   string `json:"expected"`
	} `json:"annotationCases"`
	LineNumberAddCases []struct {
		ID        string `json:"id"`
		Source    string `json:"source"`
		Expected  string `json:"expected"`
		Unsafe    bool   `json:"unsafe"`
		IssueLine int    `json:"issueLine"`
	} `json:"lineNumberAddCases"`
	DocumentCases []struct {
		ID       string `json:"id"`
		Source   string `json:"source"`
		Expected string `json:"expected"`
	} `json:"documentCases"`
}

func TestPushParityFixtures(t *testing.T) {
	body, err := os.ReadFile(filepath.Join("..", "..", "testdata", "push-parity", "cases.json"))
	if err != nil {
		t.Fatal(err)
	}
	var fixture pushParityFixture
	if err := json.Unmarshal(body, &fixture); err != nil {
		t.Fatal(err)
	}
	normalize := func(text string) string {
		return strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n")
	}
	for _, test := range fixture.AnnotationCases {
		t.Run("annotation/"+test.ID, func(t *testing.T) {
			if got := normalize(updateFolderAnnotationText(test.Source, test.Mode, test.Annotation)); got != test.Expected {
				t.Fatalf("got %q, want %q", got, test.Expected)
			}
		})
	}
	for _, test := range fixture.LineNumberAddCases {
		t.Run("line-numbers/"+test.ID, func(t *testing.T) {
			got, issue := tryAddLineNumbers(test.Source)
			if test.Unsafe {
				if issue == nil {
					t.Fatalf("unsafe input was instrumented: %q", got)
				}
				if test.IssueLine != 0 && issue.Line != test.IssueLine {
					t.Fatalf("issue line = %d, want %d", issue.Line, test.IssueLine)
				}
				return
			}
			if issue != nil {
				t.Fatalf("unexpected issue: %+v", issue)
			}
			if normalize(got) != test.Expected {
				t.Fatalf("got %q, want %q", normalize(got), test.Expected)
			}
		})
	}
	for _, test := range fixture.DocumentCases {
		t.Run("document/"+test.ID, func(t *testing.T) {
			if got := normalize(normalizeDocumentModuleContent(test.Source)); got != test.Expected {
				t.Fatalf("got %q, want %q", got, test.Expected)
			}
		})
	}
}
