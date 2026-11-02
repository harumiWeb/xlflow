package intel

import (
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func BenchmarkConditionalProjectionInitial(b *testing.B) {
	for _, size := range []int{4096, 256 << 10, 2 << 20} {
		for _, directives := range []bool{false, true} {
			b.Run(fmt.Sprintf("bytes-%d/directives-%t", size, directives), func(b *testing.B) {
				constants := map[string]string{"Feature": "False"}
				source := strings.Repeat("' padding\n", size/10)
				if directives {
					source += "#If Feature Then\nInactive prose\n#End If\n"
				}
				doc := Document{Path: "Main.bas", Source: source}
				var prepared Document
				b.ReportAllocs()
				for b.Loop() {
					prepared = doc.WithConditionalProjection(constants)
				}
				if prepared.Source != source || prepared.ProjectConditionals != directives {
					b.Fatal("unexpected initial preparation")
				}
			})
		}
	}
}

// BenchmarkConditionalEditorRequests measures the request after immutable
// revision preparation, including the entrypoint's repeated preparation.
// It is intentionally separate from initial parsing and edit invalidation.
func BenchmarkConditionalEditorRequests(b *testing.B) {
	for _, padding := range []int{0, 100000} {
		for _, mode := range []string{"none", "resolved", "unresolved"} {
			source := "Public Function Target(ByVal value As Long) As Long\nTarget = value + 1\nEnd Function\n" +
				"Public Sub Caller()\nCall Target(1)\nEnd Sub\n" + strings.Repeat("' padding\n", padding)
			switch mode {
			case "resolved":
				source += "#If Feature Then\nPrivate Const Selected = 1\n#Else\nPrivate Const Selected = 2\n#End If\n"
			case "unresolved":
				source += "#If Win64 Then\nPrivate Const First = 1\n#Else\nPrivate Const Second = 2\n#End If\n"
			}
			constants := map[string]string{"Feature": "True"}
			cfg := config.Default()
			cfg.VBA.ConditionalConstants = constants
			analyzer := Analyzer{Config: cfg, ConditionalProjection: true}
			for _, request := range []struct {
				name string
				run  func(Document, []Document) error
			}{
				{"hover", func(doc Document, open []Document) error {
					result, err := analyzer.Hover(doc, Position{Line: 4, Character: 7}, open)
					if err == nil && result == nil {
						return fmt.Errorf("missing hover")
					}
					return err
				}},
				{"definition", func(doc Document, open []Document) error {
					result, err := analyzer.Definition(doc, Position{Line: 4, Character: 7}, open, func(path string) string { return path })
					if err == nil && len(result) == 0 {
						return fmt.Errorf("missing definition")
					}
					return err
				}},
				{"completion", func(doc Document, open []Document) error {
					result, err := analyzer.Completions(doc, Position{Line: 4, Character: 8}, open)
					if err == nil && len(result) == 0 {
						return fmt.Errorf("missing completions")
					}
					return err
				}},
				{"signature-help", func(doc Document, open []Document) error {
					result, err := analyzer.SignatureHelp(doc, Position{Line: 4, Character: 12}, open)
					if err == nil && result == nil {
						return fmt.Errorf("missing signature help")
					}
					return err
				}},
			} {
				b.Run(fmt.Sprintf("padding-%d/%s/%s", padding, mode, request.name), func(b *testing.B) {
					snapshot := NewAnalysisSnapshot(Document{Path: "Main.bas", URI: "file:///Main.bas", Source: source}.WithConditionalProjection(constants))
					defer snapshot.RetireAndWait()
					doc := snapshot.Document()
					open := []Document{doc}
					if err := request.run(doc, open); err != nil {
						b.Fatal(err)
					}
					beforeBuilds, _ := snapshot.InteractiveIndexStats()
					beforeParses := snapshot.ParseCount()
					b.ReportAllocs()
					for b.Loop() {
						if err := request.run(doc, open); err != nil {
							b.Fatal(err)
						}
					}
					builds, _ := snapshot.InteractiveIndexStats()
					b.ReportMetric(float64(builds-beforeBuilds)/float64(b.N), "index-builds/op")
					b.ReportMetric(float64(snapshot.ParseCount()-beforeParses)/float64(b.N), "parses/op")
				})
			}
		}
	}
}
