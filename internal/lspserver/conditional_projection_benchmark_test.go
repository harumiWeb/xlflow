package lspserver

import (
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// BenchmarkLSPConditionalEdits keeps ordinary body and branch-selection edits
// distinct. Each sample includes document publication and the next query.
func BenchmarkLSPConditionalEdits(b *testing.B) {
	for _, padding := range []int{0, 100000} {
		for _, conditional := range []bool{false, true} {
			b.Run(fmt.Sprintf("padding-%d/conditional-%t", padding, conditional), func(b *testing.B) {
				fixture, source := makeIssue756OrdinaryFixture(b)
				fixture.largeSource = "#Const Feature = False\n" + source + strings.Repeat("' padding\n", padding) +
					"#If Feature Then\nPrivate Const First = 1\n#Else\nPrivate Const Second = 2\n#End If\n"
				s := newLSPBenchmarkServer(b, fixture)
				line := benchmarkSourceLine(fixture.largeSource, "    Call ProjectTarget(1)")
				column := len("    Call ProjectTarget(")
				position := intel.Position{Line: line, Character: 12}
				var version int32 = 1
				b.ReportAllocs()
				for b.Loop() {
					text := "2"
					editLine, editColumn, length := line, column, 1
					if conditional {
						text = "True "
						editLine, editColumn, length = 0, len("#Const Feature = "), 5
					}
					if version%2 == 0 {
						if conditional {
							text = "False"
						} else {
							text = "1"
						}
					}
					version++
					changed, applied, err := s.docs.applyChanges(fixture.largeURI, []documentContentChange{{
						rng: protocolRange(editLine, editColumn, editLine, editColumn+length), text: text,
					}}, version)
					if err != nil || !applied {
						b.Fatalf("edit applied=%t, err=%v", applied, err)
					}
					result, err := s.analyzer.Hover(changed, position, []intel.Document{changed})
					if err != nil || result == nil {
						b.Fatalf("hover=%v, err=%v", result, err)
					}
				}
			})
		}
	}
}
