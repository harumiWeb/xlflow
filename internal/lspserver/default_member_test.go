package lspserver

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
)

func TestRealtimeFindingPreservesDefaultMemberContext(t *testing.T) {
	got := realtimeFinding(analyze.Finding{
		Code: "VBA253", Severity: "warning",
		Line: 4, Column: 9, EndLine: 4, EndColumn: 14,
		Message: "implicit default member",
		DefaultMember: &analyze.DefaultMemberContext{
			Kind: "recursive", Binding: "known", ExpectedContext: "value",
			Member: "Item", Depth: 2,
		},
	})

	if got.DefaultMember == nil || got.DefaultMember.Kind != "recursive" ||
		got.DefaultMember.Binding != "known" || got.DefaultMember.ExpectedContext != "value" ||
		got.DefaultMember.Member != "Item" || got.DefaultMember.Depth != 2 {
		t.Fatalf("default-member context = %+v", got.DefaultMember)
	}
}
