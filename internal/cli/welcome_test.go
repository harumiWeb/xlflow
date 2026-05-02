package cli

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/output"
)

func TestShouldRenderScaffoldWelcome(t *testing.T) {
	tests := []struct {
		name    string
		command string
		opts    output.Options
		want    bool
	}{
		{
			name:    "interactive init",
			command: "init",
			opts:    output.Options{Interactive: true},
			want:    true,
		},
		{
			name:    "interactive new",
			command: "new",
			opts:    output.Options{Interactive: true},
			want:    true,
		},
		{
			name:    "json output",
			command: "init",
			opts:    output.Options{Interactive: true, JSON: true},
			want:    false,
		},
		{
			name:    "non interactive",
			command: "init",
			opts:    output.Options{},
			want:    false,
		},
		{
			name:    "other command",
			command: "run",
			opts:    output.Options{Interactive: true},
			want:    false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := shouldRenderScaffoldWelcome(tt.command, tt.opts); got != tt.want {
				t.Fatalf("shouldRenderScaffoldWelcome(%q, %+v) = %v, want %v", tt.command, tt.opts, got, tt.want)
			}
		})
	}
}

func TestRenderScaffoldWelcomeIncludesBadgeAndLogo(t *testing.T) {
	got := renderScaffoldWelcome(false)
	for _, want := range []string{
		"* Welcome to xlflow",
		"XX   XX LL      FFFFFF LL       OOOOO  WW      WW",
		"XXXXXXX LL      FFFF   LL      OO   OO WW  WW  WW",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("welcome output missing %q:\n%s", want, got)
		}
	}
}
