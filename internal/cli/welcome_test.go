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
		" ██╗  ██╗ ██╗      ███████╗ ██╗       ██████╗  ██╗    ██╗",
		" ╚═╝  ╚═╝ ╚══════╝ ╚═╝      ╚══════╝  ╚═════╝   ╚══╝╚══╝",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("welcome output missing %q:\n%s", want, got)
		}
	}
}

func TestInterpolateColorPreservesGradientEndpoints(t *testing.T) {
	if got := interpolateColor(welcomeTitleStart, welcomeTitleEnd, 0, 10); got != welcomeTitleStart {
		t.Fatalf("gradient start = %#v, want %#v", got, welcomeTitleStart)
	}
	if got := interpolateColor(welcomeTitleStart, welcomeTitleEnd, 9, 10); got != welcomeTitleEnd {
		t.Fatalf("gradient end = %#v, want %#v", got, welcomeTitleEnd)
	}
}

func TestRenderGradientBlockKeepsOriginalText(t *testing.T) {
	got := renderGradientBlock([]string{"XL"}, welcomeTitleStart, welcomeTitleEnd)
	for _, want := range []string{"X", "L"} {
		if !strings.Contains(got, want) {
			t.Fatalf("gradient output missing %q:\n%s", want, got)
		}
	}
}

func TestGradientColorsForLineUsesRunePositionsForUnicodeArt(t *testing.T) {
	got := gradientColorsForLine("██", welcomeTitleStart, welcomeTitleEnd)
	if len(got) != 2 {
		t.Fatalf("gradient color count = %d, want 2", len(got))
	}
	if got[0] != welcomeTitleStart {
		t.Fatalf("first unicode gradient color = %#v, want %#v", got[0], welcomeTitleStart)
	}
	if got[1] != welcomeTitleEnd {
		t.Fatalf("last unicode gradient color = %#v, want %#v", got[1], welcomeTitleEnd)
	}
}
