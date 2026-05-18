package cli

import (
	"strings"
	"testing"

	"github.com/charmbracelet/lipgloss"
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

func TestRenderScaffoldWelcomeIncludesLogoAndVersion(t *testing.T) {
	got := renderScaffoldWelcome(scaffoldWelcomeModel{Version: "1.2.3"}, false)
	for _, want := range []string{
		"Version: 1.2.3",
		" ██╗  ██╗ ██╗      ███████╗ ██╗       ██████╗  ██╗    ██╗",
		" ╚═╝  ╚═╝ ╚══════╝ ╚═╝      ╚══════╝  ╚═════╝   ╚══╝╚══╝",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("welcome output missing %q:\n%s", want, got)
		}
	}
}

func TestRenderScaffoldWelcomeOmitsWelcomeBadge(t *testing.T) {
	got := renderScaffoldWelcome(scaffoldWelcomeModel{Version: "1.2.3"}, false)
	if strings.Contains(got, "Welcome to xlflow") {
		t.Fatalf("welcome output should not include welcome badge:\n%s", got)
	}
}

func TestRenderScaffoldWelcomePlacesVersionBelowLogo(t *testing.T) {
	got := renderScaffoldWelcome(scaffoldWelcomeModel{Version: "1.2.3"}, false)
	versionIndex := strings.Index(got, "Version: 1.2.3")
	logoIndex := strings.Index(got, " ██╗  ██╗ ██╗      ███████╗ ██╗       ██████╗  ██╗    ██╗")
	if versionIndex < 0 || logoIndex < 0 {
		t.Fatalf("welcome output missing logo or version:\n%s", got)
	}
	if versionIndex < logoIndex {
		t.Fatalf("expected version below logo:\n%s", got)
	}
}

func TestRenderScaffoldWelcomeIncludesUpdateNotice(t *testing.T) {
	got := renderScaffoldWelcome(scaffoldWelcomeModel{
		Version:       "1.2.3",
		UpdateVersion: "v1.2.4",
	}, false)
	for _, want := range []string{
		"Version: 1.2.3",
		"Update available: v1.2.4",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("welcome output missing %q:\n%s", want, got)
		}
	}
}

func TestRenderScaffoldWelcomeBadgeUsesDisplayWidthForEmoji(t *testing.T) {
	got := renderScaffoldWelcomeBadge("🏄‍♂️ Welcome to xlflow")
	lines := strings.Split(got, "\n")
	if len(lines) != 3 {
		t.Fatalf("badge line count = %d, want 3", len(lines))
	}
	for _, line := range lines[1:] {
		if lipgloss.Width(line) != lipgloss.Width(lines[0]) {
			t.Fatalf("badge widths should match:\n%s", got)
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
