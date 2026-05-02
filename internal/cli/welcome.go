package cli

import (
	"strings"

	"github.com/charmbracelet/lipgloss"

	"github.com/harumiWeb/xlflow/internal/output"
)

var scaffoldWelcomeLogo = []string{
	` ██╗  ██╗ ██╗      ███████╗ ██╗       ██████╗  ██╗    ██╗`,
	` ╚██╗██╔╝ ██║      ██╔════╝ ██║      ██╔═══██╗ ██║    ██║`,
	`  ╚███╔╝  ██║      █████╗   ██║      ██║   ██║ ██║ █╗ ██║`,
	`  ██╔██╗  ██║      ██╔══╝   ██║      ██║   ██║ ██║███╗██║`,
	` ██╔╝ ██╗ ███████╗ ██║      ███████╗ ╚██████╔╝ ╚███╔███╔╝`,
	` ╚═╝  ╚═╝ ╚══════╝ ╚═╝      ╚══════╝  ╚═════╝   ╚══╝╚══╝`,
}

func shouldRenderScaffoldWelcome(command string, opts output.Options) bool {
	if opts.JSON || !opts.Interactive {
		return false
	}
	switch command {
	case "new", "init":
		return true
	default:
		return false
	}
}

func renderScaffoldWelcome(color bool) string {
	badge := renderScaffoldWelcomeBadge("* Welcome to xlflow")
	logo := strings.Join(scaffoldWelcomeLogo, "\n")
	if color {
		style := lipgloss.NewStyle().Foreground(lipgloss.Color("209")).Bold(true)
		badge = style.Render(badge)
		logo = style.Render(logo)
	}
	return badge + "\n\n" + logo + "\n\n"
}

func renderScaffoldWelcomeBadge(text string) string {
	border := strings.Repeat("-", len(text)+2)
	return "+" + border + "+\n| " + text + " |\n+" + border + "+"
}
