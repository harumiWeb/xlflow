package cli

import (
	"fmt"
	"strings"
	"unicode/utf8"

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

type rgbColor struct {
	r int
	g int
	b int
}

var (
	welcomeTitleStart = rgbColor{r: 143, g: 211, b: 255}
	welcomeTitleEnd   = rgbColor{r: 184, g: 245, b: 162}
	welcomeInfoColor  = rgbColor{r: 208, g: 214, b: 220}
	welcomeAlertColor = rgbColor{r: 255, g: 165, b: 0}
)

type scaffoldWelcomeModel struct {
	Version       string
	UpdateVersion string
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

func renderScaffoldWelcome(model scaffoldWelcomeModel, color bool) string {
	logo := strings.Join(scaffoldWelcomeLogo, "\n")
	meta := renderScaffoldWelcomeMeta(model, color)
	if color {
		logo = renderGradientBlock(scaffoldWelcomeLogo, welcomeTitleStart, welcomeTitleEnd)
	}
	if meta == "" {
		return logo + "\n\n"
	}
	return logo + "\n\n" + meta + "\n\n"
}

func renderScaffoldWelcomeBadge(text string) string {
	border := strings.Repeat("-", lipgloss.Width(text)+2)
	return "+" + border + "+\n| " + text + " |\n+" + border + "+"
}

func renderScaffoldWelcomeMeta(model scaffoldWelcomeModel, color bool) string {
	lines := make([]string, 0, 2)
	if version := strings.TrimSpace(model.Version); version != "" {
		line := "Version: " + version
		if color {
			line = lipgloss.NewStyle().
				Foreground(lipgloss.Color(welcomeInfoColor.hex())).
				Render(line)
		}
		lines = append(lines, line)
	}
	if updateVersion := strings.TrimSpace(model.UpdateVersion); updateVersion != "" {
		line := "Update available: " + updateVersion
		if color {
			line = lipgloss.NewStyle().
				Foreground(lipgloss.Color(welcomeAlertColor.hex())).
				Bold(true).
				Render(line)
		}
		lines = append(lines, line)
	}
	return strings.Join(lines, "\n")
}

func renderGradientBlock(lines []string, start, end rgbColor) string {
	rendered := make([]string, 0, len(lines))
	for _, line := range lines {
		colors := gradientColorsForLine(line, start, end)
		var b strings.Builder
		runeIndex := 0
		for _, r := range line {
			style := lipgloss.NewStyle().
				Foreground(lipgloss.Color(colors[runeIndex].hex())).
				Bold(true)
			b.WriteString(style.Render(string(r)))
			runeIndex++
		}
		rendered = append(rendered, b.String())
	}
	return strings.Join(rendered, "\n")
}

func gradientColorsForLine(line string, start, end rgbColor) []rgbColor {
	width := utf8.RuneCountInString(line)
	colors := make([]rgbColor, 0, width)
	for i := 0; i < width; i++ {
		colors = append(colors, interpolateColor(start, end, i, width))
	}
	return colors
}

func interpolateColor(start, end rgbColor, index, width int) rgbColor {
	if width <= 1 {
		return start
	}
	return rgbColor{
		r: interpolateChannel(start.r, end.r, index, width-1),
		g: interpolateChannel(start.g, end.g, index, width-1),
		b: interpolateChannel(start.b, end.b, index, width-1),
	}
}

func interpolateChannel(start, end, index, steps int) int {
	if steps <= 0 {
		return start
	}
	return start + ((end-start)*index)/steps
}

func (c rgbColor) hex() string {
	return fmt.Sprintf("#%02X%02X%02X", c.r, c.g, c.b)
}
