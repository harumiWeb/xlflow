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
	welcomeBadgeColor = rgbColor{r: 184, g: 245, b: 162}
	welcomeTitleStart = rgbColor{r: 143, g: 211, b: 255}
	welcomeTitleEnd   = rgbColor{r: 184, g: 245, b: 162}
)

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
		badge = lipgloss.NewStyle().
			Foreground(lipgloss.Color(welcomeBadgeColor.hex())).
			Bold(true).
			Render(badge)
		logo = renderGradientBlock(scaffoldWelcomeLogo, welcomeTitleStart, welcomeTitleEnd)
	}
	return badge + "\n\n" + logo + "\n\n"
}

func renderScaffoldWelcomeBadge(text string) string {
	border := strings.Repeat("-", len(text)+2)
	return "+" + border + "+\n| " + text + " |\n+" + border + "+"
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
