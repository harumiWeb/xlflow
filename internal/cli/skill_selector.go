package cli

import (
	"errors"
	"fmt"
	"strings"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"

	"github.com/harumiWeb/xlflow/internal/agentskill"
)

var errSkillSelectionCanceled = errors.New("skill provider selection canceled")

type skillSelectorModel struct {
	providers []agentskill.Provider
	cursor    int
	selected  int
	canceled  bool
}

func newSkillSelectorModel(providers []agentskill.Provider) skillSelectorModel {
	return skillSelectorModel{providers: providers, selected: -1}
}

func (m skillSelectorModel) Init() tea.Cmd {
	return nil
}

func (m skillSelectorModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	key, ok := msg.(tea.KeyMsg)
	if !ok {
		return m, nil
	}
	switch key.String() {
	case "ctrl+c", "esc":
		m.canceled = true
		return m, tea.Quit
	case "up", "k":
		if m.cursor > 0 {
			m.cursor--
		}
	case "down", "j":
		if m.cursor < len(m.providers)-1 {
			m.cursor++
		}
	case "enter":
		m.selected = m.cursor
		return m, tea.Quit
	}
	return m, nil
}

func (m skillSelectorModel) View() string {
	var b strings.Builder
	title := lipgloss.NewStyle().Bold(true).Render("Install xlflow skill")
	hint := lipgloss.NewStyle().Foreground(lipgloss.Color("241")).Render("Choose the agent provider target.")
	b.WriteString(title)
	b.WriteString("\n")
	b.WriteString(hint)
	b.WriteString("\n\n")
	for i, provider := range m.providers {
		cursor := "  "
		nameStyle := lipgloss.NewStyle().Bold(true)
		if i == m.cursor {
			cursor = "> "
			nameStyle = nameStyle.Foreground(lipgloss.Color("39"))
		}
		b.WriteString(cursor)
		b.WriteString(nameStyle.Render(provider.Name))
		b.WriteString("  ")
		b.WriteString(provider.Dir)
		b.WriteString("\n")
		b.WriteString("    ")
		b.WriteString(lipgloss.NewStyle().Foreground(lipgloss.Color("244")).Render(provider.Description))
		b.WriteString("\n")
	}
	b.WriteString("\n")
	b.WriteString(lipgloss.NewStyle().Foreground(lipgloss.Color("241")).Render("Enter select  Esc cancel"))
	return b.String()
}

func runSkillSelector() (agentskill.Provider, error) {
	providers := agentskill.Providers()
	program := tea.NewProgram(newSkillSelectorModel(providers))
	model, err := program.Run()
	if err != nil {
		return agentskill.Provider{}, err
	}
	selector, ok := model.(skillSelectorModel)
	if !ok {
		return agentskill.Provider{}, fmt.Errorf("unexpected skill selector model %T", model)
	}
	if selector.canceled || selector.selected < 0 {
		return agentskill.Provider{}, errSkillSelectionCanceled
	}
	return selector.providers[selector.selected], nil
}
