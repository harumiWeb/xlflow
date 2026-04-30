package cli

import (
	"strings"
	"testing"

	tea "github.com/charmbracelet/bubbletea"

	"github.com/harumiWeb/xlflow/internal/agentskill"
)

func TestSkillSelectorSelectsProvider(t *testing.T) {
	model := newSkillSelectorModel(agentskill.Providers())
	updated, _ := model.Update(tea.KeyMsg{Type: tea.KeyDown})
	updated, _ = updated.(skillSelectorModel).Update(tea.KeyMsg{Type: tea.KeyEnter})
	got := updated.(skillSelectorModel)
	if got.selected != 1 {
		t.Fatalf("selected = %d, want 1", got.selected)
	}
}

func TestSkillSelectorCancels(t *testing.T) {
	model := newSkillSelectorModel(agentskill.Providers())
	updated, _ := model.Update(tea.KeyMsg{Type: tea.KeyEsc})
	got := updated.(skillSelectorModel)
	if !got.canceled {
		t.Fatal("expected canceled selector")
	}
}

func TestSkillSelectorViewIncludesTargets(t *testing.T) {
	view := newSkillSelectorModel(agentskill.Providers()).View()
	for _, want := range []string{"Install xlflow skill", "codex", ".codex/skills"} {
		if !strings.Contains(view, want) {
			t.Fatalf("view missing %q:\n%s", want, view)
		}
	}
}
