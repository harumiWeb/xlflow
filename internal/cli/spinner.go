package cli

import (
	"fmt"
	"io"
	"time"

	tea "github.com/charmbracelet/bubbletea"
)

type spinnerDoneMsg struct {
	err error
}

type spinnerTickMsg time.Time

type spinnerModel struct {
	label  string
	done   <-chan error
	frames []string
	index  int
	err    error
}

func newSpinnerModel(label string, done <-chan error) spinnerModel {
	return spinnerModel{
		label:  label,
		done:   done,
		frames: []string{"|", "/", "-", "\\"},
	}
}

func (m spinnerModel) Init() tea.Cmd {
	return tea.Batch(waitForSpinner(m.done), tickSpinner())
}

func (m spinnerModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case spinnerDoneMsg:
		m.err = msg.err
		return m, tea.Quit
	case spinnerTickMsg:
		m.index = (m.index + 1) % len(m.frames)
		return m, tickSpinner()
	default:
		return m, nil
	}
}

func (m spinnerModel) View() string {
	return fmt.Sprintf("%s %s\n", m.frames[m.index], m.label)
}

func waitForSpinner(done <-chan error) tea.Cmd {
	return func() tea.Msg {
		return spinnerDoneMsg{err: <-done}
	}
}

func tickSpinner() tea.Cmd {
	return tea.Tick(120*time.Millisecond, func(t time.Time) tea.Msg {
		return spinnerTickMsg(t)
	})
}

func runSpinner(w io.Writer, label string, fn func() error) error {
	done := make(chan error, 1)
	go func() {
		done <- fn()
	}()
	program := tea.NewProgram(newSpinnerModel(label, done), tea.WithInput(nil), tea.WithOutput(w))
	model, err := program.Run()
	if err != nil {
		return err
	}
	if spinner, ok := model.(spinnerModel); ok {
		return spinner.err
	}
	return nil
}
