package output

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
)

const (
	StatusOK     = "ok"
	StatusFailed = "failed"

	ExitSuccess     = 0
	ExitValidation  = 1
	ExitConfig      = 2
	ExitEnvironment = 3
)

type Error struct {
	Code    string `json:"code,omitempty"`
	Message string `json:"message"`
	Source  string `json:"source,omitempty"`
	Number  int    `json:"number,omitempty"`
	Line    int    `json:"line,omitempty"`
	Phase   string `json:"phase,omitempty"`
}

type Envelope struct {
	Status  string   `json:"status"`
	Command string   `json:"command"`
	Error   *Error   `json:"error"`
	Logs    []string `json:"logs"`

	Diagnostics any `json:"diagnostics,omitempty"`
	Workbook    any `json:"workbook,omitempty"`
	Backup      any `json:"backup,omitempty"`
	Source      any `json:"source,omitempty"`
	Macro       any `json:"macro,omitempty"`
	Macros      any `json:"macros,omitempty"`
	Issues      any `json:"issues,omitempty"`
	Tests       any `json:"tests,omitempty"`
	Diff        any `json:"diff,omitempty"`
	Trace       any `json:"trace,omitempty"`
}

type ExitError struct {
	Code int
	Err  error
}

func (e *ExitError) Error() string {
	if e.Err == nil {
		return fmt.Sprintf("exit code %d", e.Code)
	}
	return e.Err.Error()
}

func (e *ExitError) Unwrap() error {
	return e.Err
}

func WithExitCode(code int, err error) error {
	if err == nil {
		return nil
	}
	return &ExitError{Code: code, Err: err}
}

func ExitCode(err error) int {
	if err == nil {
		return ExitSuccess
	}
	var exitErr *ExitError
	if errors.As(err, &exitErr) {
		return exitErr.Code
	}
	return ExitConfig
}

func New(command string) Envelope {
	return Envelope{
		Status:  StatusOK,
		Command: command,
		Error:   nil,
		Logs:    []string{},
	}
}

func Failure(command string, err Error) Envelope {
	return Envelope{
		Status:  StatusFailed,
		Command: command,
		Error:   &err,
		Logs:    []string{},
	}
}

func Write(w io.Writer, env Envelope, jsonOutput bool) error {
	if jsonOutput {
		enc := json.NewEncoder(w)
		enc.SetIndent("", "  ")
		return enc.Encode(env)
	}
	if env.Status == StatusOK {
		if len(env.Logs) == 0 {
			_, err := fmt.Fprintln(w, "ok")
			return err
		}
		for _, line := range env.Logs {
			if _, err := fmt.Fprintln(w, line); err != nil {
				return err
			}
		}
		return nil
	}
	if env.Error != nil {
		for _, line := range env.Logs {
			if _, err := fmt.Fprintln(w, line); err != nil {
				return err
			}
		}
		_, err := fmt.Fprintln(w, env.Error.Message)
		return err
	}
	_, err := fmt.Fprintln(w, "failed")
	return err
}
