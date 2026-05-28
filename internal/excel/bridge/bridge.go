package bridge

import (
	"context"
	"errors"
	"strings"
)

const EnvBridge = "XLFLOW_EXCEL_BRIDGE"

var ErrInvalidMode = errors.New("bridge mode must be one of auto, powershell, dotnet")

type Mode string

const (
	ModeAuto       Mode = "auto"
	ModePowerShell Mode = "powershell"
	ModeDotNet     Mode = "dotnet"
)

type Request struct {
	Command string
	Args    map[string]string
}

type Response struct {
	Stdout   []byte
	Stderr   []byte
	TimedOut bool
}

type Info struct {
	Name    string
	Version string
}

type Provider interface {
	Name() string
	Supports(command string) bool
	Info(context.Context) (Info, error)
	Execute(context.Context, Request) (Response, error)
}

type ErrorKind string

const (
	ErrorUnsupportedHost   ErrorKind = "unsupported_host"
	ErrorPowerShellMissing ErrorKind = "powershell_missing"
	ErrorScriptNotFound    ErrorKind = "script_not_found"
)

type Error struct {
	Kind    ErrorKind
	Message string
	Err     error
}

func (e *Error) Error() string {
	if e == nil {
		return ""
	}
	return e.Message
}

func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Err
}

func ParseMode(raw string) (Mode, error) {
	mode := Mode(strings.ToLower(strings.TrimSpace(raw)))
	if mode == "" {
		return "", ErrInvalidMode
	}
	switch mode {
	case ModeAuto, ModePowerShell, ModeDotNet:
		return mode, nil
	default:
		return "", ErrInvalidMode
	}
}

func ResolveMode(cli, env, cfg string) (Mode, string, error) {
	for _, candidate := range []struct {
		value  string
		source string
	}{
		{value: cli, source: "cli"},
		{value: env, source: "env"},
		{value: cfg, source: "config"},
	} {
		if strings.TrimSpace(candidate.value) == "" {
			continue
		}
		mode, err := ParseMode(candidate.value)
		if err != nil {
			return "", candidate.source, err
		}
		return mode, candidate.source, nil
	}
	return ModeAuto, "default", nil
}
