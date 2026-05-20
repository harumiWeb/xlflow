//go:build !windows

package excel

import "io"

type uiStreamSession struct{}

func newUIStreamSession(io.Writer) (*uiStreamSession, error) {
	return nil, nil
}

func (s *uiStreamSession) PipePath() string {
	return ""
}

func (s *uiStreamSession) Close() error {
	return nil
}

func (s *uiStreamSession) Events() []map[string]any {
	return nil
}
