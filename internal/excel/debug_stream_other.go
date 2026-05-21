//go:build !windows

package excel

import "io"

type debugStreamSession struct{}

func newDebugStreamSession(io.Writer) (*debugStreamSession, error) {
	return nil, nil
}

func (s *debugStreamSession) PipePath() string {
	return ""
}

func (s *debugStreamSession) Close() error {
	return nil
}

func (s *debugStreamSession) Result() any {
	return nil
}
