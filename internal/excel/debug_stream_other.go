//go:build !windows

package excel

import (
	"fmt"
	"io"
	"os"
	"time"
)

type debugStreamSession struct {
	pipePath string
}

func newDebugStreamSession(io.Writer) (*debugStreamSession, error) {
	return &debugStreamSession{
		pipePath: fmt.Sprintf(`\\.\pipe\xlflow-debug-%d-%d`, os.Getpid(), time.Now().UnixNano()),
	}, nil
}

func (s *debugStreamSession) PipePath() string {
	if s == nil {
		return ""
	}
	return s.pipePath
}

func (s *debugStreamSession) Close() error {
	return nil
}

func (s *debugStreamSession) Result() any {
	return nil
}
