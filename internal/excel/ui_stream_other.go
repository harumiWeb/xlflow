//go:build !windows

package excel

import (
	"fmt"
	"io"
	"os"
	"time"
)

type uiStreamSession struct {
	pipePath string
}

func newUIStreamSession(io.Writer) (*uiStreamSession, error) {
	return &uiStreamSession{
		pipePath: fmt.Sprintf(`\\.\pipe\xlflow-ui-%d-%d`, os.Getpid(), time.Now().UnixNano()),
	}, nil
}

func (s *uiStreamSession) PipePath() string {
	if s == nil {
		return ""
	}
	return s.pipePath
}

func (s *uiStreamSession) Close() error {
	return nil
}

func (s *uiStreamSession) Events() []map[string]any {
	return nil
}
