//go:build windows

package excel

import (
	"encoding/json"
	"fmt"
	"io"
	"net"
	"os"
	"strings"
	"sync"
	"time"

	winio "github.com/Microsoft/go-winio"
)

const debugStreamMaxEvents = 200

type debugStreamSession struct {
	pipePath string
	listener net.Listener
	stderr   io.Writer

	done      chan struct{}
	closed    chan struct{}
	closeOnce sync.Once

	mu         sync.Mutex
	events     []map[string]any
	count      int
	truncated  bool
	closeErr   error
	activeConn net.Conn
}

func newDebugStreamSession(stderr io.Writer) (*debugStreamSession, error) {
	pipePath := fmt.Sprintf(`\\.\pipe\xlflow-debug-%d-%d`, os.Getpid(), time.Now().UnixNano())
	listener, err := winio.ListenPipe(pipePath, &winio.PipeConfig{InputBufferSize: 4096, OutputBufferSize: 4096})
	if err != nil {
		return nil, fmt.Errorf("failed to open debug stream named pipe: %w", err)
	}
	if stderr == nil {
		stderr = os.Stderr
	}
	session := &debugStreamSession{pipePath: pipePath, listener: listener, stderr: stderr, done: make(chan struct{}), closed: make(chan struct{})}
	go session.acceptLoop()
	return session, nil
}

func (s *debugStreamSession) PipePath() string {
	if s == nil {
		return ""
	}
	return s.pipePath
}

func (s *debugStreamSession) Close() error {
	if s == nil {
		return nil
	}
	s.closeOnce.Do(func() {
		close(s.closed)
		_ = s.listener.Close()
	})
	<-s.done
	return s.closeErr
}

func (s *debugStreamSession) Result() any {
	if s == nil {
		return nil
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.count == 0 {
		return nil
	}
	cloned := make([]map[string]any, 0, len(s.events))
	for _, event := range s.events {
		copied := make(map[string]any, len(event))
		for k, v := range event {
			copied[k] = v
		}
		cloned = append(cloned, copied)
	}
	result := map[string]any{"events": cloned, "count": s.count}
	if s.truncated {
		result["truncated"] = true
	}
	return result
}

func (s *debugStreamSession) acceptLoop() {
	defer close(s.done)
	for {
		conn, err := s.listener.Accept()
		if err != nil {
			if !isClosedPipeAccept(err) {
				s.closeErr = err
			}
			return
		}
		s.mu.Lock()
		s.activeConn = conn
		s.mu.Unlock()
		readErr := s.readEvents(conn)
		s.mu.Lock()
		if s.activeConn == conn {
			s.activeConn = nil
		}
		s.mu.Unlock()
		_ = conn.Close()
		if readErr != nil && !isClosedPipeAccept(readErr) {
			s.closeErr = readErr
			return
		}
	}
}

func (s *debugStreamSession) readEvents(conn net.Conn) error {
	chunk := make([]byte, 4096)
	pending := make([]byte, 0, len(chunk))
	encoding := uiStreamEncodingUnknown
	for {
		if err := conn.SetReadDeadline(time.Now().Add(uiStreamReadDeadline)); err != nil {
			return err
		}
		n, err := conn.Read(chunk)
		if n > 0 {
			pending = append(pending, chunk[:n]...)
			if len(pending) > uiStreamMaxPendingBytes {
				return fmt.Errorf("debug stream message exceeds %d bytes", uiStreamMaxPendingBytes)
			}
			if encoding == uiStreamEncodingUnknown {
				encoding = detectUIStreamEncoding(pending, false)
			}
			var lines []string
			lines, pending = splitUIStreamLines(pending, encoding, false)
			for _, line := range lines {
				s.handleDebugStreamLine(line)
			}
		}
		if err == nil {
			continue
		}
		if err == io.EOF {
			s.flushDebugStreamPending(&pending, &encoding)
			return nil
		}
		if netErr, ok := err.(net.Error); ok && netErr.Timeout() {
			select {
			case <-s.closed:
				s.flushDebugStreamPending(&pending, &encoding)
				return nil
			default:
				continue
			}
		}
		select {
		case <-s.closed:
			s.flushDebugStreamPending(&pending, &encoding)
			return nil
		default:
			return err
		}
	}
}

func (s *debugStreamSession) flushDebugStreamPending(pending *[]byte, encoding *uiStreamEncoding) {
	if len(*pending) == 0 {
		return
	}
	if *encoding == uiStreamEncodingUnknown {
		*encoding = detectUIStreamEncoding(*pending, true)
	}
	lines, rest := splitUIStreamLines(*pending, *encoding, true)
	*pending = rest
	for _, line := range lines {
		s.handleDebugStreamLine(line)
	}
}

func (s *debugStreamSession) handleDebugStreamLine(line string) {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" {
		return
	}
	var event map[string]any
	if err := json.Unmarshal([]byte(trimmed), &event); err != nil {
		return
	}
	s.mu.Lock()
	s.count++
	if len(s.events) >= debugStreamMaxEvents {
		copy(s.events, s.events[1:])
		s.events[len(s.events)-1] = event
		s.truncated = true
	} else {
		s.events = append(s.events, event)
	}
	s.mu.Unlock()
	if rendered := formatDebugStreamEvent(event); rendered != "" {
		_, _ = fmt.Fprintln(s.stderr, rendered)
	}
}

func formatDebugStreamEvent(event map[string]any) string {
	if len(event) == 0 {
		return ""
	}
	parts := []string{"xlflow: debug"}
	if source := escapeUIStreamToken(stringifyUIStreamField(event, "source")); source != "" {
		parts = append(parts, "source="+source)
	}
	if mode := escapeUIStreamToken(stringifyUIStreamField(event, "runtime_mode")); mode != "" {
		parts = append(parts, "mode="+mode)
	}
	if message := escapeUIStreamToken(stringifyUIStreamField(event, "message")); message != "" {
		parts = append(parts, "message="+message)
	}
	if errText := escapeUIStreamToken(stringifyUIStreamField(event, "error")); errText != "" {
		parts = append(parts, "error="+errText)
	}
	return strings.Join(parts, " ")
}
