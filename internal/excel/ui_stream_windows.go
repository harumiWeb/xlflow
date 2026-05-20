//go:build windows

package excel

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net"
	"os"
	"strings"
	"sync"
	"time"
	"unicode/utf16"

	winio "github.com/Microsoft/go-winio"
)

type uiStreamSession struct {
	pipePath string
	listener net.Listener
	stderr   io.Writer

	done      chan struct{}
	closed    chan struct{}
	closeOnce sync.Once

	mu         sync.Mutex
	events     []map[string]any
	closeErr   error
	activeConn net.Conn
}

func newUIStreamSession(stderr io.Writer) (*uiStreamSession, error) {
	pipePath := fmt.Sprintf(`\\.\pipe\xlflow-ui-%d-%d`, os.Getpid(), time.Now().UnixNano())
	listener, err := winio.ListenPipe(pipePath, &winio.PipeConfig{InputBufferSize: 4096, OutputBufferSize: 4096})
	if err != nil {
		return nil, fmt.Errorf("failed to open UI stream named pipe: %w", err)
	}
	if stderr == nil {
		stderr = os.Stderr
	}
	session := &uiStreamSession{pipePath: pipePath, listener: listener, stderr: stderr, done: make(chan struct{}), closed: make(chan struct{})}
	go session.acceptLoop()
	return session, nil
}

func (s *uiStreamSession) PipePath() string {
	if s == nil {
		return ""
	}
	return s.pipePath
}

func (s *uiStreamSession) Close() error {
	if s == nil {
		return nil
	}
	s.closeOnce.Do(func() {
		close(s.closed)
		_ = s.listener.Close()
		s.mu.Lock()
		if s.activeConn != nil {
			_ = s.activeConn.Close()
		}
		s.mu.Unlock()
	})
	<-s.done
	return s.closeErr
}

func (s *uiStreamSession) Events() []map[string]any {
	if s == nil {
		return nil
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	if len(s.events) == 0 {
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
	return cloned
}

func (s *uiStreamSession) acceptLoop() {
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
		payload, readErr := s.readPayload(conn)
		s.mu.Lock()
		if s.activeConn == conn {
			s.activeConn = nil
		}
		s.mu.Unlock()
		_ = conn.Close()
		for _, line := range decodeUIStreamLines(payload) {
			trimmed := strings.TrimSpace(line)
			if trimmed == "" {
				continue
			}
			var event map[string]any
			if err := json.Unmarshal([]byte(trimmed), &event); err != nil {
				continue
			}
			s.mu.Lock()
			s.events = append(s.events, event)
			s.mu.Unlock()
			if rendered := formatUIStreamEvent(event); rendered != "" {
				_, _ = fmt.Fprintln(s.stderr, rendered)
			}
		}
		if readErr != nil && !isClosedPipeAccept(readErr) {
			s.closeErr = readErr
			return
		}
	}
}

func (s *uiStreamSession) readPayload(conn net.Conn) ([]byte, error) {
	var buffer bytes.Buffer
	chunk := make([]byte, 4096)
	for {
		if err := conn.SetReadDeadline(time.Now().Add(250 * time.Millisecond)); err != nil {
			return buffer.Bytes(), err
		}
		n, err := conn.Read(chunk)
		if n > 0 {
			_, _ = buffer.Write(chunk[:n])
		}
		if err == nil {
			continue
		}
		if err == io.EOF {
			return buffer.Bytes(), nil
		}
		if netErr, ok := err.(net.Error); ok && netErr.Timeout() {
			select {
			case <-s.closed:
				return buffer.Bytes(), nil
			default:
				continue
			}
		}
		select {
		case <-s.closed:
			return buffer.Bytes(), nil
		default:
			return buffer.Bytes(), err
		}
	}
}

func decodeUIStreamLines(payload []byte) []string {
	if len(payload) == 0 {
		return nil
	}
	text := string(payload)
	if looksLikeUTF16LE(payload) {
		text = decodeUTF16LE(payload)
	}
	return strings.Split(text, "\n")
}

func looksLikeUTF16LE(payload []byte) bool {
	if len(payload) < 4 || len(payload)%2 != 0 {
		return false
	}
	limit := len(payload)
	if limit > 32 {
		limit = 32
	}
	zeroHighBytes := 0
	for i := 1; i < limit; i += 2 {
		if payload[i] == 0 {
			zeroHighBytes++
		}
	}
	return zeroHighBytes >= 2
}

func decodeUTF16LE(payload []byte) string {
	if len(payload)%2 != 0 {
		payload = payload[:len(payload)-1]
	}
	words := make([]uint16, 0, len(payload)/2)
	for i := 0; i+1 < len(payload); i += 2 {
		words = append(words, uint16(payload[i])|uint16(payload[i+1])<<8)
	}
	return string(utf16.Decode(words))
}

func isClosedPipeAccept(err error) bool {
	if err == nil {
		return false
	}
	message := strings.ToLower(err.Error())
	return strings.Contains(message, "file has been closed") || strings.Contains(message, "closed network connection")
}

func formatUIStreamEvent(event map[string]any) string {
	if len(event) == 0 {
		return ""
	}
	kind := escapeUIStreamToken(stringifyUIStreamField(event, "kind"))
	id := escapeUIStreamToken(stringifyUIStreamField(event, "dialog_id"))
	source := escapeUIStreamToken(stringifyUIStreamField(event, "response_source"))
	result := escapeUIStreamToken(stringifyUIStreamField(event, "resolved_result"))
	value := escapeUIStreamToken(stringifyUIStreamField(event, "resolved_value"))
	if value != "" && truthyUIStreamField(event, "redacted") {
		value = "[redacted]"
	}
	parts := []string{"xlflow: ui"}
	if kind != "" {
		parts = append(parts, "kind="+kind)
	}
	if id != "" {
		parts = append(parts, "id="+id)
	}
	if source != "" {
		parts = append(parts, "source="+source)
	}
	if result != "" {
		parts = append(parts, "result="+result)
	}
	if value != "" {
		parts = append(parts, "value="+value)
	}
	return strings.Join(parts, " ")
}

func escapeUIStreamToken(value string) string {
	if value == "" {
		return ""
	}
	var builder strings.Builder
	for _, r := range value {
		switch r {
		case '\n':
			builder.WriteString(`\n`)
		case '\r':
			builder.WriteString(`\r`)
		case '\t':
			builder.WriteString(`\t`)
		default:
			if r < 0x20 || r == 0x7f {
				fmt.Fprintf(&builder, `\\x%02X`, r)
				continue
			}
			builder.WriteRune(r)
		}
	}
	return builder.String()
}

func stringifyUIStreamField(event map[string]any, key string) string {
	if event == nil {
		return ""
	}
	value, ok := event[key]
	if !ok || value == nil {
		return ""
	}
	switch typed := value.(type) {
	case string:
		return strings.TrimSpace(typed)
	case fmt.Stringer:
		return strings.TrimSpace(typed.String())
	default:
		return strings.TrimSpace(fmt.Sprint(value))
	}
}

func truthyUIStreamField(event map[string]any, key string) bool {
	if event == nil {
		return false
	}
	value, ok := event[key]
	if !ok || value == nil {
		return false
	}
	switch typed := value.(type) {
	case bool:
		return typed
	case string:
		return strings.EqualFold(strings.TrimSpace(typed), "true")
	default:
		return false
	}
}
