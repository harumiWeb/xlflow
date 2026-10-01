//go:build windows

package filepull

import (
	"errors"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/coordination"
)

func TestConcurrentPullsSerializeSameSourceTreeAndAllowDisjointRoots(t *testing.T) {
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	firstRoot := t.TempDir()
	secondRoot := t.TempDir()
	firstWorkbook := writeWorkbook(t, firstRoot, readFixture(t, "p1_compiled.bin"))
	secondWorkbook := writeWorkbook(t, secondRoot, readFixture(t, "p1_compiled.bin"))

	originalPublish := publishArtifact
	t.Cleanup(func() { publishArtifact = originalPublish })
	firstStarted := make(chan struct{})
	releaseFirst := make(chan struct{})
	secondPublished := make(chan struct{}, 1)
	blocked := false
	publishArtifact = func(path string, body []byte, validate func(string) error) (coordination.PublishResult, error) {
		if strings.HasPrefix(path, firstRoot) && !blocked {
			blocked = true
			close(firstStarted)
			<-releaseFirst
		}
		if strings.HasPrefix(path, secondRoot) {
			select {
			case secondPublished <- struct{}{}:
			default:
			}
		}
		return originalPublish(path, body, validate)
	}

	firstDone := make(chan error, 1)
	go func() {
		_, runErr := PullContext(t.Context(), firstRoot, testConfig(), firstWorkbook, PullOptions{Coordination: manager})
		firstDone <- runErr
	}()
	<-firstStarted
	if _, err := PullContext(t.Context(), firstRoot, testConfig(), firstWorkbook, PullOptions{Coordination: manager}); !errors.Is(err, coordination.ErrSourceTreeBusy) {
		t.Fatalf("same-root concurrent pull error = %v", err)
	}
	secondDone := make(chan error, 1)
	go func() {
		_, runErr := PullContext(t.Context(), secondRoot, testConfig(), secondWorkbook, PullOptions{Coordination: manager})
		secondDone <- runErr
	}()
	select {
	case <-secondPublished:
	case <-time.After(5 * time.Second):
		t.Fatal("disjoint source tree did not publish independently")
	}
	if err := <-secondDone; err != nil {
		t.Fatal(err)
	}
	close(releaseFirst)
	if err := <-firstDone; err != nil {
		t.Fatal(err)
	}
}
