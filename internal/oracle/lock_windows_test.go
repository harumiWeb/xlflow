//go:build windows

package oracle

import (
	"bufio"
	"context"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
	"time"
)

func TestFileBatchLockRejectsConcurrentProcessAndReleasesAfterCrash(t *testing.T) {
	lockPath := filepath.Join(t.TempDir(), "vbe-oracle.lock")
	cmd := exec.Command(os.Args[0], "-test.timeout=0", "-test.run=^TestOracleBatchLockHelper$")
	cmd.Env = append(os.Environ(),
		"XLFLOW_ORACLE_LOCK_HELPER=1",
		"XLFLOW_ORACLE_LOCK_PATH="+lockPath,
	)
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		t.Fatalf("StdoutPipe: %v", err)
	}
	cmd.Stderr = os.Stderr
	if err := cmd.Start(); err != nil {
		t.Fatalf("start helper: %v", err)
	}
	ready := make(chan string, 1)
	scanDone := make(chan struct{})
	go func() {
		defer close(scanDone)
		scanner := bufio.NewScanner(stdout)
		if scanner.Scan() {
			ready <- scanner.Text()
			return
		}
		ready <- ""
	}()
	killed := false
	stopHelper := func() {
		if killed {
			return
		}
		killed = true
		if cmd.Process != nil {
			_ = cmd.Process.Kill()
		}
		<-scanDone
		_ = cmd.Wait()
	}
	defer stopHelper()
	select {
	case line := <-ready:
		if line != "READY" {
			t.Fatalf("helper output = %q", line)
		}
	case <-time.After(10 * time.Second):
		stopHelper()
		t.Fatal("helper did not acquire oracle lock")
	}

	lock := fileBatchLock{path: lockPath}
	if _, err := lock.Acquire(context.Background()); !errors.Is(err, errOracleAlreadyRunning) {
		t.Fatalf("contention error = %v, want oracle_already_running", err)
	}

	stopHelper()
	if cmd.ProcessState == nil || cmd.ProcessState.Success() {
		t.Fatal("killed helper unexpectedly exited successfully")
	}

	deadline := time.Now().Add(5 * time.Second)
	for {
		release, acquireErr := lock.Acquire(context.Background())
		if acquireErr == nil {
			release()
			return
		}
		if !errors.Is(acquireErr, errOracleAlreadyRunning) || time.Now().After(deadline) {
			t.Fatalf("acquire after crash: %v", acquireErr)
		}
		time.Sleep(50 * time.Millisecond)
	}
}

func TestOracleBatchLockHelper(t *testing.T) {
	if os.Getenv("XLFLOW_ORACLE_LOCK_HELPER") != "1" {
		t.Skip("subprocess helper")
	}
	lock := fileBatchLock{path: os.Getenv("XLFLOW_ORACLE_LOCK_PATH")}
	release, err := lock.Acquire(context.Background())
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	defer release()
	fmt.Println("READY")
	timer := time.NewTimer(10 * time.Minute)
	defer timer.Stop()
	<-timer.C
}
