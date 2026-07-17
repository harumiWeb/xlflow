package lspserver

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestDocumentsReuseReplaceCloseAndReopenSnapshots(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	saved := "Option Explicit\nPublic Sub Saved()\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(saved), 0o644); err != nil {
		t.Fatal(err)
	}
	uri := pathToFileURI(path)
	docs := newDocuments(root)

	opened, err := docs.open(uri, "Option Explicit\nPublic Sub Opened()\nEnd Sub\n", 4)
	if err != nil {
		t.Fatal(err)
	}
	again, err := docs.getOrRead(uri)
	if err != nil || again.Snapshot != opened.Snapshot {
		t.Fatalf("same revision snapshot = %p, want %p (err=%v)", again.Snapshot, opened.Snapshot, err)
	}
	changed, err := docs.change(uri, "Option Explicit\nPublic Sub Changed()\nEnd Sub\n", 4)
	if err != nil {
		t.Fatal(err)
	}
	if changed.Snapshot == opened.Snapshot || !opened.Snapshot.Retired() {
		t.Fatal("same-version source change did not replace and retire the old snapshot")
	}
	docs.close(uri)
	if !changed.Snapshot.Retired() || len(docs.openDocuments()) != 0 {
		t.Fatal("close did not retire and remove the open snapshot")
	}
	disk, err := docs.getOrRead(uri)
	if err != nil || !strings.Contains(disk.Source, "Saved") {
		t.Fatalf("disk snapshot = %+v, err=%v", disk, err)
	}
	diskAgain, err := docs.getOrRead(uri)
	if err != nil || diskAgain.Snapshot != disk.Snapshot {
		t.Fatalf("disk snapshot was not reused: %p != %p", diskAgain.Snapshot, disk.Snapshot)
	}
	reopened, err := docs.open(uri, "Option Explicit\nPublic Sub Reopened()\nEnd Sub\n", 5)
	if err != nil {
		t.Fatal(err)
	}
	if reopened.Snapshot == disk.Snapshot || !disk.Snapshot.Retired() {
		t.Fatal("reopen did not replace the file-backed snapshot")
	}
	docs.closeAll()
	if !reopened.Snapshot.Retired() {
		t.Fatal("shutdown cleanup did not retire the active snapshot")
	}
}

func TestSnapshotLazySymbolsDoNotPublishIntoReplacement(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	old, err := docs.open(uri, "Sub OldName()\nEnd Sub\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	var loads atomic.Int32
	result := make(chan error, 1)
	go func() {
		_, _, err := old.Snapshot.SourceSymbols(func() ([]intel.Symbol, error) {
			loads.Add(1)
			close(started)
			<-release
			return []intel.Symbol{{Name: "OldName"}}, nil
		})
		result <- err
	}()
	<-started
	newer, err := docs.change(uri, "Sub NewName()\nEnd Sub\n", 2)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-result; err != nil {
		t.Fatal(err)
	}
	if !old.Snapshot.Retired() || newer.Snapshot == old.Snapshot {
		t.Fatal("replacement did not isolate the in-flight snapshot")
	}
	syms, hit, err := newer.Snapshot.SourceSymbols(func() ([]intel.Symbol, error) {
		loads.Add(1)
		return []intel.Symbol{{Name: "NewName"}}, nil
	})
	if err != nil || hit || len(syms) != 1 || syms[0].Name != "NewName" || loads.Load() != 2 {
		t.Fatalf("new snapshot symbols = (%+v, hit=%v, err=%v, loads=%d)", syms, hit, err, loads.Load())
	}
}

func TestDocumentsConcurrentReadersShareSnapshot(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	doc, err := docs.open(uri, "Option Explicit\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	const readers = 32
	var wg sync.WaitGroup
	for range readers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			got, err := docs.getOrRead(uri)
			if err != nil {
				t.Error(err)
				return
			}
			if got.Snapshot != doc.Snapshot {
				t.Errorf("snapshot = %p, want %p", got.Snapshot, doc.Snapshot)
			}
		}()
	}
	wg.Wait()
}

func TestOpenSnapshotWinsOverConcurrentDiskPublication(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "Main.bas")
	uri := pathToFileURI(path)
	docs := newDocuments(root)
	started := make(chan struct{})
	release := make(chan struct{})
	docs.readFile = func(string) ([]byte, error) {
		close(started)
		<-release
		return []byte("Sub Saved()\nEnd Sub\n"), nil
	}
	type result struct {
		doc intel.Document
		err error
	}
	diskResult := make(chan result, 1)
	go func() {
		doc, err := docs.getOrRead(uri)
		diskResult <- result{doc: doc, err: err}
	}()
	<-started
	opened, err := docs.open(uri, "Sub Unsaved()\nEnd Sub\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	got := <-diskResult
	if got.err != nil || got.doc.Snapshot != opened.Snapshot || !strings.Contains(got.doc.Source, "Unsaved") {
		t.Fatalf("disk race returned %+v, err=%v; want open snapshot %p", got.doc, got.err, opened.Snapshot)
	}
}

func TestSnapshotDocumentSymbolsKeepUserFormControlsFresh(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	codePath := filepath.Join(codeDir, "CustomerForm.bas")
	if err := os.WriteFile(codePath, []byte("Option Explicit\nPrivate Sub UserForm_Initialize()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	formPath := filepath.Join(formsDir, "CustomerForm.frm")
	writeForm := func(control string) {
		t.Helper()
		body := "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm\n" +
			"   Begin MSForms.TextBox " + control + "\n   End\nEnd\nAttribute VB_Name = \"CustomerForm\"\n"
		if err := os.WriteFile(formPath, []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	writeForm("txtOld")
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	uri := pathToFileURI(codePath)
	doc, err := s.docs.getOrRead(uri)
	if err != nil {
		t.Fatal(err)
	}
	assertSymbol := func(name string, want bool) {
		t.Helper()
		syms, err := s.analyzer.DocumentSymbols(doc)
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, symbol := range syms {
			found = found || symbol.Name == name
		}
		if found != want {
			t.Fatalf("symbol %q found=%v, want %v; symbols=%+v", name, found, want, syms)
		}
	}
	assertSymbol("txtOld", true)
	writeForm("txtNew")
	assertSymbol("txtOld", false)
	assertSymbol("txtNew", true)
}

func TestServerShutdownRetiresAllSnapshots(t *testing.T) {
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	doc, err := s.docs.open(uri, "Option Explicit\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	if err := s.shutdown(nil); err != nil {
		t.Fatal(err)
	}
	if !doc.Snapshot.Retired() || len(s.docs.openDocuments()) != 0 {
		t.Fatal("server shutdown did not retire and remove snapshots")
	}
}

func TestCloseAllPreventsInFlightDiskSnapshotPublication(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "Main.bas")
	uri := pathToFileURI(path)
	docs := newDocuments(root)
	started := make(chan struct{})
	release := make(chan struct{})
	docs.readFile = func(string) ([]byte, error) {
		close(started)
		<-release
		return []byte("Option Explicit\n"), nil
	}
	errResult := make(chan error, 1)
	go func() {
		_, err := docs.getOrRead(uri)
		errResult <- err
	}()
	<-started
	docs.closeAll()
	close(release)
	if err := <-errResult; !errors.Is(err, errDocumentsClosed) {
		t.Fatalf("in-flight read error = %v, want %v", err, errDocumentsClosed)
	}
	docs.mu.RLock()
	defer docs.mu.RUnlock()
	if !docs.closed || len(docs.docs) != 0 {
		t.Fatalf("closed store republished entries: closed=%v entries=%d", docs.closed, len(docs.docs))
	}
}

func TestConcurrentStaleChangeCannotOverwriteNewerSnapshot(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	if _, err := docs.open(uri, "Sub Original()\nEnd Sub\n", 1); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	var publications atomic.Int32
	docs.beforePublish = func() {
		if publications.Add(1) == 1 {
			close(started)
			<-release
		}
	}
	staleResult := make(chan intel.Document, 1)
	staleErr := make(chan error, 1)
	go func() {
		doc, err := docs.change(uri, "Sub OlderConcurrentChange()\nEnd Sub\n", 2)
		staleResult <- doc
		staleErr <- err
	}()
	<-started
	newer, err := docs.change(uri, "Sub NewerConcurrentChange()\nEnd Sub\n", 3)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-staleErr; err != nil {
		t.Fatal(err)
	}
	stale := <-staleResult
	if stale.Snapshot != newer.Snapshot || !strings.Contains(stale.Source, "NewerConcurrentChange") {
		t.Fatalf("stale change returned %+v, want newer snapshot %p", stale, newer.Snapshot)
	}
	active, err := docs.getOrRead(uri)
	if err != nil || active.Snapshot != newer.Snapshot || active.Version != 3 {
		t.Fatalf("active snapshot = %+v, err=%v; want version 3 snapshot %p", active, err, newer.Snapshot)
	}
}

func TestDelayedNewerChangeReplacesConcurrentOlderSnapshot(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	if _, err := docs.open(uri, "Sub Original()\nEnd Sub\n", 1); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	var publications atomic.Int32
	docs.beforePublish = func() {
		if publications.Add(1) == 1 {
			close(started)
			<-release
		}
	}
	newerResult := make(chan intel.Document, 1)
	newerErr := make(chan error, 1)
	go func() {
		doc, err := docs.change(uri, "Sub DelayedNewerChange()\nEnd Sub\n", 3)
		newerResult <- doc
		newerErr <- err
	}()
	<-started
	older, err := docs.change(uri, "Sub EarlierPublishedChange()\nEnd Sub\n", 2)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-newerErr; err != nil {
		t.Fatal(err)
	}
	newer := <-newerResult
	if newer.Version != 3 || newer.Snapshot == older.Snapshot || !older.Snapshot.Retired() {
		t.Fatalf("delayed newer result = %+v; older snapshot retired=%v", newer, older.Snapshot.Retired())
	}
	active, err := docs.getOrRead(uri)
	if err != nil || active.Snapshot != newer.Snapshot || !strings.Contains(active.Source, "DelayedNewerChange") {
		t.Fatalf("active snapshot = %+v, err=%v; want delayed version 3", active, err)
	}
}

func TestChangeRegistersURIAliasForClose(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	alias := strings.Replace(uri, "Main.bas", "%4Dain.bas", 1)
	if alias == uri {
		t.Fatalf("could not create URI alias from %q", uri)
	}
	opened, err := docs.open(uri, "Sub Original()\nEnd Sub\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	changed, err := docs.change(alias, "Sub ChangedThroughAlias()\nEnd Sub\n", 2)
	if err != nil {
		t.Fatal(err)
	}
	if changed.Snapshot == opened.Snapshot {
		t.Fatal("alias change did not replace the original snapshot")
	}
	docs.close(alias)
	if !changed.Snapshot.Retired() || len(docs.openDocuments()) != 0 {
		t.Fatal("closing through the registered alias did not remove the active snapshot")
	}
}

func TestPreCloseChangeCannotPublishIntoReopenedLifecycle(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	if _, err := docs.open(uri, "Sub OriginalLifecycle()\nEnd Sub\n", 10); err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	var publications atomic.Int32
	docs.beforePublish = func() {
		if publications.Add(1) == 1 {
			close(started)
			<-release
		}
	}
	staleResult := make(chan intel.Document, 1)
	staleErr := make(chan error, 1)
	go func() {
		doc, err := docs.change(uri, "Sub PreCloseCandidate()\nEnd Sub\n", 99)
		staleResult <- doc
		staleErr <- err
	}()
	<-started
	docs.close(uri)
	reopened, err := docs.open(uri, "Sub ReopenedLifecycle()\nEnd Sub\n", 1)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-staleErr; err != nil {
		t.Fatal(err)
	}
	stale := <-staleResult
	if stale.Snapshot != reopened.Snapshot || !strings.Contains(stale.Source, "ReopenedLifecycle") {
		t.Fatalf("pre-close candidate escaped into reopened lifecycle: %+v", stale)
	}
	active, err := docs.getOrRead(uri)
	if err != nil || active.Snapshot != reopened.Snapshot || active.Version != 1 {
		t.Fatalf("active snapshot = %+v, err=%v; want reopened lifecycle", active, err)
	}
}

func TestDiskSnapshotRetriesAfterConcurrentPublication(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(filepath.Join(t.TempDir(), "Main.bas"))
	var reads atomic.Int32
	docs.readFile = func(string) ([]byte, error) {
		if reads.Add(1) == 1 {
			return []byte("Sub StaleDiskRead()\nEnd Sub\n"), nil
		}
		return []byte("Sub FreshDiskRead()\nEnd Sub\n"), nil
	}
	started := make(chan struct{})
	release := make(chan struct{})
	var publications atomic.Int32
	docs.beforePublish = func() {
		if publications.Add(1) == 1 {
			close(started)
			<-release
		}
	}
	staleResult := make(chan intel.Document, 1)
	staleErr := make(chan error, 1)
	go func() {
		doc, err := docs.getOrRead(uri)
		staleResult <- doc
		staleErr <- err
	}()
	<-started
	fresh, err := docs.getOrRead(uri)
	if err != nil {
		t.Fatal(err)
	}
	close(release)
	if err := <-staleErr; err != nil {
		t.Fatal(err)
	}
	retried := <-staleResult
	if retried.Snapshot != fresh.Snapshot || !strings.Contains(retried.Source, "FreshDiskRead") {
		t.Fatalf("stale disk read replaced fresh publication: retried=%+v fresh=%+v", retried, fresh)
	}
	if reads.Load() < 3 {
		t.Fatalf("disk reads = %d, want stale read plus fresh publication and retry", reads.Load())
	}
}
