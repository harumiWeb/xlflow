//go:build windows

package coordination

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"golang.org/x/sys/windows"
)

func TestNewWorkbookIdentityEquivalentWindowsPaths(t *testing.T) {
	baseDir := t.TempDir()
	directPath := filepath.Join(baseDir, "Project", "Book.xlsm")

	tests := []struct {
		name string
		path string
	}{
		{name: "different case", path: strings.ToLower(directPath)},
		{name: "forward separators", path: strings.ReplaceAll(directPath, `\`, "/")},
		{name: "extended prefix", path: `\\?\` + directPath},
	}
	want, err := NewWorkbookIdentity(baseDir, directPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(direct): %v", err)
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := NewWorkbookIdentity(baseDir, tt.path)
			if err != nil {
				t.Fatalf("NewWorkbookIdentity(%q): %v", tt.path, err)
			}
			if got.LockID != want.LockID {
				t.Fatalf("LockID = %q, want %q (canonical paths %q and %q)", got.LockID, want.LockID, got.CanonicalPath, want.CanonicalPath)
			}
		})
	}
	if len(want.CanonicalPath) < 2 || want.CanonicalPath[1] != ':' || want.CanonicalPath[0] < 'A' || want.CanonicalPath[0] > 'Z' {
		t.Fatalf("CanonicalPath drive is not normalized: %q", want.CanonicalPath)
	}
}

func TestNewWorkbookIdentityEquivalentShortAndLongParentPaths(t *testing.T) {
	longBase := strings.TrimSpace(os.Getenv("ProgramFiles"))
	if longBase == "" {
		longBase = t.TempDir()
	}
	shortBase, err := windowsShortPath(longBase)
	if err != nil {
		t.Skipf("Windows short path is unavailable: %v", err)
	}
	if strings.EqualFold(shortBase, longBase) {
		t.Skip("Windows short names are disabled for the selected parent directory")
	}

	fromLong, err := NewWorkbookIdentity(longBase, filepath.Join("missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(long): %v", err)
	}
	fromShort, err := NewWorkbookIdentity(shortBase, filepath.Join("missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(short): %v", err)
	}
	if fromShort != fromLong {
		t.Fatalf("short parent identity = %#v, want long parent identity %#v", fromShort, fromLong)
	}
}

func TestNewWorkbookIdentityEquivalentUNCAndExtendedUNC(t *testing.T) {
	baseDir := t.TempDir()
	direct := `\\server\share\Project\Book.xlsm`
	extended := `\\?\UNC\server\share\Project\Book.xlsm`

	first, err := NewWorkbookIdentity(baseDir, direct)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(UNC): %v", err)
	}
	second, err := NewWorkbookIdentity(baseDir, extended)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(extended UNC): %v", err)
	}
	if first.LockID != second.LockID {
		t.Fatalf("UNC identities differ: %#v != %#v", first, second)
	}
	if !strings.HasPrefix(first.CanonicalPath, `\\server\share\`) {
		t.Fatalf("CanonicalPath = %q, want UNC semantics preserved", first.CanonicalPath)
	}
}

func TestNewWorkbookIdentityResolvesWindowsSymlink(t *testing.T) {
	baseDir := t.TempDir()
	target := filepath.Join(baseDir, "target.xlsm")
	if err := os.WriteFile(target, []byte("workbook placeholder"), 0o600); err != nil {
		t.Fatalf("write target: %v", err)
	}
	link := filepath.Join(baseDir, "linked.xlsm")
	if err := os.Symlink(target, link); err != nil {
		t.Skipf("creating a Windows symlink requires privileges unavailable here: %v", err)
	}

	targetIdentity, err := NewWorkbookIdentity(baseDir, target)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(target): %v", err)
	}
	linkIdentity, err := NewWorkbookIdentity(baseDir, link)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(link): %v", err)
	}
	if targetIdentity != linkIdentity {
		t.Fatalf("symlink identity = %#v, want %#v", linkIdentity, targetIdentity)
	}
}

func TestNewWorkbookIdentityFallsBackForBrokenWindowsSymlink(t *testing.T) {
	baseDir := t.TempDir()
	target := filepath.Join(baseDir, "missing-target.xlsm")
	link := filepath.Join(baseDir, "broken-link.xlsm")
	if err := os.Symlink(target, link); err != nil {
		t.Skipf("creating a Windows symlink requires privileges unavailable here: %v", err)
	}

	first, err := NewWorkbookIdentity(baseDir, link)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(first): %v", err)
	}
	second, err := NewWorkbookIdentity(baseDir, link)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(second): %v", err)
	}
	if first != second {
		t.Fatalf("broken symlink fallback is not deterministic: %#v != %#v", first, second)
	}
	if filepath.Base(first.CanonicalPath) != filepath.Base(link) {
		t.Fatalf("CanonicalPath = %q, want broken link basename %q preserved", first.CanonicalPath, filepath.Base(link))
	}
	targetIdentity, err := NewWorkbookIdentity(baseDir, target)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(target): %v", err)
	}
	if first.LockID == targetIdentity.LockID {
		t.Fatalf("broken link identity %q unexpectedly matched missing target identity", first.LockID)
	}
}

func TestNewWorkbookIdentityResolvesWindowsJunction(t *testing.T) {
	baseDir := t.TempDir()
	targetDir := filepath.Join(baseDir, "target")
	if err := os.Mkdir(targetDir, 0o700); err != nil {
		t.Fatalf("mkdir target: %v", err)
	}
	target := filepath.Join(targetDir, "book.xlsm")
	if err := os.WriteFile(target, []byte("workbook placeholder"), 0o600); err != nil {
		t.Fatalf("write target: %v", err)
	}

	junction := filepath.Join(baseDir, "junction")
	if output, err := exec.Command("cmd.exe", "/d", "/c", "mklink", "/J", junction, targetDir).CombinedOutput(); err != nil {
		t.Skipf("creating a Windows junction is unavailable: %v (%s)", err, strings.TrimSpace(string(output)))
	}
	defer func() {
		if err := os.Remove(junction); err != nil && !os.IsNotExist(err) {
			t.Errorf("remove junction: %v", err)
		}
	}()

	directIdentity, err := NewWorkbookIdentity(baseDir, target)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(target): %v", err)
	}
	junctionIdentity, err := NewWorkbookIdentity(baseDir, filepath.Join(junction, "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(junction): %v", err)
	}
	if directIdentity != junctionIdentity {
		t.Fatalf("junction identity = %#v, want %#v", junctionIdentity, directIdentity)
	}
}

func TestNewWorkbookIdentityResolvesWindowsJunctionParentForMissingWorkbook(t *testing.T) {
	baseDir := t.TempDir()
	targetDir := filepath.Join(baseDir, "target-missing")
	if err := os.Mkdir(targetDir, 0o700); err != nil {
		t.Fatalf("mkdir target: %v", err)
	}

	junction := filepath.Join(baseDir, "junction-missing")
	if output, err := exec.Command("cmd.exe", "/d", "/c", "mklink", "/J", junction, targetDir).CombinedOutput(); err != nil {
		t.Skipf("creating a Windows junction is unavailable: %v (%s)", err, strings.TrimSpace(string(output)))
	}
	defer func() {
		if err := os.Remove(junction); err != nil && !os.IsNotExist(err) {
			t.Errorf("remove junction: %v", err)
		}
	}()

	directIdentity, err := NewWorkbookIdentity(baseDir, filepath.Join(targetDir, "missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(target): %v", err)
	}
	junctionIdentity, err := NewWorkbookIdentity(baseDir, filepath.Join(junction, "missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(junction): %v", err)
	}
	if junctionIdentity != directIdentity {
		t.Fatalf("missing workbook junction identity = %#v, want %#v", junctionIdentity, directIdentity)
	}
}

func windowsShortPath(path string) (string, error) {
	input, err := windows.UTF16PtrFromString(path)
	if err != nil {
		return "", err
	}
	buffer := make([]uint16, 512)
	for {
		length, err := windows.GetShortPathName(input, &buffer[0], uint32(len(buffer)))
		if err != nil {
			return "", err
		}
		if length == 0 {
			return "", fmt.Errorf("GetShortPathName returned an empty path")
		}
		if length < uint32(len(buffer)) {
			return windows.UTF16ToString(buffer[:length]), nil
		}
		buffer = make([]uint16, length+1)
	}
}
