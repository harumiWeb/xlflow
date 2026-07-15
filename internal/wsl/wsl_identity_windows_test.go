//go:build windows

package wsl

import (
	"context"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/coordination"
)

func TestDelegatedWSLPathSharesWindowsWorkbookIdentity(t *testing.T) {
	translated, err := ToWindowsPathWith(context.Background(), "/mnt/c/dev/project/Book.xlsm", func(_ context.Context, name string, args ...string) ([]byte, error) {
		if name != "wslpath" || !reflect.DeepEqual(args, []string{"-w", "/mnt/c/dev/project/Book.xlsm"}) {
			t.Fatalf("unexpected translation command: %s %v", name, args)
		}
		return []byte("C:\\dev\\project\\Book.xlsm\n"), nil
	})
	if err != nil {
		t.Fatalf("ToWindowsPathWith: %v", err)
	}

	fromWSL, err := coordination.NewWorkbookIdentity(`C:\dev\project`, translated)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(translated): %v", err)
	}
	direct, err := coordination.NewWorkbookIdentity(`C:\dev\project`, `c:/DEV/project/book.xlsm`)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(direct): %v", err)
	}
	if fromWSL.LockID != direct.LockID {
		t.Fatalf("delegated identity = %q, direct identity = %q", fromWSL.LockID, direct.LockID)
	}
}
