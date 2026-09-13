package sourceencoding

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"unicode/utf8"

	"golang.org/x/text/encoding/japanese"
)

func TestValidateReportsFirstInvalidUTF8Position(t *testing.T) {
	source := []byte("Option Explicit\r\nSub Run()\r\n  x = \xff\r\nEnd Sub\r\n")
	err := Validate("src/modules/Main.bas", source)
	var encodingErr *Error
	if encodingErr, _ = errors.AsType[*Error](err); encodingErr == nil {
		t.Fatalf("Validate() error = %v, want *Error", err)
	}
	if encodingErr.Status != StatusInvalidUTF8 || encodingErr.Position.Offset != 34 || encodingErr.Position.Line != 3 || encodingErr.Position.ByteColumn != 7 {
		t.Fatalf("encoding error = %+v", encodingErr)
	}
}

func TestInspectReportsUTF16BOMStatus(t *testing.T) {
	result := Inspect("src/modules/Utf16.bas", []byte{0xff, 0xfe, 'x'})
	if result.Status != StatusUTF16BOM || result.Offset != 0 || result.Line != 1 || result.ByteColumn != 1 {
		t.Fatalf("result = %+v, want UTF-16 BOM at the first byte", result)
	}
	if err := Validate("src/modules/Utf16.bas", []byte{0xfe, 0xff, 'x'}); err == nil {
		t.Fatal("Validate accepted UTF-16 BOM")
	} else {
		var encodingErr *Error
		if encodingErr, _ = errors.AsType[*Error](err); encodingErr == nil || encodingErr.Status != StatusUTF16BOM {
			t.Fatalf("Validate error = %v, want UTF-16 BOM", err)
		}
	}
}

func TestDecodeCP932IsStrictAndPreservesJapaneseText(t *testing.T) {
	// Use actual Japanese text rather than a preexisting mojibake fixture.
	const japaneseText = "' \u65e5\u672c\u8a9e\u30b3\u30e1\u30f3\u30c8\r\nvalue = \"\u4ee3\u5165\"\r\n"
	want := japaneseText
	encoded, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte(want))
	if err != nil {
		t.Fatal(err)
	}
	decoded, err := DecodeCP932("Main.cls", encoded)
	if err != nil {
		t.Fatal(err)
	}
	if string(decoded) != want || !utf8.Valid(decoded) {
		t.Fatalf("decoded = %q, want %q", decoded, want)
	}
	if _, err := DecodeCP932("Main.cls", []byte{0x81}); err == nil {
		t.Fatal("invalid trailing CP932 lead byte was accepted")
	}
	if _, err := DecodeCP932("Main.cls", []byte{0xff, 0xfe, 0x00, 0x00}); err != nil {
		var encodingErr *Error
		if encodingErr, _ = errors.AsType[*Error](err); encodingErr == nil || encodingErr.Status != StatusUTF16BOM {
			t.Fatalf("UTF-16 BOM error = %v", err)
		}
	} else {
		t.Fatal("UTF-16 BOM was accepted")
	}
}

func TestCheckDiscoversNestedRootsTestsAndSidecarForms(t *testing.T) {
	root := t.TempDir()
	write := func(path string, body []byte) {
		t.Helper()
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, body, 0o644); err != nil {
			t.Fatal(err)
		}
	}
	write(filepath.Join(root, "src", "modules", "nested", "Main.bas"), []byte("Sub Main()\nEnd Sub\n"))
	write(filepath.Join(root, "src", "forms", "Form1.frm"), []byte("VERSION 5.00\nBegin VB.Form Form1\nEnd\n"))
	write(filepath.Join(root, "src", "forms", "code", "Form1.bas"), []byte("Private Sub Clicked()\nEnd Sub\n"))
	write(filepath.Join(root, "src", "forms", "Form1.frx"), []byte{0x00, 0xff})
	write(filepath.Join(root, "tests", "nested", "Test.cls"), []byte("VERSION 1.0\n"))

	result, err := Check(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules", Forms: "src/forms"}})
	if err != nil {
		t.Fatal(err)
	}
	got := make([]string, 0, len(result.Files))
	for _, file := range result.Files {
		got = append(got, file.Path)
	}
	want := []string{"src/forms/code/Form1.bas", "src/forms/Form1.frm", "src/modules/nested/Main.bas", "tests/nested/Test.cls"}
	if runtime.GOOS != "windows" {
		want = []string{"src/forms/Form1.frm", "src/forms/code/Form1.bas", "src/modules/nested/Main.bas", "tests/nested/Test.cls"}
	}
	if len(got) != len(want) {
		t.Fatalf("files = %#v, want %#v", got, want)
	}
	for index := range want {
		if got[index] != want[index] {
			t.Fatalf("files = %#v, want %#v", got, want)
		}
	}
}

func TestCheckReportsBOMAndMultipleInvalidFilesWithoutScanningFRX(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "B.bas"), []byte{0xef, 0xbb, 0xbf, 'x'}, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "A.cls"), []byte{'x', 0xff}, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "ignored.frx"), []byte{0xff}, 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := Check(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}})
	if err == nil {
		t.Fatal("Check accepted invalid files")
	}
	var violations *ValidationErrors
	if violations, _ = errors.AsType[*ValidationErrors](err); violations == nil || len(violations.Errors) != 2 {
		t.Fatalf("Check error = %v, want two validation errors", err)
	}
	if result.Summary.Total != 2 || result.Summary.Invalid != 2 || result.Files[0].Path != "src/modules/A.cls" || result.Files[1].Path != "src/modules/B.bas" {
		t.Fatalf("result = %+v", result)
	}
	if result.Files[0].Status != StatusInvalidUTF8 || result.Files[1].Status != StatusUTF8BOM {
		t.Fatalf("statuses = %#v", result.Files)
	}
	_, err = Check(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}, Paths: []string{"src/modules/ignored.frx"}})
	var scope *ScopeError
	if scope, _ = errors.AsType[*ScopeError](err); scope == nil {
		t.Fatalf("explicit .frx path error = %v, want ScopeError", err)
	}
}

func TestConvertDoesNotMutateWhenAnyCP932InputIsInvalid(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	goodText := "' 日本語\r\n"
	good, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte(goodText))
	if err != nil {
		t.Fatal(err)
	}
	goodPath := filepath.Join(moduleDir, "A.cls")
	badPath := filepath.Join(moduleDir, "B.cls")
	if err := os.WriteFile(goodPath, good, 0o644); err != nil {
		t.Fatal(err)
	}
	bad := []byte{0x81}
	if err := os.WriteFile(badPath, bad, 0o644); err != nil {
		t.Fatal(err)
	}

	_, err = Convert(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}}, CP932)
	var violations *ValidationErrors
	if violations, _ = errors.AsType[*ValidationErrors](err); violations == nil || len(violations.Errors) != 1 {
		t.Fatalf("Convert error = %v, want one validation error", err)
	}
	after, readErr := os.ReadFile(goodPath)
	if readErr != nil {
		t.Fatal(readErr)
	}
	if string(after) != string(good) {
		t.Fatalf("valid file was changed before validation completed: %x -> %x", good, after)
	}
	if afterBad, readErr := os.ReadFile(badPath); readErr != nil || string(afterBad) != string(bad) {
		t.Fatalf("invalid file changed: %x, err=%v", afterBad, readErr)
	}
}

func TestConvertStagesAllFilesBeforeChangingOriginals(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Japanese.cls")
	originalText := "' 日本語\r\nvalue = \"代入\"\r\n"
	original, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte(originalText))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, original, 0o640); err != nil {
		t.Fatal(err)
	}
	unchangedPath := filepath.Join(moduleDir, "Ascii.bas")
	unchanged := []byte("Sub Run()\r\nEnd Sub\r\n")
	if err := os.WriteFile(unchangedPath, unchanged, 0o600); err != nil {
		t.Fatal(err)
	}

	result, err := Convert(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}}, CP932)
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Converted != 1 || result.Summary.Unchanged != 1 {
		t.Fatalf("summary = %+v", result.Summary)
	}
	converted, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(converted) != originalText || !utf8.Valid(converted) {
		t.Fatalf("converted = %q", converted)
	}
	unchangedAfter, err := os.ReadFile(unchangedPath)
	if err != nil {
		t.Fatal(err)
	}
	if string(unchangedAfter) != string(unchanged) {
		t.Fatalf("UTF-8 source changed: %q", unchangedAfter)
	}
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if runtime.GOOS != "windows" && info.Mode().Perm() != 0o640 {
		t.Fatalf("mode = %o, want 640", info.Mode().Perm())
	}
}

func TestDiscoverRejectsExplicitPathOutsideManagedRoots(t *testing.T) {
	root := t.TempDir()
	outside := filepath.Join(filepath.Dir(root), "outside.bas")
	if err := os.WriteFile(outside, []byte("Sub Run()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = os.Remove(outside) })
	_, err := Check(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}, Paths: []string{outside}})
	var scope *ScopeError
	if scope, _ = errors.AsType[*ScopeError](err); scope == nil {
		t.Fatalf("Check() error = %v, want ScopeError", err)
	}
}

func TestDiscoverRejectsEscapingSourceSymlink(t *testing.T) {
	root := t.TempDir()
	outside := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	outsidePath := filepath.Join(outside, "Escape.bas")
	if err := os.WriteFile(outsidePath, []byte("Sub Escape()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	linkPath := filepath.Join(moduleDir, "Escape.bas")
	if err := os.Symlink(outsidePath, linkPath); err != nil {
		t.Skipf("symbolic links are unavailable: %v", err)
	}
	_, err := Check(context.Background(), Options{RootDir: root, Roots: Roots{Modules: "src/modules"}})
	var scope *ScopeError
	if scope, _ = errors.AsType[*ScopeError](err); scope == nil {
		t.Fatalf("Check() error = %v, want ScopeError", err)
	}
}

func TestConvertPreservesManagedSourceSymlink(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	targetPath := filepath.Join(moduleDir, "Target.cls")
	linkPath := filepath.Join(moduleDir, "Alias.cls")
	encoded, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte("' 日本語\r\n"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(targetPath, encoded, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(targetPath, linkPath); err != nil {
		t.Skipf("symbolic links are unavailable: %v", err)
	}

	result, err := Convert(context.Background(), Options{
		RootDir: root,
		Roots:   Roots{Modules: "src/modules"},
		Paths:   []string{"src/modules/Alias.cls"},
	}, CP932)
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Files) != 1 || result.Files[0].Path != "src/modules/Alias.cls" || result.Files[0].Status != StatusConverted {
		t.Fatalf("result = %+v", result)
	}
	linkInfo, err := os.Lstat(linkPath)
	if err != nil {
		t.Fatal(err)
	}
	if linkInfo.Mode()&os.ModeSymlink == 0 {
		t.Fatal("conversion replaced the managed source symlink")
	}
	want := "' 日本語\r\n"
	converted, err := os.ReadFile(targetPath)
	if err != nil {
		t.Fatal(err)
	}
	if string(converted) != want {
		t.Fatalf("target = %q, want %q", converted, want)
	}
}
