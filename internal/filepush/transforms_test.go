package filepush

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestUpdateFolderAnnotationTextModes(t *testing.T) {
	source := "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n"
	annotated := updateFolderAnnotationText(source, "update", "Domain")
	if !strings.Contains(annotated, `'@Folder("Domain")`) {
		t.Fatalf("annotation not inserted: %q", annotated)
	}
	// Attribute headers stay first.
	if idx := strings.Index(annotated, `'@Folder`); idx < strings.Index(annotated, "Attribute VB_Name") {
		t.Fatalf("annotation precedes attribute block: %q", annotated)
	}

	replaced := updateFolderAnnotationText("Attribute VB_Name = \"Module1\"\r\n'@Folder(\"Old\")\r\nOption Explicit\r\n", "update", "New.Path")
	if strings.Count(replaced, "@Folder") != 1 || !strings.Contains(replaced, `'@Folder("New.Path")`) {
		t.Fatalf("annotation not replaced: %q", replaced)
	}

	removed := updateFolderAnnotationText("Attribute VB_Name = \"Module1\"\r\n'@Folder(\"Old\")\r\nOption Explicit\r\n", "update", "")
	if strings.Contains(removed, "@Folder") {
		t.Fatalf("root-level file kept annotation: %q", removed)
	}

	for _, mode := range []string{"ignore", "preserve"} {
		if got := updateFolderAnnotationText(source, mode, "Domain"); got != source {
			t.Fatalf("mode %s changed text: %q", mode, got)
		}
	}
}

func TestFolderAnnotationForPath(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	for _, tc := range []struct {
		path string
		want string
	}{
		{filepath.Join(modules, "A.bas"), ""},
		{filepath.Join(modules, "Domain", "A.bas"), "Domain"},
		{filepath.Join(modules, "Domain", "Sub", "A.bas"), "Domain.Sub"},
		{filepath.Join(modules, "we ird", "A.bas"), "we ird"},
	} {
		if got := folderAnnotationForPath(modules, tc.path); got != tc.want {
			t.Fatalf("annotation(%s) = %q, want %q", tc.path, got, tc.want)
		}
	}
	// A file outside its root can never produce an annotation.
	if got := folderAnnotationForPath(modules, filepath.Join(root, "elsewhere", "A.bas")); got != "" {
		t.Fatalf("outside-root annotation = %q", got)
	}
}

func TestTryAddLineNumbers(t *testing.T) {
	source := "Option Explicit\nPublic Sub Run()\n    Debug.Print \"x\"\nEnd Sub\n"
	out, issue := tryAddLineNumbers(source)
	if issue != nil {
		t.Fatal(issue)
	}
	lines := strings.Split(out, "\n")
	if lines[2] != "3     Debug.Print \"x\"" {
		t.Fatalf("instrumented line = %q", lines[2])
	}
	if lines[0] != "Option Explicit" || lines[1] != "Public Sub Run()" || lines[3] != "End Sub" {
		t.Fatalf("headers/procedure edges were numbered: %q", out)
	}
}

func TestTryAddLineNumbersSkipsDeclarationsLabelsAndContinuations(t *testing.T) {
	source := "Sub Run()\nDim x As Long\nDone:\nx = 1 _\n    + 2\nIf x > 0 Then\n    Debug.Print x\nEnd If\nEnd Sub\n"
	out, issue := tryAddLineNumbers(source)
	if issue != nil {
		t.Fatal(issue)
	}
	lines := strings.Split(out, "\n")
	// The statement's first physical line carries the number; only the
	// continuation tail is skipped.
	if lines[1] != "Dim x As Long" || lines[2] != "Done:" || lines[3] != " 4 x = 1 _" || lines[4] != "    + 2" || lines[5] != "If x > 0 Then" || lines[7] != "End If" {
		t.Fatalf("ineligible lines were numbered: %q", out)
	}
	if lines[6] != " 7     Debug.Print x" {
		t.Fatalf("eligible line = %q", lines[6])
	}
}

func TestTryAddLineNumbersRejectsUnsafeInput(t *testing.T) {
	for name, source := range map[string]string{
		"numeric label":  "Sub Run()\n10  Debug.Print 1\nEnd Sub\n",
		"numeric goto":   "Sub Run()\nGoTo 10\nEnd Sub\n",
		"numeric resume": "Sub Run()\nResume 5\nEnd Sub\n",
	} {
		if _, issue := tryAddLineNumbers(source); issue == nil {
			t.Fatalf("%s: expected safety issue", name)
		}
	}
	// Numeric targets inside strings/comments are inert.
	if _, issue := tryAddLineNumbers("Sub Run()\nDebug.Print \"GoTo 10\"\n' Resume 5\nEnd Sub\n"); issue != nil {
		t.Fatalf("quoted/commented numeric target rejected: %+v", issue)
	}
}

func TestFingerprintEqualsIsOrderInsensitive(t *testing.T) {
	left := sourceFingerprint{
		WorkbookPath: "C:/book/Book.xlsm",
		Files: []sourceFileEntry{
			{Kind: "module", Path: "A.bas", Hash: "1"},
			{Kind: "class", Path: "B.cls", Hash: "2"},
		},
	}
	right := sourceFingerprint{
		WorkbookPath: "C:/book/Book.xlsm",
		Files: []sourceFileEntry{
			{Kind: "class", Path: "B.cls", Hash: "2"},
			{Kind: "module", Path: "A.bas", Hash: "1"},
		},
	}
	if !fingerprintEquals(left, right) {
		t.Fatal("equivalent fingerprints rejected")
	}
	right.LineNumbersEnabled = true
	if fingerprintEquals(left, right) {
		t.Fatal("line-number toggle ignored")
	}
	right.LineNumbersEnabled = false
	right.Files[0].Hash = "changed"
	if fingerprintEquals(left, right) {
		t.Fatal("hash drift ignored")
	}
}

func TestReadPushStateAcceptsLegacyBareFingerprint(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "push.json")
	legacy := `{"workbook_path":"C:/book/Book.xlsm","files":[],"line_numbers_enabled":false}`
	if err := os.WriteFile(path, []byte(legacy), 0o644); err != nil {
		t.Fatal(err)
	}
	state, err := readPushState(path)
	if err != nil || state == nil {
		t.Fatalf("readPushState = %+v, %v", state, err)
	}
	if state.AppliedTo != nil {
		t.Fatal("legacy state must not claim delivery evidence")
	}
	if shouldSkipUnchanged(path, sourceFingerprint{WorkbookPath: "C:/book/Book.xlsm"}, "whatever") {
		t.Fatal("legacy state justified a skip")
	}
}

func TestDiscoverSourceFilesKindsAndOrdering(t *testing.T) {
	root := newSourceTree(t)
	writeTestFile(t, filepath.Join(root, "src", "modules", "B.bas"), "x")
	writeTestFile(t, filepath.Join(root, "src", "modules", "A.bas"), "x")
	writeTestFile(t, filepath.Join(root, "src", "classes", "C.cls"), "x")
	writeTestFile(t, filepath.Join(root, "src", "forms", "F.frm"), "x")
	writeTestFile(t, filepath.Join(root, "src", "forms", "F.frx"), "x")
	writeTestFile(t, filepath.Join(root, "src", "forms", "code", "F.bas"), "x")
	writeTestFile(t, filepath.Join(root, "src", "workbook", "ThisWorkbook.bas"), "x")

	cfg := testConfig()
	files, err := discoverSourceFiles(resolvedRoots(root, cfg), "sidecar")
	if err != nil {
		t.Fatal(err)
	}
	var kinds []string
	for _, f := range files {
		kinds = append(kinds, f.Kind+":"+f.RelativePath)
	}
	want := []string{
		"module:A.bas", "module:B.bas",
		"class:C.cls",
		"form:F.frm", "form:F.frx",
		"document:ThisWorkbook.bas",
		"form_code:F.bas",
	}
	if strings.Join(kinds, ",") != strings.Join(want, ",") {
		t.Fatalf("discovery = %v, want %v", kinds, want)
	}

	// Non-sidecar mode fingerprints code/*.bas as a plain form file set entry.
	files, err = discoverSourceFiles(resolvedRoots(root, cfg), "embedded")
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range files {
		if f.Kind == "form_code" {
			t.Fatalf("non-sidecar discovery produced form_code: %+v", f)
		}
	}
}
