package filepull

import (
	"encoding/json"
	"os"
	"path/filepath"
	"slices"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

type pullParityFixture struct {
	SourceCases     []pullParitySourceCase     `json:"sourceCases"`
	FolderCases     []pullParityFolderCase     `json:"folderCases"`
	LineNumberCases []pullParityLineNumberCase `json:"lineNumberCases"`
}

type pullParitySourceCase struct {
	ID       string `json:"id"`
	Kind     string `json:"kind"`
	Name     string `json:"name"`
	GoSource string `json:"goSource"`
	Expected string `json:"expected"`
}

type pullParityFolderCase struct {
	ID               string   `json:"id"`
	Source           string   `json:"source"`
	ExpectedSegments []string `json:"expectedSegments"`
}

type pullParityLineNumberCase struct {
	ID       string `json:"id"`
	Numbered string `json:"numbered"`
	Expected string `json:"expected"`
}

func TestPullParityFixtures(t *testing.T) {
	fixture := readPullParityFixture(t)
	for _, test := range fixture.SourceCases {
		t.Run("source/"+test.ID, func(t *testing.T) {
			kind, ok := map[string]vbaproject.ModuleType{
				"standard": vbaproject.ModuleStd,
				"class":    vbaproject.ModuleClass,
				"document": vbaproject.ModuleDocument,
			}[test.Kind]
			if !ok {
				t.Fatalf("unknown fixture kind %q", test.Kind)
			}
			got, err := vbaproject.ExportModuleSource(vbaproject.Module{Name: test.Name, Type: kind, Source: test.GoSource})
			if err != nil {
				t.Fatal(err)
			}
			if got != test.Expected {
				t.Fatalf("tracked source mismatch\n--- got ---\n%s\n--- want ---\n%s", got, test.Expected)
			}
		})
	}
	for _, test := range fixture.FolderCases {
		t.Run("folder/"+test.ID, func(t *testing.T) {
			got, err := folderSegments(test.Source)
			if err != nil {
				t.Fatal(err)
			}
			if !slices.Equal(got, test.ExpectedSegments) {
				t.Fatalf("folder segments = %v, want %v", got, test.ExpectedSegments)
			}
		})
	}
	for _, test := range fixture.LineNumberCases {
		t.Run("line-numbers/"+test.ID, func(t *testing.T) {
			got, err := removeGeneratedLineNumbers(test.Numbered)
			if err != nil {
				t.Fatal(err)
			}
			if got != test.Expected {
				t.Fatalf("line-number result = %q, want %q", got, test.Expected)
			}
		})
	}
}

func readPullParityFixture(t *testing.T) pullParityFixture {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "..", "testdata", "pull-parity", "cases.json"))
	if err != nil {
		t.Fatal(err)
	}
	var fixture pullParityFixture
	if err := json.Unmarshal(body, &fixture); err != nil {
		t.Fatal(err)
	}
	return fixture
}
