package cli

import (
	"bytes"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/agentskill"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
)

func skillFencedExamples(t *testing.T, reference, language string) []string {
	t.Helper()
	body := strings.ReplaceAll(string(readFileForTest(t, reference)), "\r\n", "\n")
	var examples []string
	for {
		_, rest, found := strings.Cut(body, "```"+language+"\n")
		if !found {
			break
		}
		example, remaining, closed := strings.Cut(rest, "```")
		if !closed {
			t.Fatalf("unclosed %s example in %s", language, reference)
		}
		examples = append(examples, strings.TrimSpace(example)+"\n")
		body = remaining
	}
	if len(examples) == 0 {
		t.Fatalf("no %s example in %s", language, reference)
	}
	return examples
}

func TestBundledSkillFormAndPackExamples(t *testing.T) {
	for _, scenario := range []string{"frame-picture", "multipage-tabs"} {
		t.Run(scenario, func(t *testing.T) {
			dir := t.TempDir()
			if _, err := agentskill.Install(agentskill.InstallOptions{RootDir: dir, Agent: "agents"}); err != nil {
				t.Fatal(err)
			}
			references := filepath.Join(dir, ".agents", "skills", "xlflow", "references")
			formDoc := filepath.Join(references, "forms.md")
			packDoc := filepath.Join(references, "pack.md")
			specs := skillFencedExamples(t, formDoc, "yaml")
			spec := specs[0]
			code := skillFencedExamples(t, formDoc, "vba")[0]
			if scenario == "multipage-tabs" {
				header, _, ok := strings.Cut(spec, "controls:\n")
				if !ok || len(specs) != 2 {
					t.Fatal("container example requires a complete spec header")
				}
				spec = header + specs[1]
				code = "Option Explicit\n"
			}
			writePackSourceModule(t, dir, "xlflow.toml", []byte(skillFencedExamples(t, packDoc, "toml")[0]))
			writePushSourceTree(t, dir)
			writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
			writePackSourceModule(t, dir, "templates/Base.xlsm", readFileForTest(t, filepath.Join(dir, "build", "Book.xlsm")))
			writePackSourceModule(t, dir, "src/forms/specs/Login.yaml", []byte(spec))
			writePackSourceModule(t, dir, "src/forms/code/Login.bas", []byte(code))
			bmp := readFileForTest(t, filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
			writePackSourceModule(t, dir, "src/forms/assets/logo.bmp", bmp)
			stubClosedWorkbook(t)
			verify := func(path string) {
				t.Helper()
				project := readPackedProjectForTest(t, path)
				persistedCode := strings.ReplaceAll(moduleSourceForTest(project, "Login"), "\r\n", "\n")
				if len(project.Forms) != 1 || project.Forms[0].Name != "Login" || !strings.Contains(persistedCode, strings.TrimSpace(code)) {
					t.Fatalf("documented form/code missing from %s", path)
				}
				state, err := projection.Project(project.Forms[0])
				if err != nil {
					t.Fatal(err)
				}
				if scenario == "frame-picture" {
					if len(state.Controls) != 3 || state.Controls[1].ParentID != state.Controls[0].ID {
						t.Fatalf("documented Frame hierarchy not persisted: %+v", state.Controls)
					}
					decoded, err := picture.Decode(project.Forms[0].Controls[1].Record.Pictures["Picture"])
					if err != nil || !bytes.Equal(decoded.Data, bmp) {
						t.Fatalf("documented Image asset not persisted: %v", err)
					}
				} else if len(state.Controls) != 4 || state.Controls[1].ParentID != state.Controls[0].ID || state.Controls[2].ParentID != state.Controls[1].ID || len(state.Controls[3].Tabs) != 1 {
					t.Fatalf("documented Page/TabStrip hierarchy not persisted: %+v", state.Controls)
				}
			}
			for _, command := range skillFencedExamples(t, packDoc, "bash") {
				args := strings.Fields(command)
				if len(args) == 0 || args[0] != "xlflow" {
					t.Fatalf("unexpected documented command: %s", command)
				}
				stdout, err := runBuildCommandForTest(dir, args[1:]...)
				if err != nil {
					t.Fatalf("%s: %v\n%s", command, err, stdout)
				}
				verify(filepath.Join(dir, "dist", "Release.xlsm"))
			}
			stdout, err := runBuildCommandForTest(dir, "push", "--backend", "file", "--json")
			if err != nil {
				t.Fatalf("documented file push: %v\n%s", err, stdout)
			}
			verify(filepath.Join(dir, "build", "Book.xlsm"))
			// Repack the generated form as an existing template, exercising an
			// actual Designer edit rather than only new-form generation.
			writePackSourceModule(t, dir, "templates/Base.xlsm", readFileForTest(t, filepath.Join(dir, "dist", "Release.xlsm")))
			writePackSourceModule(t, dir, "src/forms/specs/Login.yaml", []byte(strings.Replace(spec, "caption: Sign in", "caption: Updated sign in", 1)))
			stdout, err = runBuildCommandForTest(dir, "pack", "--template", "templates/Base.xlsm", "--out", "dist/Edited.xlsm", "--json")
			if err != nil {
				t.Fatalf("documented template edit: %v\n%s", err, stdout)
			}
			verify(filepath.Join(dir, "dist", "Edited.xlsm"))
			state, err := projection.Project(readPackedProjectForTest(t, filepath.Join(dir, "dist", "Edited.xlsm")).Forms[0])
			if err != nil || state.Form.Caption == nil || *state.Form.Caption != "Updated sign in" {
				t.Fatalf("template caption edit missing: %+v, %v", state.Form, err)
			}
		})
	}
}
