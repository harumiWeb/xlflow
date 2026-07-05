package formulas

import (
	"bufio"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
)

func Pull(workbookPath, outputDir string) (Result, error) {
	manifest, names, regionsByPath, err := Extract(workbookPath)
	if err != nil {
		return Result{}, err
	}
	parent := filepath.Dir(outputDir)
	if err := os.MkdirAll(parent, 0o755); err != nil {
		return Result{}, err
	}
	staging, err := os.MkdirTemp(parent, "."+filepath.Base(outputDir)+".tmp-")
	if err != nil {
		return Result{}, err
	}
	defer func() { _ = os.RemoveAll(staging) }()

	for relPath, regions := range regionsByPath {
		if err := writeJSONL(filepath.Join(staging, filepath.FromSlash(relPath)), regions); err != nil {
			return Result{}, err
		}
	}
	if err := writeJSONL(filepath.Join(staging, "names.jsonl"), names); err != nil {
		return Result{}, err
	}
	manifestPath := filepath.Join(outputDir, "manifest.json")
	if err := writeJSONFile(filepath.Join(staging, "manifest.json"), manifest); err != nil {
		return Result{}, err
	}
	if err := replaceGeneratedOutputs(staging, outputDir); err != nil {
		return Result{}, err
	}
	count := 0
	for _, sheet := range manifest.Sheets {
		count += sheet.FormulaRegionCount
	}
	return Result{
		Manifest:           manifest,
		Names:              names,
		OutputDir:          outputDir,
		ManifestPath:       manifestPath,
		FormulaRegionCount: count,
	}, nil
}

func replaceGeneratedOutputs(staging, outputDir string) error {
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return err
	}
	backup, err := os.MkdirTemp(filepath.Dir(outputDir), "."+filepath.Base(outputDir)+".bak-")
	if err != nil {
		return err
	}
	defer func() { _ = os.RemoveAll(backup) }()

	names := []string{"manifest.json", "names.jsonl", "sheets"}
	moved := []string{}
	for _, name := range names {
		target := filepath.Join(outputDir, name)
		if !pathExists(target) {
			continue
		}
		if err := os.Rename(target, filepath.Join(backup, name)); err != nil {
			return err
		}
		moved = append(moved, name)
	}
	rollback := func() {
		for _, name := range names {
			_ = os.RemoveAll(filepath.Join(outputDir, name))
		}
		for _, name := range moved {
			_ = os.Rename(filepath.Join(backup, name), filepath.Join(outputDir, name))
		}
	}
	for _, name := range names {
		target := filepath.Join(outputDir, name)
		if err := os.Rename(filepath.Join(staging, name), target); err != nil {
			rollback()
			return err
		}
	}
	return nil
}

func pathExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}

func writeJSONL[T any](path string, values []T) (err error) {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := file.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	w := bufio.NewWriter(file)
	for _, value := range values {
		line, err := marshalJSON(value, "", "")
		if err != nil {
			return err
		}
		if _, err := w.Write(line); err != nil {
			return err
		}
		if err := w.WriteByte('\n'); err != nil {
			return err
		}
	}
	return w.Flush()
}

func writeJSONFile(path string, value any) error {
	body, err := marshalJSON(value, "", "  ")
	if err != nil {
		return err
	}
	body = append(body, '\n')
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	return os.WriteFile(path, body, 0o644)
}

func marshalJSON(value any, prefix, indent string) ([]byte, error) {
	var b bytes.Buffer
	encoder := json.NewEncoder(&b)
	encoder.SetEscapeHTML(false)
	if indent != "" || prefix != "" {
		encoder.SetIndent(prefix, indent)
	}
	if err := encoder.Encode(value); err != nil {
		return nil, err
	}
	body := b.Bytes()
	return bytes.TrimSuffix(body, []byte{'\n'}), nil
}
