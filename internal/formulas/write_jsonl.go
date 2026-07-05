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
	if err := os.RemoveAll(outputDir); err != nil {
		return Result{}, err
	}
	if err := os.MkdirAll(filepath.Join(outputDir, "sheets"), 0o755); err != nil {
		return Result{}, err
	}
	for relPath, regions := range regionsByPath {
		if err := writeJSONL(filepath.Join(outputDir, filepath.FromSlash(relPath)), regions); err != nil {
			return Result{}, err
		}
	}
	if err := writeJSONL(filepath.Join(outputDir, "names.jsonl"), names); err != nil {
		return Result{}, err
	}
	manifestPath := filepath.Join(outputDir, "manifest.json")
	if err := writeJSONFile(manifestPath, manifest); err != nil {
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
