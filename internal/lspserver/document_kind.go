package lspserver

import (
	"path/filepath"
	"strings"
)

type DocumentKind int

const (
	DocumentKindUnknown DocumentKind = iota
	DocumentKindVBA
	DocumentKindUserFormYAML
	DocumentKindUserFormJSON
)

// DetectDocumentKind classifies only the configured, direct spec children as
// UserForm candidates. This path-based recognition intentionally includes
// malformed and wrong-kind specs so their kind or syntax can be diagnosed;
// other YAML and JSON files remain ignored.
func DetectDocumentKind(root, configuredFormsRoot, path, _ string) DocumentKind {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".yaml" && ext != ".yml" && ext != ".json" {
		return DocumentKindVBA
	}
	if !isUserFormSpecPath(root, configuredFormsRoot, path) {
		return DocumentKindUnknown
	}
	if ext == ".json" {
		return DocumentKindUserFormJSON
	}
	return DocumentKindUserFormYAML
}

func isUserFormSpecPath(root, configuredFormsRoot, path string) bool {
	if strings.TrimSpace(root) == "" {
		root = "."
	}
	formsRoot := strings.TrimSpace(configuredFormsRoot)
	if formsRoot == "" {
		formsRoot = filepath.Join("src", "forms")
	}
	if !filepath.IsAbs(formsRoot) {
		formsRoot = filepath.Join(root, formsRoot)
	}
	specsDir, err := filepath.Abs(filepath.Join(formsRoot, "specs"))
	if err != nil {
		return false
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(root, path)
	}
	path, err = filepath.Abs(path)
	if err != nil {
		return false
	}
	return strings.EqualFold(filepath.Clean(filepath.Dir(path)), filepath.Clean(specsDir))
}
