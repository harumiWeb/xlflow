package lspserver

import (
	"encoding/json"
	"path/filepath"
	"strings"

	"gopkg.in/yaml.v3"
)

type DocumentKind int

const (
	DocumentKindUnknown DocumentKind = iota
	DocumentKindVBA
	DocumentKindUserFormYAML
	DocumentKindUserFormJSON
)

// DetectDocumentKind classifies only the configured, direct spec children as
// UserForm candidates. Other YAML and JSON files are deliberately ignored so
// the LSP never claims general-purpose structured-data files.
func DetectDocumentKind(root, configuredFormsRoot, path, source string) DocumentKind {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".yaml" && ext != ".yml" && ext != ".json" {
		return DocumentKindVBA
	}
	if !isUserFormSpecPath(root, configuredFormsRoot, path) {
		return DocumentKindUnknown
	}
	if kind, found := declaredDocumentKind(ext, source); found && kind != "xlflow.userform" {
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

func declaredDocumentKind(ext, source string) (string, bool) {
	switch ext {
	case ".json":
		var root map[string]any
		if err := json.Unmarshal([]byte(source), &root); err != nil {
			return "", false
		}
		value, ok := root["kind"].(string)
		return value, ok
	case ".yaml", ".yml":
		var root yaml.Node
		if err := yaml.Unmarshal([]byte(source), &root); err != nil || len(root.Content) == 0 {
			return "", false
		}
		node := root.Content[0]
		if node.Kind != yaml.MappingNode {
			return "", false
		}
		for i := 0; i+1 < len(node.Content); i += 2 {
			if node.Content[i].Value == "kind" && node.Content[i+1].Kind == yaml.ScalarNode {
				return node.Content[i+1].Value, true
			}
		}
	}
	return "", false
}
