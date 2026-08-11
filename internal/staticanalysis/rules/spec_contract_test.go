package rules

import (
	"bytes"
	"encoding/json"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"testing"
)

var specContractPattern = regexp.MustCompile(`<!-- xlflow-rule-contract: (\{[^\n]+\}) -->`)

var requiredSpecContracts = map[string]string{
	"VBA225": "excel-loop-performance-analysis.md",
	"VBA238": "excel-loop-performance-analysis.md",
	"VBA242": "excel-loop-performance-analysis.md",
	"VBA243": "excel-loop-performance-analysis.md",
	"VBA244": "vba-procedure-call-cycles.md",
}

type specRuleContract struct {
	ID                 string       `json:"id"`
	Family             RuleFamily   `json:"family"`
	Category           RuleCategory `json:"category"`
	DefaultSeverity    RuleSeverity `json:"default_severity"`
	Scope              RuleScope    `json:"scope"`
	Realtime           bool         `json:"realtime"`
	ConfigurationKey   string       `json:"configuration_key"`
	InlineSuppressible bool         `json:"inline_suppressible"`
	PreflightBlocking  bool         `json:"preflight_blocking"`
}

func TestSpecificationRuleContractsMatchRegistry(t *testing.T) {
	root := filepath.Join("..", "..", "..", "docs", "specs")
	seen := make(map[string]string)
	markers := 0
	err := filepath.WalkDir(root, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil || entry.IsDir() || filepath.Ext(path) != ".md" {
			return walkErr
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		for _, match := range specContractPattern.FindAllSubmatch(body, -1) {
			markers++
			var contract specRuleContract
			if !json.Valid(match[1]) {
				t.Errorf("%s: rule contract is not one complete JSON value", path)
				continue
			}
			decoder := json.NewDecoder(bytes.NewReader(match[1]))
			decoder.DisallowUnknownFields()
			if err := decoder.Decode(&contract); err != nil {
				t.Errorf("%s: decode rule contract: %v", path, err)
				continue
			}
			if previous, ok := seen[contract.ID]; ok {
				t.Errorf("%s: duplicate contract for %s (already in %s)", path, contract.ID, previous)
				continue
			}
			seen[contract.ID] = path
			metadata, ok := Lookup(contract.ID)
			if !ok {
				t.Errorf("%s: unknown rule %q", path, contract.ID)
				continue
			}
			want := specRuleContract{ID: metadata.ID, Family: metadata.Family, Category: metadata.Category, DefaultSeverity: metadata.DefaultSeverity, Scope: metadata.Scope, Realtime: metadata.Realtime, ConfigurationKey: metadata.ConfigurationKey, InlineSuppressible: metadata.InlineSuppressible, PreflightBlocking: metadata.PreflightBlocking}
			if contract != want {
				t.Errorf("%s: contract = %#v, registry = %#v", path, contract, want)
			}
		}
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	if markers == 0 {
		t.Fatal("no xlflow-rule-contract markers found")
	}
	for id, filename := range requiredSpecContracts {
		path, ok := seen[id]
		if !ok {
			t.Errorf("required specification contract %s is missing from %s", id, filename)
			continue
		}
		if filepath.Base(path) != filename {
			t.Errorf("specification contract %s is in %s, want %s", id, path, filename)
		}
	}
}
