// Package events contains the VBE-observed event declarations supported by
// the UserForm Designer. It deliberately stays independent of vbadb: event
// signatures preserve the observed parameter passing syntax.
package events

import (
	_ "embed"
	"encoding/json"
	"fmt"
	"strings"
)

//go:embed catalogue.json
var catalogueJSON []byte

type catalogueEntry struct {
	Type      string `json:"type"`
	Name      string `json:"name"`
	Signature string `json:"signature"`
	IsDefault bool   `json:"isDefault"`
}

// Descriptor is the event shape sent by xlflow/userFormEvents.
type Descriptor struct {
	Name      string `json:"name"`
	Signature string `json:"signature"`
	IsDefault bool   `json:"isDefault"`
}

// Declaration includes the observed host type and is used to validate the
// embedded evidence catalogue.
type Declaration struct {
	Type      string
	Name      string
	Signature string
	IsDefault bool
}

var catalogue = mustLoadCatalogue()

func mustLoadCatalogue() []catalogueEntry {
	var result []catalogueEntry
	if err := json.Unmarshal(catalogueJSON, &result); err != nil {
		panic(fmt.Sprintf("decode embedded UserForm event catalogue: %v", err))
	}
	return result
}

// Declarations returns a copy of all observed event declarations.
func Declarations() []Declaration {
	result := make([]Declaration, 0, len(catalogue))
	for _, entry := range catalogue {
		result = append(result, Declaration(entry))
	}
	return result
}

// SupportedTypes returns the host types represented in the catalogue in
// first-observed order.
func SupportedTypes() []string {
	seen := make(map[string]struct{}, len(catalogue))
	types := make([]string, 0, 15)
	for _, entry := range catalogue {
		key := strings.ToLower(entry.Type)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		types = append(types, entry.Type)
	}
	return types
}

// ForType returns observed event declarations for typeName, rendering the
// object placeholder with objectName. The bool is false for unsupported types.
func ForType(typeName, objectName string) ([]Descriptor, bool) {
	var result []Descriptor
	found := false
	for _, entry := range catalogue {
		if !strings.EqualFold(entry.Type, typeName) {
			continue
		}
		found = true
		result = append(result, Descriptor{
			Name: entry.Name, Signature: strings.Replace(entry.Signature, "{object}", objectName, 1), IsDefault: entry.IsDefault,
		})
	}
	return result, found
}
