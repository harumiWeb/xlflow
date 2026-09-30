package pack

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"
)

const (
	vbaProjectRelationshipsPath = "xl/_rels/vbaProject.bin.rels"
	contentTypesPath            = "[Content_Types].xml"
)

var vbaSignatureRelationshipTypes = map[string]struct{}{
	"http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature": {},
	// Accept the older compatibility spelling as well as the Office 2014 Agile URI.
	"http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile": {},
	"http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile": {},
	"http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3": {},
}

var vbaSignatureContentTypes = map[string]struct{}{
	"application/vnd.ms-office.vbaprojectsignature":      {},
	"application/vnd.ms-office.vbaprojectsignatureagile": {},
	"application/vnd.ms-office.vbaprojectsignaturev3":    {},
}

func hasPackageVBASignature(reader *zip.Reader) (bool, error) {
	var relationshipParts, contentTypeParts []*zip.File
	for _, entry := range reader.File {
		switch entry.Name {
		case vbaProjectRelationshipsPath:
			relationshipParts = append(relationshipParts, entry)
		case contentTypesPath:
			contentTypeParts = append(contentTypeParts, entry)
		}
	}

	var inspectErrors []error
	for _, relationshipPart := range relationshipParts {
		signed, err := zipXMLHasAttributeValue(relationshipPart, "Type", func(name string) bool {
			return name == "Relationship"
		}, func(value string) bool {
			_, ok := vbaSignatureRelationshipTypes[value]
			return ok
		})
		if signed {
			return true, nil
		}
		if err != nil {
			inspectErrors = append(inspectErrors, err)
		}
	}
	for _, contentTypePart := range contentTypeParts {
		signed, err := zipXMLHasAttributeValue(contentTypePart, "ContentType", func(name string) bool {
			return name == "Default" || name == "Override"
		}, func(value string) bool {
			_, ok := vbaSignatureContentTypes[strings.ToLower(strings.TrimSpace(value))]
			return ok
		})
		if signed {
			return true, nil
		}
		if err != nil {
			inspectErrors = append(inspectErrors, err)
		}
	}
	if err := errors.Join(inspectErrors...); err != nil {
		return false, fmt.Errorf("inspect VBA signature package metadata: %w", err)
	}
	return false, nil
}

func zipXMLHasAttributeValue(entry *zip.File, attributeName string, matchesElement, matchesValue func(string) bool) (bool, error) {
	reader, err := entry.Open()
	if err != nil {
		return false, fmt.Errorf("%s: %w", entry.Name, err)
	}
	defer func() { _ = reader.Close() }()

	decoder := xml.NewDecoder(reader)
	for {
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			return false, nil
		}
		if err != nil {
			return false, fmt.Errorf("%s: %w", entry.Name, err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || !matchesElement(start.Name.Local) {
			continue
		}
		for _, attribute := range start.Attr {
			if attribute.Name.Local == attributeName && matchesValue(attribute.Value) {
				return true, nil
			}
		}
	}
}
