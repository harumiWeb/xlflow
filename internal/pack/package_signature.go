package pack

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"

	"golang.org/x/text/encoding/unicode"
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
	data, readErr := io.ReadAll(reader)
	closeErr := reader.Close()
	if readErr != nil {
		return false, fmt.Errorf("%s: %w", entry.Name, readErr)
	}
	if closeErr != nil {
		return false, fmt.Errorf("%s: %w", entry.Name, closeErr)
	}

	decoder := xml.NewDecoder(bytes.NewReader(xmlMetadataBytes(data)))
	// OPC allows package metadata to be stored as UTF-16. The stream is already
	// normalized to UTF-8, so an `encoding="utf-16"` declaration describes the
	// original storage and the normalized input is used unchanged. Other declared
	// encodings stay fail-closed instead of being guessed.
	decoder.CharsetReader = func(encoding string, input io.Reader) (io.Reader, error) {
		if strings.EqualFold(encoding, "utf-16") {
			return input, nil
		}
		return nil, fmt.Errorf("%s: unsupported xml encoding %q", entry.Name, encoding)
	}
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

// xmlMetadataBytes normalizes OPC XML metadata to UTF-8 for encoding/xml, which
// only accepts UTF-8 input. Detection is bounded to a BOM or the leading `\x00`
// next to `<` — every XML document starts with an element or the `<?xml`
// prolog. All other bytes are parsed as UTF-8, so malformed input still fails
// loudly in the XML decoder.
func xmlMetadataBytes(data []byte) []byte {
	switch {
	case bytes.HasPrefix(data, []byte{0xEF, 0xBB, 0xBF}):
		return data[3:]
	case bytes.HasPrefix(data, []byte{0xFE, 0xFF}):
		return decodeUTF16XML(data, unicode.BigEndian, unicode.ExpectBOM)
	case bytes.HasPrefix(data, []byte{0xFF, 0xFE}):
		return decodeUTF16XML(data, unicode.LittleEndian, unicode.ExpectBOM)
	case len(data) >= 2 && data[0] == '<' && data[1] == 0:
		return decodeUTF16XML(data, unicode.LittleEndian, unicode.IgnoreBOM)
	case len(data) >= 2 && data[0] == 0 && data[1] == '<':
		return decodeUTF16XML(data, unicode.BigEndian, unicode.IgnoreBOM)
	default:
		return data
	}
}

func decodeUTF16XML(data []byte, endianness unicode.Endianness, bomPolicy unicode.BOMPolicy) []byte {
	decoded, err := unicode.UTF16(endianness, bomPolicy).NewDecoder().Bytes(data)
	if err != nil {
		// Malformed UTF-16 still fails loudly when the XML decoder sees the raw bytes.
		return data
	}
	return decoded
}
