package intel

import (
	"encoding/json"
	"fmt"
	"io"
	"slices"
	"strings"
	"unicode/utf16"
	"unicode/utf8"
)

// JSONDocument indexes object fields by the same dotted and indexed paths used
// by FormSpec validation issues. It records source ranges only; validation
// remains the responsibility of the canonical FormSpec parser.
type JSONDocument struct {
	Source     string
	Fields     map[string]JSONFieldNodes
	ParseError error
}

type JSONFieldNodes struct {
	KeyRange   Range
	ValueRange Range
}

func ParseJSON(source string) *JSONDocument {
	doc := &JSONDocument{Source: source, Fields: make(map[string]JSONFieldNodes)}
	fieldOffsets := make(map[string]jsonFieldOffsets)
	decoder := json.NewDecoder(strings.NewReader(source))
	token, err := decoder.Token()
	if err != nil {
		doc.ParseError = err
		return doc
	}
	if _, err := doc.indexValue(decoder, token, "", fieldOffsets); err != nil {
		doc.ParseError = err
		return doc
	}
	if _, err := decoder.Token(); err != io.EOF {
		if err == nil {
			err = fmt.Errorf("multiple JSON values")
		}
		doc.ParseError = err
		return doc
	}
	offsets := make([]int, 0, len(fieldOffsets)*4)
	for _, field := range fieldOffsets {
		offsets = append(offsets, field.keyStart, field.keyEnd, field.valueStart, field.valueEnd)
	}
	positions := jsonPositionsAt(source, offsets)
	for path, field := range fieldOffsets {
		doc.Fields[path] = JSONFieldNodes{
			KeyRange:   Range{Start: positions[field.keyStart], End: positions[field.keyEnd]},
			ValueRange: Range{Start: positions[field.valueStart], End: positions[field.valueEnd]},
		}
	}
	return doc
}

func (d *JSONDocument) Field(path string) (JSONFieldNodes, bool) {
	field, ok := d.Fields[path]
	return field, ok
}

type jsonFieldOffsets struct {
	keyStart   int
	keyEnd     int
	valueStart int
	valueEnd   int
}

func (d *JSONDocument) indexValue(decoder *json.Decoder, token json.Token, path string, fields map[string]jsonFieldOffsets) (int, error) {
	delimiter, ok := token.(json.Delim)
	if !ok {
		return int(decoder.InputOffset()), nil
	}
	switch delimiter {
	case '{':
		for {
			keyStart := skipJSONSeparator(d.Source, int(decoder.InputOffset()), ',')
			keyToken, err := decoder.Token()
			if err != nil {
				return 0, err
			}
			if closing, ok := keyToken.(json.Delim); ok && closing == '}' {
				return int(decoder.InputOffset()), nil
			}
			key, ok := keyToken.(string)
			if !ok {
				return 0, fmt.Errorf("expected JSON object key")
			}
			keyEnd := int(decoder.InputOffset())
			valueStart := skipJSONColon(d.Source, keyEnd)
			valueToken, err := decoder.Token()
			if err != nil {
				return 0, err
			}
			fieldPath := joinPath(path, key)
			valueEnd, err := d.indexValue(decoder, valueToken, fieldPath, fields)
			if err != nil {
				return 0, err
			}
			fields[fieldPath] = jsonFieldOffsets{keyStart: keyStart, keyEnd: keyEnd, valueStart: valueStart, valueEnd: valueEnd}
		}
	case '[':
		index := 0
		for {
			valueToken, err := decoder.Token()
			if err != nil {
				return 0, err
			}
			if closing, ok := valueToken.(json.Delim); ok && closing == ']' {
				return int(decoder.InputOffset()), nil
			}
			itemPath := fmt.Sprintf("%s[%d]", path, index)
			if _, err := d.indexValue(decoder, valueToken, itemPath, fields); err != nil {
				return 0, err
			}
			index++
		}
	default:
		return 0, fmt.Errorf("unexpected JSON delimiter %q", delimiter)
	}
}

func jsonPositionsAt(source string, offsets []int) map[int]Position {
	slices.Sort(offsets)
	offsets = slices.Compact(offsets)
	positions := make(map[int]Position, len(offsets))
	byteOffset, line, character := 0, 0, 0
	for _, target := range offsets {
		for byteOffset < target {
			switch source[byteOffset] {
			case '\r':
				byteOffset++
				if byteOffset < len(source) && source[byteOffset] == '\n' {
					byteOffset++
				}
				line++
				character = 0
			case '\n':
				byteOffset++
				line++
				character = 0
			default:
				r, size := utf8.DecodeRuneInString(source[byteOffset:])
				byteOffset += size
				character += utf16.RuneLen(r)
			}
		}
		positions[target] = Position{Line: line, Character: character}
	}
	return positions
}

func skipJSONWhitespace(source string, offset int) int {
	for offset < len(source) && isJSONWhitespace(source[offset]) {
		offset++
	}
	return offset
}

func skipJSONSeparator(source string, offset int, separator byte) int {
	for offset < len(source) {
		if isJSONWhitespace(source[offset]) || source[offset] == separator {
			offset++
			continue
		}
		break
	}
	return offset
}

func skipJSONColon(source string, offset int) int {
	offset = skipJSONWhitespace(source, offset)
	if offset < len(source) && source[offset] == ':' {
		offset++
	}
	return skipJSONWhitespace(source, offset)
}

func isJSONWhitespace(value byte) bool {
	return value == ' ' || value == '\t' || value == '\r' || value == '\n'
}
