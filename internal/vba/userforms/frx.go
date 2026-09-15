package userforms

import (
	"bytes"
	"encoding/binary"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

var frxCFBSignature = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// FindFRXControlNames returns the candidate names found in UserForm .frx
// streams. The text .frm export may omit nested Begin blocks and keep the
// control metadata in the binary OLE compound file instead. A candidate is
// accepted only when it is the complete name in a serialized control record;
// a control named ComboBoxBitmapVector must not satisfy a ComboBox candidate.
// The boolean is false when the FRX container could not be parsed.
func FindFRXControlNames(data []byte, candidates []string) (map[string]struct{}, bool) {
	needles := make(map[string]struct{}, len(candidates))
	for _, candidate := range candidates {
		name := strings.ToLower(strings.TrimSpace(candidate))
		if name == "" {
			continue
		}
		needles[name] = struct{}{}
	}
	if len(needles) == 0 {
		_, parsed, complete := ExtractFRXControlNames(data)
		return map[string]struct{}{}, parsed && complete
	}

	names, parsed, complete := ExtractFRXControlNames(data)
	if !parsed || !complete {
		return nil, false
	}
	found := make(map[string]struct{}, len(needles))
	for name := range names {
		if _, ok := needles[name]; ok {
			found[name] = struct{}{}
		}
	}
	return found, true
}

// ExtractFRXControlNames returns control names decoded from serialized
// UserForm control records. The second result reports whether the compound
// file opened; the third reports whether every recognized control-record
// marker used a supported name layout.
func ExtractFRXControlNames(data []byte) (map[string]struct{}, bool, bool) {
	container, err := openFRXContainer(data)
	if err != nil {
		return nil, false, false
	}

	names := make(map[string]struct{})
	complete := true
	for _, path := range container.Paths() {
		if !isFRXPropertyStream(path) {
			continue
		}
		stream, ok := container.Stream(path)
		if !ok {
			continue
		}
		streamNames, streamComplete := extractFRXControlNamesFromStream(stream)
		if !streamComplete {
			complete = false
		}
		for _, name := range streamNames {
			names[name] = struct{}{}
		}
	}
	return names, true, complete
}

func isFRXPropertyStream(path string) bool {
	last := path[strings.LastIndexAny(path, "/\\")+1:]
	return strings.EqualFold(last, "f")
}

const (
	// MSForms stores the control name length four bytes into each d5/e5/f5
	// record. The name follows one of these fixed record prefixes depending on
	// whether the control is a form root, a text box, or a nested-page control.
	frxRecordNameLengthOffset  = 4
	frxRecordNameOffset        = 20
	frxTextBoxNameOffset       = 24
	frxNestedControlNameOffset = 28
	frxMaxControlNameLength    = 256
)

func extractFRXControlNamesFromStream(stream []byte) ([]string, bool) {
	if len(stream) < frxRecordNameLengthOffset+4 {
		return nil, true
	}
	seen := make(map[string]struct{})
	complete := true
	for offset := 0; offset+frxRecordNameLengthOffset+4 <= len(stream); offset++ {
		if stream[offset] != 0xd5 && stream[offset] != 0xe5 && stream[offset] != 0xf5 {
			continue
		}
		rawLength := binary.LittleEndian.Uint32(stream[offset+frxRecordNameLengthOffset:])
		if rawLength&0x80000000 == 0 {
			continue
		}
		nameLength := int(rawLength & 0x7fff)
		if nameLength == 0 || nameLength > frxMaxControlNameLength {
			continue
		}
		matched := false
		for _, nameOffset := range []int{frxRecordNameOffset, frxTextBoxNameOffset, frxNestedControlNameOffset} {
			start := offset + nameOffset
			end := start + nameLength
			if end > len(stream) || !isFRXControlName(stream[start:end]) {
				continue
			}
			seen[strings.ToLower(string(stream[start:end]))] = struct{}{}
			matched = true
			break
		}
		if !matched {
			complete = false
		}
	}

	names := make([]string, 0, len(seen))
	for name := range seen {
		names = append(names, name)
	}
	return names, complete
}

func isFRXControlName(value []byte) bool {
	if len(value) == 0 || !isFRXControlNameStart(value[0]) {
		return false
	}
	for _, char := range value[1:] {
		if !isFRXControlNamePart(char) {
			return false
		}
	}
	return true
}

func isFRXControlNameStart(char byte) bool {
	return (char >= 'A' && char <= 'Z') || (char >= 'a' && char <= 'z') || char == '_'
}

func isFRXControlNamePart(char byte) bool {
	return isFRXControlNameStart(char) || (char >= '0' && char <= '9')
}

func openFRXContainer(data []byte) (*cfb.Container, error) {
	// A VBA .frx file has a small LB wrapper before the embedded CFB. The
	// wrapper contains offsets whose size varies between exported artifacts, so
	// locate the signature only in the fixed-size header region.
	header := data
	if len(header) > 64 {
		header = header[:64]
	}
	if offset := bytes.Index(header, frxCFBSignature); offset > 0 {
		data = data[offset:]
	}
	return cfb.Open(data)
}
