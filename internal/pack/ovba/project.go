package ovba

import (
	"encoding/binary"
	"fmt"
)

// VBAProjectStub returns the fixed 7-byte stub for the _VBA_PROJECT stream.
// It marks the project as source-only (no p-code). The value is verified
// against a known-good binary: cc61ffff000300.
func VBAProjectStub() []byte {
	return []byte{0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00}
}

// recU32 / recU16 build an "id + size + fixed-width integer (LE)" record.
func recU32(id uint16, size uint32, v uint32) []byte {
	p := make([]byte, 4)
	binary.LittleEndian.PutUint32(p, v)
	return sizedRecord(id, p[:size])
}

func recU16(id uint16, size uint32, v uint16) []byte {
	p := make([]byte, 2)
	binary.LittleEndian.PutUint16(p, v)
	return sizedRecord(id, p[:size])
}

// modRecord builds one module entry within dir (MODULENAME..MODULE terminator).
// All MBCS payloads are encoded with the project code page; the paired Unicode
// payloads carry the same text as UTF-16LE, so the two forms are consistent by
// construction. modType is 0x0021 (procedural/std) or 0x0022
// (document/class/form). [MS-OVBA] §2.3.4.2.3.2.
func modRecord(spec ModuleSpec, codepage uint16) ([]byte, error) {
	if spec.TypeID != 0x0021 && spec.TypeID != 0x0022 {
		return nil, fmt.Errorf("ovba: module %q has invalid MODULETYPE 0x%04X", spec.Name, spec.TypeID)
	}
	nameMBCS, err := EncodeMBCS(spec.Name, codepage)
	if err != nil {
		return nil, fmt.Errorf("ovba: module %q name: %w", spec.Name, err)
	}
	streamMBCS, err := EncodeMBCS(spec.StreamName, codepage)
	if err != nil {
		return nil, fmt.Errorf("ovba: module %q stream name: %w", spec.Name, err)
	}
	docMBCS, err := EncodeMBCS(spec.DocString, codepage)
	if err != nil {
		return nil, fmt.Errorf("ovba: module %q doc string: %w", spec.Name, err)
	}
	var b []byte
	b = append(b, sizedRecord(0x0019, nameMBCS)...)                 // MODULENAME
	b = append(b, sizedRecord(0x0047, utf16le(spec.Name))...)       // MODULENAMEUNICODE
	b = append(b, sizedRecord(0x001A, streamMBCS)...)               // MODULESTREAMNAME
	b = append(b, sizedRecord(0x0032, utf16le(spec.StreamName))...) // embedded StreamNameUnicode
	b = append(b, sizedRecord(0x001C, docMBCS)...)                  // MODULEDOCSTRING
	b = append(b, sizedRecord(0x0048, utf16le(spec.DocString))...)  // embedded DocStringUnicode
	b = append(b, recU32(0x0031, 4, 0x00000000)...)                 // MODULEOFFSET = 0 (source-only)
	b = append(b, recU32(0x001E, 4, spec.HelpContext)...)           // MODULEHELPCONTEXT
	b = append(b, recU16(0x002C, 2, 0xFFFF)...)                     // MODULECOOKIE (always 0xFFFF on write)
	b = append(b, sizedRecord(spec.TypeID, nil)...)                 // MODULETYPE
	if spec.ReadOnly {
		b = append(b, sizedRecord(0x0025, nil)...) // MODULEREADONLY
	}
	if spec.Private {
		b = append(b, sizedRecord(0x0028, nil)...) // MODULEPRIVATE
	}
	for _, extra := range spec.Extra {
		switch extra.ID {
		// Record IDs the writer itself emits (or that delimit sections) must
		// not be supplied as opaque extras; parsed Extras can never contain
		// them, but a caller-built ModuleSpec could otherwise emit a second
		// MODULENAME or a stray terminator and corrupt the module record.
		case 0x000F, 0x0010, 0x0013, 0x0019, 0x0047, 0x001A, 0x0032, 0x001C,
			0x0048, 0x0031, 0x001E, 0x002C, 0x0021, 0x0022, 0x0025, 0x0028, 0x002B:
			return nil, fmt.Errorf("ovba: module %q: reserved record 0x%04X cannot be an extra record", spec.Name, extra.ID)
		}
		b = append(b, sizedRecord(extra.ID, extra.Payload)...)
	}
	b = append(b, sizedRecord(0x002B, nil)...) // MODULE terminator
	return b, nil
}

// ModuleSpec specifies one module used to build the PROJECTMODULES section.
type ModuleSpec struct {
	Name        string
	StreamName  string
	TypeID      uint16 // 0x0021=std / 0x0022=class, document, or form
	DocString   string
	HelpContext uint32
	ReadOnly    bool
	Private     bool
	Extra       []ModuleExtraRecord
}

// BuildProjectModules builds the PROJECTMODULES section of the dir stream:
// the MODULES count, PROJECTCOOKIE, one record per module (MODULEOFFSET=0),
// and the terminator.
func BuildProjectModules(specs []ModuleSpec, codepage uint16) ([]byte, error) {
	b := recU16(0x000F, 2, uint16(len(specs)))  // MODULES count
	b = append(b, recU16(0x0013, 2, 0xFFFF)...) // PROJECTCOOKIE
	for _, m := range specs {
		rec, err := modRecord(m, codepage)
		if err != nil {
			return nil, err
		}
		b = append(b, rec...)
	}
	b = append(b, sizedRecord(0x0010, nil)...) // terminator
	return b, nil
}
