package ovba

import "encoding/binary"

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
// streamName is the module stream name (it can differ from Name in real bins).
// modType is 0x0021 (procedural/std) or 0x0022 (document/class). [MS-OVBA] §2.3.4.2.3.2.8.
func modRecord(name, streamName string, modType uint16) []byte {
	nu := utf16le(name)
	su := utf16le(streamName)
	var b []byte
	b = append(b, sizedRecord(0x0019, []byte(name))...)       // MODULENAME
	b = append(b, sizedRecord(0x0047, nu)...)                 // MODULENAMEUNICODE
	b = append(b, sizedRecord(0x001A, []byte(streamName))...) // MODULESTREAMNAME
	b = append(b, sizedRecord(0x0032, su)...)                 // ...Unicode
	b = append(b, sizedRecord(0x001C, nil)...)                // MODULEDOCSTRING
	b = append(b, sizedRecord(0x0048, nil)...)                // ...Unicode
	b = append(b, recU32(0x0031, 4, 0x00000000)...)           // MODULEOFFSET = 0 (source-only)
	b = append(b, recU32(0x001E, 4, 0x00000000)...)           // MODULEHELPCONTEXT
	b = append(b, recU16(0x002C, 2, 0xFFFF)...)               // MODULECOOKIE
	b = append(b, sizedRecord(modType, nil)...)               // MODULETYPE
	b = append(b, sizedRecord(0x002B, nil)...)                // MODULE terminator
	return b
}

// ModuleSpec specifies one module used to build the PROJECTMODULES section.
type ModuleSpec struct {
	Name       string
	StreamName string
	TypeID     uint16 // 0x0021=std / 0x0022=class or document
}

// BuildProjectModules builds the PROJECTMODULES section of the dir stream:
// the MODULES count, PROJECTCOOKIE, one record per module (MODULEOFFSET=0),
// and the terminator.
func BuildProjectModules(specs []ModuleSpec) []byte {
	b := recU16(0x000F, 2, uint16(len(specs)))  // MODULES count
	b = append(b, recU16(0x0013, 2, 0xFFFF)...) // PROJECTCOOKIE
	for _, m := range specs {
		b = append(b, modRecord(m.Name, m.StreamName, m.TypeID)...)
	}
	b = append(b, sizedRecord(0x0010, nil)...) // terminator
	return b
}
