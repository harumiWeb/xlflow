package ovba

import (
	"encoding/binary"
	"strings"
)

type record struct {
	id      uint16
	start   int // byte position of the record's start within buf (used to slice out the references-section span)
	payload []byte
}

// walkRecords breaks dir.plain into a sequence of (id, start, payload).
// PROJECTVERSION (0x0009) is special-cased because it has no size field and is a fixed 12 bytes.
func walkRecords(buf []byte) []record {
	var recs []record
	i := 0
	for i+6 <= len(buf) {
		id := binary.LittleEndian.Uint16(buf[i:])
		if id == 0x0009 { // id(2)+Reserved(4)+Major(4)+Minor(2)
			i += 12
			continue
		}
		size := int(binary.LittleEndian.Uint32(buf[i+2:]))
		if i+6+size > len(buf) {
			break
		}
		recs = append(recs, record{id: id, start: i, payload: buf[i+6 : i+6+size]})
		i += 6 + size
	}
	return recs
}

// DirModule holds the metadata of one module found in the dir stream.
type DirModule struct {
	Name       string
	StreamName string
	Offset     uint32 // MODULEOFFSET (start of the source within the module stream)
	TypeID     uint16 // 0x0021=procedural / 0x0022=non-procedural
}

// DirInfo holds the metadata extracted from the decompressed dir stream.
type DirInfo struct {
	SysKind        uint32
	LCID           uint32
	CodePage       uint16
	RefNames       []string // reference names (for display, deduplicated)
	RefsRaw        []byte   // verbatim byte span of the references section (first 0x0016 up to 0x000F)
	ProjectInfoRaw []byte   // verbatim span of PROJECTINFORMATION (start up to PROJECTREFERENCES)
	Modules        []DirModule
}

// ParseDir extracts metadata from the decompressed dir stream (dir.plain).
func ParseDir(plain []byte) DirInfo {
	var di DirInfo
	recs := walkRecords(plain)
	var cur *DirModule
	flush := func() {
		if cur != nil {
			di.Modules = append(di.Modules, *cur)
			cur = nil
		}
	}
	// References section = from the first REFERENCENAME(0x0016) up to just before the first PROJECTMODULES(0x000F).
	// The internal structure (nested REFERENCECONTROL, etc.) is not interpreted; the byte span is preserved verbatim.
	// 0x000F is a top-level marker that never appears in a reference sub-record, so it is safe as a terminator.
	refStart, refEnd := -1, -1
	seen := map[string]bool{}
	for _, r := range recs {
		switch r.id {
		case 0x0001:
			di.SysKind = le32(r.payload)
		case 0x0014:
			di.LCID = le32(r.payload)
		case 0x0003:
			di.CodePage = le16(r.payload)
		case 0x0016: // REFERENCENAME (duplicate names inside REFERENCECONTROL are folded by dedup)
			if refStart < 0 {
				refStart = r.start
			}
			if name := string(r.payload); !seen[name] {
				seen[name] = true
				di.RefNames = append(di.RefNames, name)
			}
		case 0x000F: // PROJECTMODULES -> end of the references section
			if refEnd < 0 {
				refEnd = r.start
			}
		case 0x0019: // MODULENAME -> new module
			flush()
			cur = &DirModule{Name: string(r.payload)}
		case 0x001A:
			if cur != nil {
				cur.StreamName = string(r.payload)
			}
		case 0x0031:
			if cur != nil {
				cur.Offset = le32(r.payload)
			}
		case 0x0021, 0x0022:
			if cur != nil {
				cur.TypeID = r.id
			}
		}
	}
	flush()
	// PROJECTINFORMATION runs from the start to just before PROJECTREFERENCES; with no references, up to just before PROJECTMODULES.
	infoEnd := refStart
	if infoEnd < 0 {
		infoEnd = refEnd // no references: the information section ends at PROJECTMODULES (0x000F)
	}
	if infoEnd > 0 {
		di.ProjectInfoRaw = plain[:infoEnd]
	}
	if refStart >= 0 && refEnd > refStart {
		di.RefsRaw = plain[refStart:refEnd]
	}
	return di
}

func le16(b []byte) uint16 {
	if len(b) < 2 {
		return 0
	}
	return binary.LittleEndian.Uint16(b)
}

func le32(b []byte) uint32 {
	if len(b) < 4 {
		return 0
	}
	return binary.LittleEndian.Uint32(b)
}

// ProjectText holds the information extracted from the (textual) PROJECT stream.
type ProjectText struct {
	ID, Name     string
	CMG, DPB, GC string
	Kinds        map[string]string // module name → "Module"/"Class"/"Document"/"BaseClass"
}

// ParseProjectText parses the PROJECT stream line by line. Surrounding "..." quotes on values are stripped.
func ParseProjectText(raw []byte) ProjectText {
	pt := ProjectText{Kinds: map[string]string{}}
	unq := func(s string) string { return strings.Trim(strings.TrimSpace(s), "\"") }
	for _, line := range strings.Split(string(raw), "\r\n") {
		eq := strings.IndexByte(line, '=')
		if eq < 0 {
			continue
		}
		key, val := line[:eq], line[eq+1:]
		switch key {
		case "ID":
			pt.ID = unq(val)
		case "Name":
			pt.Name = unq(val)
		case "CMG":
			pt.CMG = unq(val)
		case "DPB":
			pt.DPB = unq(val)
		case "GC":
			pt.GC = unq(val)
		case "Module", "Class", "BaseClass":
			pt.Kinds[unq(val)] = key
		case "Document":
			name := val
			if i := strings.IndexByte(val, '/'); i >= 0 { // "Sheet1/&H00000000"
				name = val[:i]
			}
			pt.Kinds[unq(name)] = "Document"
		}
	}
	return pt
}
