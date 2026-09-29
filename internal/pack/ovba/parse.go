package ovba

import (
	"encoding/binary"
	"fmt"
	"strings"
)

type record struct {
	id      uint16
	start   int // byte position of the record's start within buf (used to slice out the references-section span)
	payload []byte
}

// walkRecords breaks dir.plain into a sequence of (id, start, payload).
// PROJECTVERSION (0x0009) is special-cased because it has no size field and is a fixed 12 bytes.
func walkRecords(buf []byte) ([]record, error) {
	var recs []record
	i := 0
	for i < len(buf) {
		if len(buf)-i < 6 {
			return nil, fmt.Errorf("ovba: record header truncated at offset %d", i)
		}
		id := binary.LittleEndian.Uint16(buf[i:])
		if id == 0x0009 { // id(2)+Reserved(4)+Major(4)+Minor(2)
			if len(buf)-i < 12 {
				return nil, fmt.Errorf("ovba: PROJECTVERSION record truncated at offset %d", i)
			}
			i += 12
			continue
		}
		size := uint64(binary.LittleEndian.Uint32(buf[i+2:]))
		end := uint64(i) + 6 + size
		if end > uint64(len(buf)) {
			return nil, fmt.Errorf("ovba: record 0x%04X at offset %d declares %d bytes beyond the stream", id, i, size)
		}
		endInt := int(end)
		recs = append(recs, record{id: id, start: i, payload: buf[i+6 : endInt]})
		i = endInt
	}
	return recs, nil
}

// DirModule holds the metadata of one module found in the dir stream.
type DirModule struct {
	Name        string
	StreamName  string
	DocString   string
	HelpContext uint32
	Offset      uint32 // MODULEOFFSET (start of the source within the module stream)
	TypeID      uint16 // 0x0021=procedural / 0x0022=non-procedural
	ReadOnly    bool   // MODULEREADONLY (0x0025) present
	Private     bool   // MODULEPRIVATE (0x0028) present
	// Extra holds module-level records this model does not interpret, in
	// original order. The writer re-emits them before the module terminator
	// so unknown records are never silently stripped.
	Extra []ModuleExtraRecord
}

// ModuleExtraRecord is one uninterpreted record inside a MODULE Record: the
// record ID plus its raw payload bytes.
type ModuleExtraRecord struct {
	ID      uint16
	Payload []byte
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
// MBCS module fields are decoded with PROJECTCODEPAGE and the MODULE record
// shape is validated per MS-OVBA; malformed or unrepresentable input is
// rejected rather than silently misread.
func ParseDir(plain []byte) (DirInfo, error) {
	var di DirInfo
	recs, err := walkRecords(plain)
	if err != nil {
		return di, err
	}
	// References section = from the first REFERENCENAME(0x0016) up to just before the first PROJECTMODULES(0x000F).
	// The internal structure (nested REFERENCECONTROL, etc.) is not interpreted; the byte span is preserved verbatim.
	// 0x000F is a top-level marker that never appears in a reference sub-record, so it is safe as a terminator.
	refStart, refEnd := -1, -1
	moduleStart := -1
	var refNameRaw [][]byte
	for _, r := range recs {
		switch r.id {
		case 0x0001:
			if err := requirePayloadSize(r, 4); err != nil {
				return di, err
			}
			di.SysKind = le32(r.payload)
		case 0x0014:
			if err := requirePayloadSize(r, 4); err != nil {
				return di, err
			}
			di.LCID = le32(r.payload)
		case 0x0003:
			if err := requirePayloadSize(r, 2); err != nil {
				return di, err
			}
			di.CodePage = le16(r.payload)
		case 0x0016: // REFERENCENAME (duplicate names inside REFERENCECONTROL are folded by dedup)
			if refStart < 0 {
				refStart = r.start
			}
			refNameRaw = append(refNameRaw, r.payload)
		case 0x000F: // PROJECTMODULES -> end of the references section
			if refEnd < 0 {
				refEnd = r.start
				moduleStart = r.start
			}
		case 0x0019: // MODULENAME is only valid inside PROJECTMODULES
			if moduleStart < 0 {
				return di, fmt.Errorf("ovba: MODULENAME record before PROJECTMODULES at offset %d", r.start)
			}
		}
	}
	if di.CodePage == 0 {
		return di, fmt.Errorf("ovba: dir stream has no PROJECTCODEPAGE record")
	}
	// Resolve the codec before decoding any MBCS field; unsupported code pages
	// fail deterministically here instead of producing mojibake later.
	if di.CodePage != 65001 {
		if _, err := codePageEncoding(di.CodePage); err != nil {
			return di, err
		}
	}
	seen := map[string]bool{}
	for _, payload := range refNameRaw {
		name, err := DecodeMBCS(payload, di.CodePage)
		if err != nil {
			return di, fmt.Errorf("ovba: REFERENCENAME decode: %w", err)
		}
		if !seen[name] {
			seen[name] = true
			di.RefNames = append(di.RefNames, name)
		}
	}
	if moduleStart >= 0 {
		if err := di.parseModules(recs, moduleStart); err != nil {
			return di, err
		}
	}
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
	return di, nil
}

// moduleFieldRank orders the known records of a MODULE Record per MS-OVBA
// §2.3.4.2.3.2; 0 marks records that are either embedded payloads (0x0032,
// 0x0048) validated by position, or unknown records preserved verbatim.
func moduleFieldRank(id uint16) int {
	switch id {
	case 0x0019:
		return 1
	case 0x0047:
		return 2
	case 0x001A:
		return 3
	case 0x001C:
		return 4
	case 0x0031:
		return 5
	case 0x001E:
		return 6
	case 0x002C:
		return 7
	case 0x0021, 0x0022:
		return 8
	case 0x0025:
		return 9
	case 0x0028:
		return 10
	}
	return 0
}

// parseModules walks the PROJECTMODULES section starting after the 0x000F
// record at sectionStart, decoding each MODULE Record until 0x0010.
func (di *DirInfo) parseModules(recs []record, sectionStart int) error {
	var cur *DirModule
	lastRank := 0
	lastID := uint16(0)
	seenFields := map[uint16]bool{}

	finalize := func() error {
		for _, id := range []uint16{0x0019, 0x001A, 0x001C, 0x0031, 0x001E, 0x002C} {
			if !seenFields[id] {
				return fmt.Errorf("ovba: module %q is missing required record 0x%04X", cur.Name, id)
			}
		}
		if cur.TypeID != 0x0021 && cur.TypeID != 0x0022 {
			return fmt.Errorf("ovba: module %q has invalid MODULETYPE 0x%04X", cur.Name, cur.TypeID)
		}
		di.Modules = append(di.Modules, *cur)
		cur = nil
		lastRank = 0
		lastID = 0
		seenFields = map[uint16]bool{}
		return nil
	}

	inSection := false
	for _, r := range recs {
		if !inSection {
			if r.start == sectionStart {
				inSection = true
			}
			continue
		}
		if r.id == 0x0010 { // PROJECTMODULES terminator
			if cur != nil {
				return fmt.Errorf("ovba: module %q has no terminator before PROJECTMODULES end", cur.Name)
			}
			return nil
		}
		if cur == nil {
			switch r.id {
			case 0x0013: // PROJECTCOOKIE
				continue
			case 0x0019: // MODULENAME: start a module
				name, err := decodeModuleText(r.payload, di.CodePage, "MODULENAME")
				if err != nil {
					return err
				}
				cur = &DirModule{Name: name}
				seenFields[r.id] = true
				lastRank = moduleFieldRank(r.id)
				lastID = r.id
				continue
			default:
				return fmt.Errorf("ovba: record 0x%04X in PROJECTMODULES outside a module at offset %d", r.id, r.start)
			}
		}
		rank := moduleFieldRank(r.id)
		switch {
		case r.id == 0x002B: // module terminator
			if err := finalize(); err != nil {
				return err
			}
			continue
		case r.id == 0x0032: // embedded StreamNameUnicode: must follow MODULESTREAMNAME
			if lastID != 0x001A {
				return fmt.Errorf("ovba: module %q: MODULESTREAMNAMEUNICODE out of position (after 0x%04X)", cur.Name, lastID)
			}
			uni, err := utf16leString(r.payload)
			if err != nil {
				return fmt.Errorf("ovba: module %q stream name unicode: %w", cur.Name, err)
			}
			if uni != cur.StreamName {
				return fmt.Errorf("ovba: module %q stream name mismatch: MBCS %q vs Unicode %q", cur.Name, cur.StreamName, uni)
			}
		case r.id == 0x0048: // embedded DocStringUnicode: must follow MODULEDOCSTRING
			if lastID != 0x001C {
				return fmt.Errorf("ovba: module %q: MODULEDOCSTRINGUNICODE out of position (after 0x%04X)", cur.Name, lastID)
			}
			uni, err := utf16leString(r.payload)
			if err != nil {
				return fmt.Errorf("ovba: module %q doc string unicode: %w", cur.Name, err)
			}
			if uni != cur.DocString {
				return fmt.Errorf("ovba: module %q doc string mismatch: MBCS %q vs Unicode %q", cur.Name, cur.DocString, uni)
			}
		case rank == 0:
			cur.Extra = append(cur.Extra, ModuleExtraRecord{ID: r.id, Payload: append([]byte(nil), r.payload...)})
		case rank <= lastRank:
			return fmt.Errorf("ovba: module %q: record 0x%04X out of order (after 0x%04X)", cur.Name, r.id, lastID)
		default:
			seenFields[r.id] = true
			if err := applyModuleField(cur, r, di.CodePage); err != nil {
				return err
			}
		}
		lastID = r.id
		if rank > 0 {
			lastRank = rank
		}
	}
	return fmt.Errorf("ovba: PROJECTMODULES section has no terminator")
}

// applyModuleField decodes one known-position module record into cur.
func applyModuleField(cur *DirModule, r record, codepage uint16) error {
	switch r.id {
	case 0x0047: // MODULENAMEUNICODE
		uni, err := utf16leString(r.payload)
		if err != nil {
			return fmt.Errorf("ovba: module %q name unicode: %w", cur.Name, err)
		}
		if uni != cur.Name {
			return fmt.Errorf("ovba: module name mismatch: MBCS %q vs Unicode %q", cur.Name, uni)
		}
	case 0x001A:
		name, err := decodeModuleText(r.payload, codepage, "MODULESTREAMNAME")
		if err != nil {
			return fmt.Errorf("ovba: module %q: %w", cur.Name, err)
		}
		cur.StreamName = name
	case 0x001C:
		doc, err := decodeModuleText(r.payload, codepage, "MODULEDOCSTRING")
		if err != nil {
			return fmt.Errorf("ovba: module %q: %w", cur.Name, err)
		}
		cur.DocString = doc
	case 0x0031:
		if err := requirePayloadSize(r, 4); err != nil {
			return err
		}
		cur.Offset = le32(r.payload)
	case 0x001E:
		if err := requirePayloadSize(r, 4); err != nil {
			return err
		}
		cur.HelpContext = le32(r.payload)
	case 0x002C:
		if err := requirePayloadSize(r, 2); err != nil {
			return err
		}
		// MODULECOOKIE is ignored on read; the writer always emits 0xFFFF.
	case 0x0021, 0x0022:
		if err := requirePayloadSize(r, 0); err != nil {
			return err
		}
		cur.TypeID = r.id
	case 0x0025:
		if err := requirePayloadSize(r, 0); err != nil {
			return err
		}
		cur.ReadOnly = true
	case 0x0028:
		if err := requirePayloadSize(r, 0); err != nil {
			return err
		}
		cur.Private = true
	}
	return nil
}

func requirePayloadSize(r record, want int) error {
	if len(r.payload) != want {
		return fmt.Errorf("ovba: record 0x%04X at offset %d has payload size %d (want %d)", r.id, r.start, len(r.payload), want)
	}
	return nil
}

// decodeModuleText decodes an MBCS text field and rejects embedded NULs.
func decodeModuleText(b []byte, codepage uint16, field string) (string, error) {
	s, err := DecodeMBCS(b, codepage)
	if err != nil {
		return "", fmt.Errorf("ovba: %s: %w", field, err)
	}
	if strings.IndexByte(s, 0) >= 0 {
		return "", fmt.Errorf("ovba: %s contains a NUL character", field)
	}
	return s, nil
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
	Components   []ProjectComponent
}

// ProjectComponent is one component declaration from the textual PROJECT stream.
type ProjectComponent struct {
	Kind string
	Name string
}

// ParseProjectText parses the PROJECT stream line by line. The stream is MBCS
// text in the project code page; component names and property values are
// decoded with that code page and surrounding "..." quotes are stripped.
// Undecodable component declarations return an error.
func ParseProjectText(raw []byte, codepage uint16) (ProjectText, error) {
	pt := ProjectText{Kinds: map[string]string{}}
	unq := func(s string) string { return strings.Trim(strings.TrimSpace(s), "\"") }
	for line := range strings.SplitSeq(strings.ReplaceAll(string(raw), "\r\n", "\n"), "\n") {
		eq := strings.IndexByte(line, '=')
		if eq < 0 {
			continue
		}
		key, val := line[:eq], line[eq+1:]
		switch key {
		case "ID":
			pt.ID = unq(val)
		case "Name":
			decoded, err := DecodeMBCS([]byte(unq(val)), codepage)
			if err != nil {
				return pt, fmt.Errorf("ovba: PROJECT Name decode: %w", err)
			}
			pt.Name = decoded
		case "CMG":
			pt.CMG = unq(val)
		case "DPB":
			pt.DPB = unq(val)
		case "GC":
			pt.GC = unq(val)
		case "Module", "Class", "BaseClass":
			name, err := DecodeMBCS([]byte(unq(val)), codepage)
			if err != nil {
				return pt, fmt.Errorf("ovba: PROJECT component %s decode: %w", key, err)
			}
			pt.Kinds[name] = key
			pt.Components = append(pt.Components, ProjectComponent{Kind: key, Name: name})
		case "Document":
			name := val
			if i := strings.IndexByte(val, '/'); i >= 0 { // "Sheet1/&H00000000"
				name = val[:i]
			}
			decoded, err := DecodeMBCS([]byte(unq(name)), codepage)
			if err != nil {
				return pt, fmt.Errorf("ovba: PROJECT component %s decode: %w", key, err)
			}
			pt.Kinds[decoded] = "Document"
			pt.Components = append(pt.Components, ProjectComponent{Kind: key, Name: decoded})
		}
	}
	return pt, nil
}
