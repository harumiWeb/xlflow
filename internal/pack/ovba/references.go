package ovba

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"strings"
)

// ProjectReference retains one complete reference, including optional names.
// Raw is independently owned and is the authoritative serialization input.
type ProjectReference struct {
	Raw             []byte
	Name            string
	Kind            uint16
	LibID           string
	OriginalTypeLib [16]byte
}

// ParseProjectReferences reads complete references rather than interpreting
// nested CONTROL fields as independent references.
func ParseProjectReferences(raw []byte, codePage uint16) ([]ProjectReference, error) {
	var result []ProjectReference
	offset := 0
	decode := func(payload []byte) (string, error) {
		value, err := DecodeMBCS(payload, codePage)
		if err != nil {
			return "", err
		}
		if strings.ContainsRune(value, 0) {
			return "", fmt.Errorf("ovba: NUL in reference")
		}
		return value, nil
	}
	read := func() (uint16, []byte, error) {
		if len(raw)-offset < 6 {
			return 0, nil, fmt.Errorf("ovba: truncated reference at %d", offset)
		}
		id := binary.LittleEndian.Uint16(raw[offset:])
		n := uint64(binary.LittleEndian.Uint32(raw[offset+2:]))
		// CONTROL aggregate sizes are ignored on read per MS-OVBA.
		if id == 0x002F || id == 0x0030 || id == 0x000D {
			if len(raw)-offset < 10 {
				return 0, nil, fmt.Errorf("ovba: truncated LIBID length")
			}
			n = 4 + uint64(binary.LittleEndian.Uint32(raw[offset+6:])) + 6
			if id == 0x0030 {
				n += 20
			}
		}
		if n > uint64(len(raw)-offset-6) {
			return 0, nil, fmt.Errorf("ovba: reference 0x%04X exceeds boundary", id)
		}
		payload := raw[offset+6 : offset+6+int(n)]
		offset += 6 + int(n)
		return id, payload, nil
	}
	name := func(ref *ProjectReference) error {
		id, payload, err := read()
		if err != nil {
			return err
		}
		if id != 0x0016 {
			return fmt.Errorf("ovba: expected reference name")
		}
		ref.Name, err = decode(payload)
		if err != nil {
			return err
		}
		if offset+2 <= len(raw) && binary.LittleEndian.Uint16(raw[offset:]) == 0x003E {
			_, payload, err = read()
			if err != nil {
				return err
			}
			unicode, err := utf16leString(payload)
			if err != nil {
				return err
			}
			if unicode != ref.Name {
				return fmt.Errorf("ovba: reference name Unicode mismatch")
			}
		}
		return nil
	}
	for offset < len(raw) {
		start := offset
		ref := ProjectReference{}
		if offset+2 <= len(raw) && binary.LittleEndian.Uint16(raw[offset:]) == 0x0016 {
			if err := name(&ref); err != nil {
				return nil, err
			}
		}
		id, payload, err := read()
		if err != nil {
			return nil, err
		}
		if id == 0x0033 { // REFERENCEORIGINAL must precede CONTROL.
			if _, err := decode(payload); err != nil {
				return nil, err
			}
			id, payload, err = read()
			if err != nil {
				return nil, err
			}
			if id != 0x002F {
				return nil, fmt.Errorf("ovba: ORIGINAL without CONTROL")
			}
		}
		ref.Kind = id
		switch id {
		case 0x002F:
			if _, err := decode(payload[4 : len(payload)-6]); err != nil {
				return nil, err
			}
			if offset+2 <= len(raw) && binary.LittleEndian.Uint16(raw[offset:]) == 0x0016 {
				extended := ProjectReference{}
				if err := name(&extended); err != nil {
					return nil, err
				}
			}
			id, payload, err = read()
			if err != nil {
				return nil, err
			}
			if id != 0x0030 {
				return nil, fmt.Errorf("ovba: CONTROL missing extended record")
			}
			copy(ref.OriginalTypeLib[:], payload[len(payload)-20:len(payload)-4])
			ref.LibID, err = decode(payload[4 : len(payload)-26])
		case 0x000D:
			ref.LibID, err = decode(payload[4 : len(payload)-6])
		case 0x000E:
			// Two sized paths followed by MajorVersion and MinorVersion.
			p := payload
			for range 2 {
				if len(p) < 4 {
					return nil, fmt.Errorf("ovba: truncated PROJECT reference")
				}
				n := uint64(binary.LittleEndian.Uint32(p))
				if n > uint64(len(p)-4) {
					return nil, fmt.Errorf("ovba: PROJECT path exceeds boundary")
				}
				if _, err := decode(p[4 : 4+int(n)]); err != nil {
					return nil, err
				}
				p = p[4+int(n):]
			}
			if len(p) != 6 {
				return nil, fmt.Errorf("ovba: invalid PROJECT reference tail")
			}
		default:
			return nil, fmt.Errorf("ovba: unsupported reference record 0x%04X", id)
		}
		if err != nil {
			return nil, err
		}
		ref.Raw = bytes.Clone(raw[start:offset])
		result = append(result, ref)
	}
	return result, nil
}
