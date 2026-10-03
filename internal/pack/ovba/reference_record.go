package ovba

import (
	"encoding/binary"
	"fmt"
	"regexp"
	"strings"
)

// referencePayloadSize ignores aggregate sizes that MS-OVBA marks ignored on
// read, bounding each embedded field against the remaining stream instead.
func referencePayloadSize(raw []byte) (uint64, error) {
	if len(raw) < 6 {
		return 0, fmt.Errorf("ovba: truncated reference header")
	}
	id := binary.LittleEndian.Uint16(raw)
	size := uint64(binary.LittleEndian.Uint32(raw[2:]))
	switch id {
	case 0x000D, 0x002F, 0x0030:
		if len(raw) < 10 {
			return 0, fmt.Errorf("ovba: truncated LIBID length")
		}
		size = 4 + uint64(binary.LittleEndian.Uint32(raw[6:])) + 6
		if id == 0x0030 {
			size += 20
		}
	case 0x000E:
		offset := uint64(6)
		for range 2 {
			if uint64(len(raw))-offset < 4 {
				return 0, fmt.Errorf("ovba: truncated PROJECT path length")
			}
			n := uint64(binary.LittleEndian.Uint32(raw[int(offset):]))
			offset += 4
			if n > uint64(len(raw))-offset {
				return 0, fmt.Errorf("ovba: PROJECT path exceeds boundary")
			}
			offset += n
		}
		size = offset // subtract header and add six-byte version tail
	}
	if size > uint64(len(raw)-6) {
		return 0, fmt.Errorf("ovba: reference 0x%04X exceeds boundary", id)
	}
	return size, nil
}

var libIDPrefix = regexp.MustCompile(`^\*\\[GH]\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$`)

// validateReferenceLIBID checks MS-OVBA LibidReference on encoded bytes, so
// the registration-name limit is a byte limit even for multibyte code pages.
func validateReferenceLIBID(raw []byte) error {
	parts := strings.SplitN(string(raw), "#", 5)
	if len(parts) != 5 || !libIDPrefix.MatchString(parts[0]) || strings.ContainsRune(string(raw), 0) || len(parts[4]) > 255 {
		return fmt.Errorf("ovba: invalid reference LIBID")
	}
	hex := func(value string, limit int) bool {
		if len(value) == 0 || len(value) > limit {
			return false
		}
		for _, b := range []byte(value) {
			switch {
			case b >= '0' && b <= '9', b >= 'a' && b <= 'f', b >= 'A' && b <= 'F':
			default:
				return false
			}
		}
		return true
	}
	major, minor, ok := strings.Cut(parts[1], ".")
	if !ok || !hex(major, 4) || !hex(minor, 4) || !hex(parts[2], 8) {
		return fmt.Errorf("ovba: invalid reference LIBID version or LCID")
	}
	return nil
}
