package ovba

import (
	"encoding/binary"
	"fmt"
	"math"
	"unicode/utf16"
)

// sizedRecord builds a record in the form "2-byte id + 4-byte size + payload".
func sizedRecord(id uint16, payload []byte) []byte {
	if len(payload) > math.MaxUint32 {
		panic(fmt.Sprintf("ovba: record payload too large for uint32 length: %d", len(payload)))
	}
	if len(payload) > math.MaxInt-6 {
		panic(fmt.Sprintf("ovba: record payload size overflows allocation: %d", len(payload)))
	}
	size := 6 + len(payload)
	out := make([]byte, size)
	binary.LittleEndian.PutUint16(out[0:], id)
	binary.LittleEndian.PutUint32(out[2:], uint32(len(payload)))
	copy(out[6:], payload)
	return out
}

// utf16le converts s to its UTF-16LE byte encoding.
func utf16le(s string) []byte {
	units := utf16.Encode([]rune(s))
	out := make([]byte, 0, len(units)*2)
	for _, u := range units {
		out = append(out, byte(u), byte(u>>8))
	}
	return out
}

// utf16leString decodes a UTF-16LE byte sequence. Odd lengths and unpaired
// surrogates are rejected so a corrupt dir stream cannot be misread.
func utf16leString(b []byte) (string, error) {
	if len(b)%2 != 0 {
		return "", fmt.Errorf("ovba: UTF-16 field has odd length %d", len(b))
	}
	units := make([]uint16, len(b)/2)
	for i := range units {
		units[i] = binary.LittleEndian.Uint16(b[i*2:])
	}
	for i := 0; i < len(units); i++ {
		u := units[i]
		if u >= 0xD800 && u <= 0xDBFF {
			if i+1 >= len(units) || units[i+1] < 0xDC00 || units[i+1] > 0xDFFF {
				return "", fmt.Errorf("ovba: UTF-16 field has an unpaired surrogate at unit %d", i)
			}
			i++ // consume the paired low surrogate
		} else if u >= 0xDC00 && u <= 0xDFFF {
			return "", fmt.Errorf("ovba: UTF-16 field has a lone low surrogate at unit %d", i)
		}
	}
	return string(utf16.Decode(units)), nil
}
