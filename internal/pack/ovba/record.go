package ovba

import "encoding/binary"

// sizedRecord builds a record in the form "2-byte id + 4-byte size + payload".
func sizedRecord(id uint16, payload []byte) []byte {
	out := make([]byte, 6+len(payload))
	binary.LittleEndian.PutUint16(out[0:], id)
	binary.LittleEndian.PutUint32(out[2:], uint32(len(payload)))
	copy(out[6:], payload)
	return out
}

// utf16le converts an ASCII string to a UTF-16LE byte sequence (module names are within ASCII).
func utf16le(s string) []byte {
	out := make([]byte, 0, len(s)*2)
	for _, r := range s {
		out = append(out, byte(r), byte(r>>8))
	}
	return out
}
