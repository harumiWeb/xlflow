package ovba

import (
	"encoding/binary"
	"errors"
)

// Decompress expands an MS-OVBA CompressedContainer ([MS-OVBA] §2.4.1).
// It is the inverse of Compress.
func Decompress(data []byte) ([]byte, error) {
	if len(data) == 0 {
		return nil, errors.New("ovba: empty input")
	}
	if data[0] != 0x01 {
		return nil, errors.New("ovba: SignatureByte is not 0x01")
	}
	var out []byte
	pos := 1
	for pos < len(data) {
		if pos+2 > len(data) {
			return nil, errors.New("ovba: chunk header truncated")
		}
		hdr := binary.LittleEndian.Uint16(data[pos:])
		pos += 2
		size := int(hdr&0x0FFF) + 3 // total byte count including the header
		compressed := hdr&0x8000 != 0
		body := size - 2
		if pos+body > len(data) {
			return nil, errors.New("ovba: chunk body truncated")
		}
		if !compressed {
			// Per spec a raw chunk has a fixed body=4096. Guard against an out-of-range panic on malformed short input.
			if pos+4096 > len(data) {
				return nil, errors.New("ovba: raw chunk shorter than 4096 bytes")
			}
			out = append(out, data[pos:pos+4096]...)
			pos += 4096
			continue
		}
		decoded, err := decompressChunk(data[pos : pos+body])
		if err != nil {
			return nil, err
		}
		out = append(out, decoded...)
		pos += body
	}
	return out, nil
}

// decompressChunk expands one compressed chunk body. [MS-OVBA] §2.4.1.3.2.
func decompressChunk(body []byte) ([]byte, error) {
	var out []byte
	i := 0
	for i < len(body) {
		flag := body[i]
		i++
		for bit := 0; bit < 8 && i < len(body); bit++ {
			if flag&(1<<uint(bit)) == 0 {
				out = append(out, body[i]) // literal
				i++
				continue
			}
			if i+2 > len(body) {
				return nil, errors.New("ovba: CopyToken truncated")
			}
			token := binary.LittleEndian.Uint16(body[i:])
			i += 2
			bitCount, _ := copyTokenHelp(len(out)) // current position from the start of the chunk
			lengthMask := uint16(0xFFFF) >> uint(bitCount)
			length := int(token&lengthMask) + 3
			offset := int(token>>uint(16-bitCount)) + 1
			start := len(out) - offset
			if start < 0 {
				return nil, errors.New("ovba: CopyToken offset out of range")
			}
			for k := 0; k < length; k++ { // one byte at a time, to support overlapping copies (RLE)
				out = append(out, out[start+k])
			}
		}
	}
	return out, nil
}
