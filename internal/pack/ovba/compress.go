// Package ovba implements the MS-OVBA content layer (compression and
// dir/PROJECT assembly). It knows nothing about the CFB envelope.
package ovba

// Compress produces an MS-OVBA CompressedContainer ([MS-OVBA] §2.4.1):
// a leading SignatureByte 0x01 followed by a sequence of CompressedChunks,
// one per 4096-byte block of input.
func Compress(data []byte) []byte {
	out := []byte{0x01}
	for start := 0; start < len(data); start += 4096 {
		end := start + 4096
		if end > len(data) {
			end = len(data)
		}
		out = append(out, compressChunk(data[start:end])...)
	}
	return out
}

// compressChunk compresses one chunk (at most 4096 bytes) and returns it with a
// 2-byte CompressedChunkHeader. It falls back to a raw chunk only when the
// compressed form does not fit in 4096 bytes ([MS-OVBA] §2.4.1.3.7; the decision
// is "does it fit in 4096", not "did it shrink").
func compressChunk(chunk []byte) []byte {
	var body []byte
	pos := 0
	for pos < len(chunk) {
		flagIndex := len(body)
		body = append(body, 0) // reserve the slot for the FlagByte
		var flag byte
		for i := 0; i < 8 && pos < len(chunk); i++ {
			off, length := longestMatch(chunk, pos)
			if length >= 3 {
				token := packCopyToken(pos, off, length)
				body = append(body, byte(token), byte(token>>8))
				flag |= 1 << uint(i)
				pos += length
			} else {
				body = append(body, chunk[pos])
				pos++
			}
		}
		body[flagIndex] = flag
	}

	// The chunk flag is decided by whether the compressed form fits in 4096 bytes ([MS-OVBA] §2.4.1.3.7).
	// Even if body ends up longer than the input, it stays compressed as long as it is under 4096.
	// CompressedChunkSize is the total byte count including the header. The low 12 bits = size-3.
	if len(body) < 4096 {
		size := len(body) + 2
		header := uint16(0xB000) | uint16((size-3)&0x0FFF) // 0b1011<<12: flag=1, sig=011
		return append([]byte{byte(header), byte(header >> 8)}, body...)
	}
	// raw chunk: flag=0, data padded to 4096. size=4098 -> low12=4095.
	raw := make([]byte, 4096)
	copy(raw, chunk)
	header := uint16(0x3000) | uint16(0x0FFF)
	return append([]byte{byte(header), byte(header >> 8)}, raw...)
}

// packCopyToken packs offset/length into a 2-byte CopyToken. [MS-OVBA] §2.4.1.3.19.2.
func packCopyToken(pos, offset, length int) uint16 {
	bitCount, _ := copyTokenHelp(pos)
	lengthBitCount := 16 - bitCount
	return (uint16(offset-1) << uint(lengthBitCount)) | uint16(length-3)
}

// copyTokenHelp returns the CopyToken bit split at position pos (measured from the start of the chunk).
// [MS-OVBA] §2.4.1.3.19.3. bitCount = max(ceil(log2(pos)), 4).
func copyTokenHelp(pos int) (bitCount, maxLength int) {
	bitCount = 4
	for (1 << bitCount) < pos {
		bitCount++
	}
	maxLength = (0xFFFF >> bitCount) + 3
	return bitCount, maxLength
}

// longestMatch greedily searches for the preceding subsequence that best matches chunk[pos:].
// [MS-OVBA] §2.4.1.3.19.4. It also allows overlapping copies (RLE) where candidate+length exceeds pos.
func longestMatch(chunk []byte, pos int) (bestOff, bestLen int) {
	_, maxLen := copyTokenHelp(pos)
	for candidate := pos - 1; candidate >= 0; candidate-- {
		length := 0
		for pos+length < len(chunk) &&
			length < maxLen &&
			chunk[candidate+length] == chunk[pos+length] {
			length++
		}
		if length > bestLen {
			bestLen = length
			bestOff = pos - candidate
		}
	}
	if bestLen < 3 { // a length below 3 cannot become a CopyToken (a literal is shorter)
		return 0, 0
	}
	return bestOff, bestLen
}
