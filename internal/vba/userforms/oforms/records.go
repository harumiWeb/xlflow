package oforms

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
	"unicode/utf16"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

const maxDesignerStreamSize = 64 << 20

type byteReader struct {
	data []byte
	pos  int
	end  int
}

func newByteReader(data []byte) *byteReader { return &byteReader{data: data, end: len(data)} }

func (r *byteReader) remaining() int { return r.end - r.pos }

func (r *byteReader) need(n int) error {
	if n < 0 || r.pos > r.end-n {
		return fmt.Errorf("read %d bytes at %d exceeds boundary %d", n, r.pos, r.end)
	}
	return nil
}

func (r *byteReader) take(n int) ([]byte, error) {
	if err := r.need(n); err != nil {
		return nil, err
	}
	b := r.data[r.pos : r.pos+n]
	r.pos += n
	return b, nil
}

func (r *byteReader) uint8() (uint8, error) {
	b, err := r.take(1)
	if err != nil {
		return 0, err
	}
	return b[0], nil
}

func (r *byteReader) uint16() (uint16, error) {
	b, err := r.take(2)
	if err != nil {
		return 0, err
	}
	return binary.LittleEndian.Uint16(b), nil
}

func (r *byteReader) int16() (int16, error) {
	v, err := r.uint16()
	return int16(v), err
}

func (r *byteReader) uint32() (uint32, error) {
	b, err := r.take(4)
	if err != nil {
		return 0, err
	}
	return binary.LittleEndian.Uint32(b), nil
}

func (r *byteReader) int32() (int32, error) {
	v, err := r.uint32()
	return int32(v), err
}

func (r *byteReader) align(base, size int) ([]byte, error) {
	over := (r.pos - base) % size
	if over == 0 {
		return nil, nil
	}
	return r.take(size - over)
}

func parseRecord(data []byte, spec *recordSpec, codePage uint16, limit int) (*Record, int, error) {
	if limit < 0 || limit > len(data) {
		return nil, 0, fmt.Errorf("record limit %d exceeds payload %d", limit, len(data))
	}
	r := newByteReader(data)
	r.end = limit
	base := r.pos
	minor, err := r.uint8()
	if err != nil {
		return nil, r.pos, err
	}
	major, err := r.uint8()
	if err != nil {
		return nil, r.pos, err
	}
	if minor != 0 || major != spec.major {
		return nil, r.pos - 2, fmt.Errorf("unexpected %s version %d.%d", spec.typeName, major, minor)
	}
	cb, err := r.uint16()
	if err != nil {
		return nil, r.pos, err
	}
	afterExtra := base + 4 + int(cb)
	if afterExtra > limit {
		return nil, r.pos, fmt.Errorf("%s cb %d exceeds record boundary %d", spec.typeName, cb, limit)
	}
	maskLow, err := r.uint32()
	if err != nil {
		return nil, r.pos, err
	}
	mask := uint64(maskLow)
	maskWidth := 4
	if spec.mask64 {
		high, err := r.uint32()
		if err != nil {
			return nil, r.pos, err
		}
		mask |= uint64(high) << 32
		maskWidth = 8
	}
	record := &Record{
		Type: spec.typeName, Minor: minor, Major: major, Mask: mask, MaskWidth: maskWidth,
		Values: map[string]int64{}, Strings: map[string]StoredString{}, Sizes: map[string]Size{},
		Arrays: map[string][]byte{}, Pictures: map[string][]byte{}, Padding: map[string][]byte{},
	}
	for bit, flag := range spec.flags {
		if mask&(uint64(1)<<bit) != 0 {
			record.Values[flag.name] = flag.value
		}
	}
	for _, field := range spec.data {
		if mask&(uint64(1)<<field.bit) == 0 {
			continue
		}
		pad, err := r.align(base, field.size)
		if err != nil {
			return nil, r.pos, err
		}
		if len(pad) > 0 {
			record.Padding["before:"+field.name] = bytes.Clone(pad)
		}
		value, err := readInteger(r, field.size, field.kind == fieldSigned)
		if err != nil {
			return nil, r.pos, err
		}
		record.Values[field.name] = value
	}
	pad, err := r.align(base, 4)
	if err != nil {
		return nil, r.pos, err
	}
	if len(pad) > 0 {
		record.Padding["data:end"] = bytes.Clone(pad)
	}
	for _, extra := range spec.extra {
		if mask&(uint64(1)<<extra.bit) == 0 {
			continue
		}
		switch extra.kind {
		case extraSize, extraPosition:
			width, err := r.int32()
			if err != nil {
				return nil, r.pos, err
			}
			height, err := r.int32()
			if err != nil {
				return nil, r.pos, err
			}
			record.Sizes[extra.name] = Size{Width: width, Height: height}
		case extraString:
			stored, err := readStoredString(r, uint32(record.Values[extra.name]), codePage)
			if err != nil {
				return nil, r.pos, fmt.Errorf("%s: %w", extra.name, err)
			}
			record.Strings[extra.name] = stored
			padding, err := r.align(base, 4)
			if err != nil {
				return nil, r.pos, err
			}
			if len(padding) > 0 {
				record.Padding["str:"+extra.name] = bytes.Clone(padding)
			}
		case extraArray:
			size := record.Values[extra.sizeFrom]
			if size < 0 || size > math.MaxInt {
				return nil, r.pos, fmt.Errorf("%s has invalid size %d", extra.name, size)
			}
			blob, err := r.take(int(size))
			if err != nil {
				return nil, r.pos, err
			}
			record.Arrays[extra.name] = bytes.Clone(blob)
		}
	}
	if r.pos != afterExtra {
		return nil, r.pos, fmt.Errorf("%s cb mismatch: consumed %d, expected %d", spec.typeName, r.pos-base-4, cb)
	}
	if spec.stopAfterExtra {
		record.Raw = bytes.Clone(data[base:afterExtra])
		return record, afterExtra, nil
	}
	for _, stream := range spec.stream {
		if mask&(uint64(1)<<stream.bit) == 0 {
			continue
		}
		picture, err := readGUIDAndPicture(r)
		if err != nil {
			return nil, r.pos, fmt.Errorf("%s: %w", stream.name, err)
		}
		record.Pictures[stream.name] = picture
	}
	if spec.textProps {
		if r.remaining() < 4 {
			return nil, r.pos, fmt.Errorf("%s missing TextProps", spec.typeName)
		}
		nestedLength := 4 + int(binary.LittleEndian.Uint16(data[r.pos+2:r.pos+4]))
		if nestedLength > r.remaining() {
			return nil, r.pos, fmt.Errorf("TextProps length %d exceeds remaining %d", nestedLength, r.remaining())
		}
		textProps, consumed, err := parseRecord(data[r.pos:r.end], &textPropsSpec, codePage, nestedLength)
		if err != nil {
			return nil, r.pos + consumed, err
		}
		record.TextProps = textProps
		r.pos += consumed
	}
	if r.pos < limit {
		if !spec.rawTail {
			return nil, r.pos, fmt.Errorf("%s has %d unexpected trailing bytes", spec.typeName, limit-r.pos)
		}
		record.TailRaw = bytes.Clone(data[r.pos:limit])
		r.pos = limit
	}
	record.Raw = bytes.Clone(data[base:limit])
	return record, r.pos, nil
}

func readInteger(r *byteReader, size int, signed bool) (int64, error) {
	switch size {
	case 1:
		v, err := r.uint8()
		if signed {
			return int64(int8(v)), err
		}
		return int64(v), err
	case 2:
		if signed {
			v, err := r.int16()
			return int64(v), err
		}
		v, err := r.uint16()
		return int64(v), err
	case 4:
		if signed {
			v, err := r.int32()
			return int64(v), err
		}
		v, err := r.uint32()
		return int64(v), err
	default:
		return 0, fmt.Errorf("unsupported integer width %d", size)
	}
}

func readStoredString(r *byteReader, packed uint32, codePage uint16) (StoredString, error) {
	size := int(packed & 0x7fffffff)
	raw, err := r.take(size)
	if err != nil {
		return StoredString{}, err
	}
	compressed := packed&0x80000000 != 0
	text, err := decodeStoredText(raw, compressed, codePage)
	if err != nil {
		return StoredString{}, err
	}
	return StoredString{Text: text, Compressed: compressed, Raw: bytes.Clone(raw)}, nil
}

func decodeStoredText(raw []byte, compressed bool, codePage uint16) (string, error) {
	if compressed {
		return ovba.DecodeMBCS(raw, codePage)
	}
	if len(raw)%2 != 0 {
		return "", fmt.Errorf("UTF-16 string has odd byte length %d", len(raw))
	}
	units := make([]uint16, len(raw)/2)
	for i := range units {
		units[i] = binary.LittleEndian.Uint16(raw[i*2:])
	}
	return string(utf16.Decode(units)), nil
}

func readGUIDAndPicture(r *byteReader) ([]byte, error) {
	start := r.pos
	if _, err := r.take(16); err != nil {
		return nil, err
	}
	if _, err := r.uint32(); err != nil {
		return nil, err
	}
	size, err := r.uint32()
	if err != nil {
		return nil, err
	}
	if uint64(size) > uint64(r.remaining()) {
		return nil, fmt.Errorf("picture size %d exceeds remaining %d", size, r.remaining())
	}
	if _, err := r.take(int(size)); err != nil {
		return nil, err
	}
	return bytes.Clone(r.data[start:r.pos]), nil
}
