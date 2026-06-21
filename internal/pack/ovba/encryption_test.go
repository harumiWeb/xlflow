package ovba

import (
	"bytes"
	"encoding/binary"
	"testing"
)

// encryptData is the §2.4.3.2 inverse of DecryptData, used only to generate
// valid Encrypted Data Structures for round-trip testing. Seed and projKey are
// caller-chosen; ignored content is arbitrary (filled with 0x00).
func encryptData(seed, projKey byte, data []byte) []byte {
	const version = byte(2)
	versionEnc := seed ^ version
	projKeyEnc := seed ^ projKey

	unencByte1 := projKey
	encByte1 := projKeyEnc
	encByte2 := versionEnc

	out := []byte{seed, versionEnc, projKeyEnc}
	encNext := func(b byte) byte {
		be := b ^ (encByte2 + unencByte1)
		encByte2 = encByte1
		encByte1 = be
		unencByte1 = b
		return be
	}

	ignoredLen := int((seed & 6) / 2)
	for i := 0; i < ignoredLen; i++ {
		out = append(out, encNext(0x00))
	}
	var lenBuf [4]byte
	binary.LittleEndian.PutUint32(lenBuf[:], uint32(len(data)))
	for _, b := range lenBuf {
		out = append(out, encNext(b))
	}
	for _, b := range data {
		out = append(out, encNext(b))
	}
	return out
}

func TestDecryptDataRoundTrip(t *testing.T) {
	cases := []struct {
		name          string
		seed, projKey byte
		data          []byte
	}{
		// Vary seed so (seed & 6)/2 exercises IgnoredLength 0..3.
		{"ignored0", 0x00, 0x7B, []byte{0x01, 0x00, 0x00, 0x00}},
		{"ignored1", 0x02, 0x10, []byte{0x00, 0x00, 0x00, 0x00}},
		{"ignored2", 0x04, 0xAB, []byte{0x04, 0x00, 0x00, 0x00}},
		{"ignored3", 0x06, 0xFF, []byte("hello world")},
		{"empty", 0x06, 0x01, []byte{}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			enc := encryptData(c.seed, c.projKey, c.data)
			got, err := DecryptData(enc)
			if err != nil {
				t.Fatalf("DecryptData: %v", err)
			}
			if !bytes.Equal(got, c.data) {
				t.Errorf("round-trip mismatch: got %x want %x", got, c.data)
			}
		})
	}
}

func TestDecryptDataRejectsBadVersion(t *testing.T) {
	// Seed=0x00 so VersionEnc must be 0x02 for Version==2; use 0x03 to force a mismatch.
	enc := []byte{0x00, 0x03, 0x10}
	if _, err := DecryptData(enc); err == nil {
		t.Fatal("expected error for Version != 2, got nil")
	}
}

func TestDecryptDataRejectsShort(t *testing.T) {
	if _, err := DecryptData([]byte{0x00, 0x02}); err == nil {
		t.Fatal("expected error for input shorter than 3 bytes, got nil")
	}
}

func TestDecryptDataRejectsOversizedLength(t *testing.T) {
	// Craft a structure whose decrypted Length is ~4 GiB but with no data bytes,
	// so DecryptData must fail loudly on the length check rather than attempt a
	// multi-gigabyte allocation before discovering the input is too short.
	const seed, projKey = byte(0), byte(0x10) // seed=0 => IgnoredLength 0
	versionEnc := seed ^ 2
	projKeyEnc := seed ^ projKey
	unencByte1, encByte1, encByte2 := projKey, projKeyEnc, versionEnc
	enc := []byte{seed, versionEnc, projKeyEnc}
	encNext := func(b byte) byte {
		be := b ^ (encByte2 + unencByte1)
		encByte2, encByte1, unencByte1 = encByte1, be, b
		return be
	}
	var lenBuf [4]byte
	binary.LittleEndian.PutUint32(lenBuf[:], 0xFFFFFFFF)
	for _, b := range lenBuf {
		enc = append(enc, encNext(b))
	}
	// No data bytes follow.
	if _, err := DecryptData(enc); err == nil {
		t.Fatal("expected error for declared length exceeding input, got nil")
	}
}
