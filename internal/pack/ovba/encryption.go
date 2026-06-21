package ovba

import (
	"encoding/binary"
	"errors"
	"fmt"
)

// DecryptData reverses the MS-OVBA "Data Encryption" scheme (§2.4.3.3 Decryption).
// Despite the spec's name this is reversible obfuscation, not cryptography: the
// key is derived from the data's own Seed byte, so no password is required.
//
// enc is the raw bytes of an Encrypted Data Structure (§2.4.3.1):
//
//	Seed(1) VersionEnc(1) ProjKeyEnc(1) IgnoredEnc(IgnoredLength) LengthEnc(4) DataEnc(Length)
//
// It returns the decrypted Data. A Version other than 2, a truncated structure,
// or a Length that overruns the input is a fail-loud error rather than a silent
// misread.
func DecryptData(enc []byte) ([]byte, error) {
	if len(enc) < 3 {
		return nil, fmt.Errorf("ovba: encrypted data too short (%d bytes)", len(enc))
	}
	seed := enc[0]
	versionEnc := enc[1]
	projKeyEnc := enc[2]

	if version := seed ^ versionEnc; version != 2 {
		return nil, fmt.Errorf("ovba: unexpected encryption version %d (want 2)", version)
	}
	projKey := seed ^ projKeyEnc

	// State per §2.4.3.3. Byte arithmetic wraps mod 256 (Go uint8).
	unencByte1 := projKey
	encByte1 := projKeyEnc
	encByte2 := versionEnc

	pos := 3
	decryptNext := func() (byte, error) {
		if pos >= len(enc) {
			return 0, errors.New("ovba: encrypted data ended prematurely")
		}
		be := enc[pos]
		pos++
		b := be ^ (encByte2 + unencByte1)
		encByte2 = encByte1
		encByte1 = be
		unencByte1 = b
		return b, nil
	}

	ignoredLen := int((seed & 6) / 2)
	for i := 0; i < ignoredLen; i++ {
		if _, err := decryptNext(); err != nil {
			return nil, err
		}
	}

	var lenBuf [4]byte
	for i := range lenBuf {
		b, err := decryptNext()
		if err != nil {
			return nil, err
		}
		lenBuf[i] = b
	}
	dataLen := binary.LittleEndian.Uint32(lenBuf[:])

	out := make([]byte, 0, dataLen)
	for i := uint32(0); i < dataLen; i++ {
		b, err := decryptNext()
		if err != nil {
			return nil, err
		}
		out = append(out, b)
	}
	return out, nil
}
