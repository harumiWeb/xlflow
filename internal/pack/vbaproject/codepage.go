package vbaproject

import (
	"fmt"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
)

// decodeMBCS converts the byte sequence of a VBA module source to a Go string according to the CODEPAGE.
func decodeMBCS(b []byte, codepage uint16) (string, error) {
	var enc encoding.Encoding
	switch codepage {
	case 932:
		enc = japanese.ShiftJIS
	case 1252:
		enc = charmap.Windows1252
	case 65001: // UTF-8 is passed through as-is (rare, since VBA source is usually MBCS)
		return string(b), nil
	// 1200 (UTF-16) would be silently corrupted if its bytes were misread as UTF-8, so it is not passed through and falls into the default error
	default:
		return "", fmt.Errorf("vbaproject: unsupported codepage %d", codepage)
	}
	out, err := enc.NewDecoder().Bytes(b)
	if err != nil {
		return "", err
	}
	return string(out), nil
}

// encodeMBCS converts a Go string to the MBCS byte sequence of the CODEPAGE (the inverse of decodeMBCS).
// If it contains characters the target codepage cannot represent, it returns an error (no silent mojibake).
func encodeMBCS(s string, codepage uint16) ([]byte, error) {
	var enc encoding.Encoding
	switch codepage {
	case 932:
		enc = japanese.ShiftJIS
	case 1252:
		enc = charmap.Windows1252
	case 65001:
		return []byte(s), nil
	default:
		return nil, fmt.Errorf("vbaproject: unsupported codepage %d", codepage)
	}
	out, err := enc.NewEncoder().Bytes([]byte(s))
	if err != nil {
		return nil, fmt.Errorf("vbaproject: cannot encode to codepage %d: %w", codepage, err)
	}
	return out, nil
}
