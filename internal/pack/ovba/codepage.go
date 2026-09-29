package ovba

import (
	"fmt"
	"strings"
	"unicode/utf8"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
)

// codePageEncoding maps a VBA PROJECTCODEPAGE value to its codec. Only the
// practically relevant Windows ANSI code pages and UTF-8 are supported; an
// unknown code page must fail deterministically rather than guess an encoding
// (misdecoding would silently corrupt names and source).
func codePageEncoding(codepage uint16) (encoding.Encoding, error) {
	switch codepage {
	case 874:
		return charmap.Windows874, nil
	case 932:
		return japanese.ShiftJIS, nil
	case 936:
		return simplifiedchinese.GBK, nil
	case 949:
		return korean.EUCKR, nil
	case 950:
		return traditionalchinese.Big5, nil
	case 1250:
		return charmap.Windows1250, nil
	case 1251:
		return charmap.Windows1251, nil
	case 1252:
		return charmap.Windows1252, nil
	case 1253:
		return charmap.Windows1253, nil
	case 1254:
		return charmap.Windows1254, nil
	case 1255:
		return charmap.Windows1255, nil
	case 1256:
		return charmap.Windows1256, nil
	case 1257:
		return charmap.Windows1257, nil
	case 1258:
		return charmap.Windows1258, nil
	case 65001:
		return nil, nil // UTF-8 is handled natively in EncodeMBCS/DecodeMBCS
	default:
		return nil, fmt.Errorf("ovba: unsupported codepage %d", codepage)
	}
}

// DecodeMBCS converts a PROJECTCODEPAGE-encoded byte sequence to a Go string.
// Invalid input returns an error so callers never operate on mojibake.
func DecodeMBCS(b []byte, codepage uint16) (string, error) {
	if codepage == 65001 {
		if !utf8.Valid(b) {
			return "", fmt.Errorf("ovba: invalid UTF-8 in codepage %d", codepage)
		}
		return string(b), nil
	}
	enc, err := codePageEncoding(codepage)
	if err != nil {
		return "", err
	}
	out, err := enc.NewDecoder().Bytes(b)
	if err != nil {
		return "", fmt.Errorf("ovba: cannot decode as codepage %d: %w", codepage, err)
	}
	s := string(out)
	// x/text decoders substitute U+FFFD for undecodable input instead of
	// failing. A dir/module field must never silently accept mojibake, so a
	// decoded U+FFFD is rejected whenever the code page cannot encode that
	// character (which is the case for every supported MBCS page).
	if strings.ContainsRune(s, '\uFFFD') {
		if _, encErr := EncodeMBCS("\uFFFD", codepage); encErr != nil {
			return "", fmt.Errorf("ovba: cannot decode as codepage %d: input contains undecodable bytes", codepage)
		}
	}
	return s, nil
}

// EncodeMBCS converts a Go string to the PROJECTCODEPAGE byte encoding. A
// string containing characters the target code page cannot represent returns
// an error instead of emitting silent replacement bytes.
func EncodeMBCS(s string, codepage uint16) ([]byte, error) {
	if codepage == 65001 {
		if !utf8.ValidString(s) {
			return nil, fmt.Errorf("ovba: invalid UTF-8 string for codepage %d", codepage)
		}
		return []byte(s), nil
	}
	enc, err := codePageEncoding(codepage)
	if err != nil {
		return nil, err
	}
	out, err := enc.NewEncoder().Bytes([]byte(s))
	if err != nil {
		return nil, fmt.Errorf("ovba: cannot encode %q to codepage %d: %w", s, codepage, err)
	}
	return out, nil
}
