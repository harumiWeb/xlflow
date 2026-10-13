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

// CodePageEncodingError identifies a source character that cannot be encoded
// with a VBA project's PROJECTCODEPAGE, or a code page the writer does not
// support. Offset is the UTF-8 byte offset in the input string when Rune is
// non-zero.
type CodePageEncodingError struct {
	CodePage            uint16
	Rune                rune
	Offset              int
	UnsupportedCodePage bool
	InvalidUTF8         bool
	Cause               error
}

func (e *CodePageEncodingError) Error() string {
	if e == nil {
		return "ovba: codepage encoding failed"
	}
	if e.UnsupportedCodePage {
		return fmt.Sprintf("ovba: unsupported codepage %d", e.CodePage)
	}
	if e.InvalidUTF8 {
		return fmt.Sprintf("ovba: invalid UTF-8 for codepage %d at byte offset %d", e.CodePage, e.Offset)
	}
	return fmt.Sprintf("ovba: rune U+%04X cannot be encoded with codepage %d", e.Rune, e.CodePage)
}

func (e *CodePageEncodingError) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

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
		return nil, &CodePageEncodingError{CodePage: codepage, UnsupportedCodePage: true}
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
	if !utf8.ValidString(s) {
		return nil, &CodePageEncodingError{
			CodePage:    codepage,
			Rune:        utf8.RuneError,
			Offset:      firstInvalidUTF8Offset(s),
			InvalidUTF8: true,
		}
	}
	if codepage == 65001 {
		return []byte(s), nil
	}
	enc, err := codePageEncoding(codepage)
	if err != nil {
		return nil, err
	}
	out, err := enc.NewEncoder().Bytes([]byte(s))
	if err != nil {
		for offset, r := range s {
			if _, runeErr := enc.NewEncoder().Bytes([]byte(string(r))); runeErr != nil {
				return nil, &CodePageEncodingError{CodePage: codepage, Rune: r, Offset: offset, Cause: runeErr}
			}
		}
		return nil, &CodePageEncodingError{CodePage: codepage, Rune: utf8.RuneError, Cause: err}
	}
	return out, nil
}

func firstInvalidUTF8Offset(s string) int {
	for offset := 0; offset < len(s); {
		r, size := utf8.DecodeRuneInString(s[offset:])
		if r == utf8.RuneError && size == 1 {
			return offset
		}
		offset += size
	}
	return len(s)
}
