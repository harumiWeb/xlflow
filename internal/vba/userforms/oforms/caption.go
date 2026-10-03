package oforms

import (
	"bytes"
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

func generatedVBFrame(name, caption string, size Size, codePage uint16) ([]byte, error) {
	if err := validateVBFrameCaption(caption, codePage); err != nil {
		return nil, err
	}
	frame := fmt.Sprintf("VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} %s\r\n   Caption = \"%s\"\r\n   ClientHeight = %d\r\n   ClientWidth = %d\r\n   StartUpPosition = 1\r\nEnd\r\n", name, escapeVBFrameString(caption),
		int64(math.Round(float64(size.Height)*1440/2540)), int64(math.Round(float64(size.Width)*1440/2540)))
	return ovba.EncodeMBCS(frame, codePage)
}

func rewriteVBFrameCaption(raw []byte, caption string, codePage uint16) ([]byte, error) {
	if err := validateVBFrameCaption(caption, codePage); err != nil {
		return nil, err
	}
	lines := bytes.SplitAfter(raw, []byte{'\n'})
	layout, err := inspectVBFrameLayout(lines, codePage)
	if err != nil {
		return nil, err
	}
	result := make([]byte, 0, len(raw))
	for index, line := range lines {
		body, ending := vbFrameLineParts(line)
		_, err := ovba.DecodeMBCS(body, codePage)
		if err != nil {
			return nil, fmt.Errorf("%w: decode VBFrame layout: %v", ErrInvalidEdit, err)
		}
		if index == layout.rootCaptionIndex {
			if len(ending) == 0 {
				return nil, fmt.Errorf("%w: root Caption line has no line ending", ErrInvalidEdit)
			}
			replacement, err := encodeVBFrameCaptionLine(caption, ending, codePage)
			if err != nil {
				return nil, err
			}
			result = append(result, replacement...)
		} else {
			result = append(result, line...)
			if index == layout.rootBeginIndex && layout.rootCaptionIndex < 0 {
				if len(ending) == 0 {
					return nil, fmt.Errorf("%w: root Begin line has no line ending", ErrInvalidEdit)
				}
				captionLine, err := encodeVBFrameCaptionLine(caption, ending, codePage)
				if err != nil {
					return nil, err
				}
				result = append(result, captionLine...)
			}
		}
	}
	return result, nil
}

type vbFrameLayout struct {
	rootBeginIndex    int
	rootCaptionIndex  int
	rootTypeInfoIndex int
}

func inspectVBFrameLayout(lines [][]byte, codePage uint16) (vbFrameLayout, error) {
	layout := vbFrameLayout{rootBeginIndex: -1, rootCaptionIndex: -1, rootTypeInfoIndex: -1}
	depth, propertyDepth := 0, 0
	for index, line := range lines {
		body, ending := vbFrameLineParts(line)
		text, err := ovba.DecodeMBCS(body, codePage)
		if err != nil {
			return vbFrameLayout{}, fmt.Errorf("%w: decode VBFrame layout: %v", ErrInvalidEdit, err)
		}
		trimmed := strings.TrimSpace(text)
		lower := strings.ToLower(trimmed)
		if strings.HasPrefix(lower, "beginproperty ") {
			if depth == 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: property block outside root in VBFrame", ErrInvalidEdit)
			}
			propertyDepth++
			continue
		}
		if lower == "endproperty" {
			propertyDepth--
			if propertyDepth < 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: unmatched EndProperty in VBFrame", ErrInvalidEdit)
			}
			continue
		}
		if propertyDepth != 0 {
			if isVBFrameBegin(trimmed) || strings.EqualFold(trimmed, "End") {
				return vbFrameLayout{}, fmt.Errorf("%w: component boundary within VBFrame property block", ErrInvalidEdit)
			}
			continue
		}
		if depth == 1 && isVBFrameProperty(trimmed, "Caption") {
			if layout.rootCaptionIndex >= 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: duplicate root Caption properties in VBFrame", ErrInvalidEdit)
			}
			if _, err := parseVBFrameCaption(trimmed); err != nil {
				return vbFrameLayout{}, fmt.Errorf("%w: malformed root Caption property: %v", ErrInvalidEdit, err)
			}
			if len(ending) == 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: root Caption line has no line ending", ErrInvalidEdit)
			}
			layout.rootCaptionIndex = index
		}
		if depth == 1 && isVBFrameProperty(trimmed, "TypeInfoVer") {
			if layout.rootTypeInfoIndex >= 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: duplicate root TypeInfoVer properties in VBFrame", ErrInvalidEdit)
			}
			layout.rootTypeInfoIndex = index
		}
		if depth == 0 && isVBFrameBegin(trimmed) {
			if layout.rootBeginIndex >= 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: multiple root Begin blocks in VBFrame", ErrInvalidEdit)
			}
			if len(ending) == 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: root Begin line has no line ending", ErrInvalidEdit)
			}
			layout.rootBeginIndex = index
		}
		if isVBFrameBegin(trimmed) {
			depth++
		} else if strings.EqualFold(trimmed, "End") {
			depth--
			if depth < 0 {
				return vbFrameLayout{}, fmt.Errorf("%w: unmatched End in VBFrame layout", ErrInvalidEdit)
			}
		}
	}
	if layout.rootBeginIndex < 0 || depth != 0 || propertyDepth != 0 {
		return vbFrameLayout{}, fmt.Errorf("%w: unknown VBFrame layout", ErrInvalidEdit)
	}
	return layout, nil
}

// Excel binds a stored TypeInfoVer to the root ShapeCookie. Keep that binding
// when membership changes; generated forms legitimately omit TypeInfoVer.
func rewriteVBFrameTypeInfo(raw []byte, cookie int64, codePage uint16) ([]byte, error) {
	lines := bytes.SplitAfter(raw, []byte{'\n'})
	layout, err := inspectVBFrameLayout(lines, codePage)
	if err != nil {
		return nil, err
	}
	if layout.rootTypeInfoIndex < 0 {
		return raw, nil
	}
	body, ending := vbFrameLineParts(lines[layout.rootTypeInfoIndex])
	text, err := ovba.DecodeMBCS(body, codePage)
	if err != nil {
		return nil, err
	}
	_, value, _ := strings.Cut(text, "=")
	value, _, _ = strings.Cut(value, "'")
	old, err := strconv.ParseUint(strings.TrimSpace(value), 10, 32)
	if err != nil || len(ending) == 0 || cookie < 0 || cookie > math.MaxUint32 {
		return nil, fmt.Errorf("%w: malformed root TypeInfoVer", ErrInvalidEdit)
	}
	if int64(old) == cookie {
		return raw, nil
	}
	line, err := ovba.EncodeMBCS(fmt.Sprintf("   TypeInfoVer = %d%s", cookie, ending), codePage)
	if err != nil {
		return nil, err
	}
	lines[layout.rootTypeInfoIndex] = line
	return bytes.Join(lines, nil), nil
}

func rootVBFrameCaption(raw []byte, codePage uint16) (string, bool, error) {
	lines := bytes.SplitAfter(raw, []byte{'\n'})
	layout, err := inspectVBFrameLayout(lines, codePage)
	if err != nil {
		return "", false, err
	}
	if layout.rootCaptionIndex < 0 {
		return "", false, nil
	}
	body, _ := vbFrameLineParts(lines[layout.rootCaptionIndex])
	text, err := ovba.DecodeMBCS(body, codePage)
	if err != nil {
		return "", false, fmt.Errorf("%w: decode root Caption: %v", ErrInvalidEdit, err)
	}
	caption, err := parseVBFrameCaption(strings.TrimSpace(text))
	if err != nil {
		return "", false, fmt.Errorf("%w: malformed root Caption property: %v", ErrInvalidEdit, err)
	}
	return caption, true, nil
}

func parseVBFrameCaption(line string) (string, error) {
	if !isVBFrameProperty(line, "Caption") {
		return "", fmt.Errorf("caption property is not an assignment")
	}
	value := strings.TrimSpace(line[len("Caption"):])
	if !strings.HasPrefix(value, "=") {
		return "", fmt.Errorf("caption property is not an assignment")
	}
	value = strings.TrimSpace(value[1:])
	if len(value) < 2 || value[0] != '"' || value[len(value)-1] != '"' {
		return "", fmt.Errorf("caption value must be a quoted string")
	}
	inner := value[1 : len(value)-1]
	var decoded strings.Builder
	decoded.Grow(len(inner))
	for index := 0; index < len(inner); index++ {
		if inner[index] != '"' {
			decoded.WriteByte(inner[index])
			continue
		}
		if index+1 >= len(inner) || inner[index+1] != '"' {
			return "", fmt.Errorf("caption value contains an unescaped quote")
		}
		decoded.WriteByte('"')
		index++
	}
	return decoded.String(), nil
}

func validateVBFrameCaption(caption string, codePage uint16) error {
	for _, r := range caption {
		if unicode.IsControl(r) {
			return fmt.Errorf("%w: root Caption cannot contain control characters", ErrInvalidEdit)
		}
	}
	if _, err := ovba.EncodeMBCS(caption, codePage); err != nil {
		return fmt.Errorf("%w: root Caption is not representable in code page %d: %v", ErrInvalidEdit, codePage, err)
	}
	return nil
}

func encodeVBFrameCaptionLine(caption string, ending []byte, codePage uint16) ([]byte, error) {
	line := "   Caption = \"" + escapeVBFrameString(caption) + "\"" + string(ending)
	encoded, err := ovba.EncodeMBCS(line, codePage)
	if err != nil {
		return nil, fmt.Errorf("%w: encode root Caption in code page %d: %v", ErrInvalidEdit, codePage, err)
	}
	return encoded, nil
}

func escapeVBFrameString(value string) string {
	return strings.ReplaceAll(value, `"`, `""`)
}

func vbFrameLineParts(line []byte) ([]byte, []byte) {
	if len(line) == 0 || line[len(line)-1] != '\n' {
		return line, nil
	}
	body := line[:len(line)-1]
	if len(body) > 0 && body[len(body)-1] == '\r' {
		return body[:len(body)-1], line[len(line)-2:]
	}
	return body, line[len(line)-1:]
}

func isVBFrameBegin(line string) bool {
	return strings.HasPrefix(strings.ToLower(line), "begin ")
}

func isVBFrameProperty(line, property string) bool {
	if len(line) < len(property) || !strings.EqualFold(line[:len(property)], property) {
		return false
	}
	return strings.HasPrefix(strings.TrimSpace(line[len(property):]), "=")
}
