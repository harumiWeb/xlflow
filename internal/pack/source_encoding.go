package pack

import (
	"bytes"
	"errors"
	"fmt"
	"maps"
	"path/filepath"
	"slices"
	"strconv"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

type sourceEncodingInput uint8

const (
	sourceEncodingCode sourceEncodingInput = iota
	sourceEncodingName
	sourceEncodingFormSpec
)

func collectSourceEncodingErrors(sources []SourceModule, codepage uint16) (*SourceEncodingErrors, error) {
	if _, err := ovba.EncodeMBCS("", codepage); err != nil {
		return nil, err
	}

	issues := &SourceEncodingErrors{}
	seenIssues := make(map[string]bool, len(sources))
	for _, source := range sources {
		if _, err := ovba.EncodeMBCS(source.Source, codepage); err != nil {
			issue := sourceEncodingErrorFor(source, err, sourceEncodingCode)
			if issue != nil {
				appendSourceEncodingIssue(issues, seenIssues, issue)
			} else {
				return nil, err
			}
		}
		if _, err := ovba.EncodeMBCS(source.Name, codepage); err != nil {
			issue := sourceEncodingErrorFor(source, err, sourceEncodingName)
			if issue != nil {
				appendSourceEncodingIssue(issues, seenIssues, issue)
			} else {
				return nil, err
			}
		}
		for _, value := range formSpecEncodingValues(source.FormSpec) {
			if _, err := ovba.EncodeMBCS(value, codepage); err == nil {
				continue
			} else if issue := sourceEncodingErrorFor(source, err, sourceEncodingFormSpec); issue != nil {
				appendSourceEncodingIssue(issues, seenIssues, issue)
			} else {
				return nil, err
			}
		}
	}
	if len(issues.Errors) == 0 {
		return nil, nil
	}
	return issues, nil
}

func appendSourceEncodingIssue(issues *SourceEncodingErrors, seenIssues map[string]bool, issue *SourceEncodingError) {
	key := strings.ToLower(filepath.Clean(issue.Path)) + "\x00" +
		strconv.Itoa(issue.Line) + "\x00" + strconv.Itoa(issue.Column) + "\x00" + strconv.FormatInt(int64(issue.Rune), 10)
	if seenIssues[key] {
		return
	}
	seenIssues[key] = true
	issues.Errors = append(issues.Errors, issue)
}

func sourceHasEncodingIssue(source SourceModule, issues *SourceEncodingErrors) bool {
	if issues == nil {
		return false
	}
	for _, issue := range issues.Errors {
		if issue == nil {
			continue
		}
		for _, path := range []string{source.SourcePath, sourceSpecPath(source), source.Name} {
			if path != "" && strings.EqualFold(issue.Path, path) {
				return true
			}
		}
	}
	return false
}

func sourceEncodingErrorFor(source SourceModule, err error, input sourceEncodingInput) *SourceEncodingError {
	encodingErr, ok := errors.AsType[*ovba.CodePageEncodingError](err)
	if !ok || encodingErr == nil || encodingErr.UnsupportedCodePage {
		return nil
	}

	path := source.SourcePath
	data := []byte(source.Source)
	field := "source"
	switch input {
	case sourceEncodingName:
		data = []byte(source.Source)
		field = "module_name"
		path = sourceLabel(source)
		if offset := strings.Index(string(data), source.Name); offset >= 0 {
			return newSourceEncodingError(path, data, offset+encodingErr.Offset, field, encodingErr)
		}
		data = []byte(source.Name)
		return newSourceEncodingError(path, data, encodingErr.Offset, field, encodingErr)
	case sourceEncodingFormSpec:
		path = sourceSpecPath(source)
		data = source.FormSpecSource
		field = "form_spec"
		if offset, found := findRune(data, encodingErr.Rune); found {
			return newSourceEncodingError(path, data, offset, field, encodingErr)
		}
		if offset, found := findEscapedRune(data, encodingErr.Rune); found {
			return newSourceEncodingError(path, data, offset, field, encodingErr)
		}
		return newSourceEncodingError(path, nil, 0, field, encodingErr)
	}

	if path == "" {
		path = source.Name
	}
	offset := encodingErr.Offset
	if foundOffset, found := findRune(data, encodingErr.Rune); found {
		offset = foundOffset
	}
	return newSourceEncodingError(path, data, offset, field, encodingErr)
}

func formSpecEncodingValues(input *spec.FormSpec) []string {
	if input == nil {
		return nil
	}
	values := make([]string, 0)
	appendString := func(value *string) {
		if value != nil {
			values = append(values, *value)
		}
	}
	appendString(&input.Form.Name)
	appendString(input.Form.Caption)
	var visitControl func(spec.FormSpecControl)
	visitControl = func(control spec.FormSpecControl) {
		values = append(values, control.Name)
		appendString(control.Caption)
		appendString(control.Tag)
		appendString(control.ControlTipText)
		appendString(control.Accelerator)
		appendString(control.Text)
		if value, ok := control.Value.(string); ok {
			values = append(values, value)
		}
		for _, key := range slices.Sorted(maps.Keys(control.Properties)) {
			value, ok := control.Properties[key].(string)
			if ok {
				if _, supported := spec.LookupControlProperty(control.Type, key); supported {
					values = append(values, value)
				}
			}
		}
		for _, tab := range control.Tabs {
			values = append(values, tab.Name)
			appendString(tab.Caption)
			appendString(tab.ControlTipText)
			appendString(tab.Tag)
			appendString(tab.Accelerator)
		}
		for _, child := range control.Controls {
			visitControl(child)
		}
	}
	for _, control := range input.Controls {
		visitControl(control)
	}
	return values
}

func sourceEncodingErrorFromError(source SourceModule, err error) *SourceEncodingError {
	encodingErr, ok := errors.AsType[*ovba.CodePageEncodingError](err)
	if !ok || encodingErr == nil || encodingErr.UnsupportedCodePage {
		return nil
	}
	if generation, ok := errors.AsType[*vbaproject.UserFormGenerationError](err); ok {
		reason := strings.ToLower(generation.Reason)
		if strings.Contains(reason, "codeonly") || strings.Contains(reason, "module source") {
			return sourceEncodingErrorFor(source, err, sourceEncodingCode)
		}
	}
	if source.Type == ModuleTypeForm && source.FormSpec != nil {
		return sourceEncodingErrorFor(source, encodingErr, sourceEncodingFormSpec)
	}
	return sourceEncodingErrorFor(source, encodingErr, sourceEncodingCode)
}

func appendSourceEncodingDiagnostic(diagnostics []error, issue *SourceEncodingError) []error {
	for _, diagnostic := range diagnostics {
		if existing, ok := errors.AsType[*SourceEncodingError](diagnostic); ok && existing != nil &&
			existing.Path == issue.Path && existing.Line == issue.Line && existing.Column == issue.Column && existing.Rune == issue.Rune {
			return diagnostics
		}
	}
	return append(diagnostics, issue)
}

func combinePackDiagnostics(diagnostics []error) error {
	var sourceIssues []*SourceEncodingError
	var otherIssues []error
	var collect func(error)
	collect = func(diagnostic error) {
		if diagnostic == nil {
			return
		}
		if aggregate, ok := errors.AsType[*DiagnosticsError](diagnostic); ok && aggregate != nil {
			for _, issue := range aggregate.Errors {
				collect(issue)
			}
			return
		}
		if aggregate, ok := errors.AsType[*SourceEncodingErrors](diagnostic); ok && aggregate != nil {
			for _, issue := range aggregate.Errors {
				if issue != nil {
					sourceIssues = appendUniqueSourceEncodingIssue(sourceIssues, issue)
				}
			}
			return
		}
		if issue, ok := errors.AsType[*SourceEncodingError](diagnostic); ok && issue != nil {
			sourceIssues = appendUniqueSourceEncodingIssue(sourceIssues, issue)
			return
		}
		otherIssues = append(otherIssues, diagnostic)
	}
	for _, diagnostic := range diagnostics {
		collect(diagnostic)
	}
	if len(sourceIssues) > 0 {
		sourceErrors := &SourceEncodingErrors{Errors: sourceIssues}
		if len(otherIssues) == 0 {
			return sourceErrors
		}
		return &DiagnosticsError{Errors: append([]error{sourceErrors}, otherIssues...)}
	}
	if len(otherIssues) == 0 {
		return nil
	}
	if len(otherIssues) == 1 {
		return otherIssues[0]
	}
	return &DiagnosticsError{Errors: otherIssues}
}

func appendUniqueSourceEncodingIssue(issues []*SourceEncodingError, candidate *SourceEncodingError) []*SourceEncodingError {
	for _, issue := range issues {
		if issue.Path == candidate.Path && issue.Line == candidate.Line && issue.Column == candidate.Column && issue.Rune == candidate.Rune {
			return issues
		}
	}
	return append(issues, candidate)
}

func hasSourceEncodingErrors(err error) bool {
	if _, ok := errors.AsType[*SourceEncodingErrors](err); ok {
		return true
	}
	_, ok := errors.AsType[*SourceEncodingError](err)
	return ok
}

func ambiguousLayoutUnlessEncoding(err error) error {
	if isCodePageEncodingError(err) || hasSourceEncodingErrors(err) {
		return err
	}
	return fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
}

func sourceSpecPath(source SourceModule) string {
	if source.FormSpecPath != "" {
		return source.FormSpecPath
	}
	for _, path := range source.RelatedPaths {
		switch strings.ToLower(filepath.Ext(path)) {
		case ".json", ".yaml", ".yml":
			return path
		}
	}
	if source.SourcePath != "" {
		return source.SourcePath
	}
	return source.Name
}

func newSourceEncodingError(path string, data []byte, offset int, field string, cause *ovba.CodePageEncodingError) *SourceEncodingError {
	if path == "" {
		path = "<source>"
	}
	if offset < 0 {
		offset = 0
	}
	if offset > len(data) {
		offset = len(data)
	}
	line, column, byteColumn := sourceLineColumn(data, offset)
	return &SourceEncodingError{
		Path:       path,
		Field:      field,
		Line:       line,
		Column:     column,
		ByteColumn: byteColumn,
		ByteOffset: offset,
		Rune:       cause.Rune,
		CodePage:   cause.CodePage,
	}
}

func findEscapedRune(data []byte, target rune) (int, bool) {
	var needles []string
	if target <= 0xffff {
		needles = []string{fmt.Sprintf("\\u%04x", target)}
	} else {
		first, second := utf16.EncodeRune(target)
		needles = []string{
			fmt.Sprintf("\\u%04x\\u%04x", first, second),
			fmt.Sprintf("\\u%08x", target),
		}
	}
	lower := bytes.ToLower(data)
	for _, needle := range needles {
		if offset := bytes.Index(lower, []byte(needle)); offset >= 0 {
			return offset, true
		}
	}
	return 0, false
}

func sourceLineColumn(data []byte, offset int) (line, column, byteColumn int) {
	line, column, byteColumn = 1, 1, 1
	for index := 0; index < offset; {
		switch data[index] {
		case '\r':
			line++
			column, byteColumn = 1, 1
			index++
			if index < offset && data[index] == '\n' {
				index++
			}
		case '\n':
			line++
			column, byteColumn = 1, 1
			index++
		default:
			_, size := utf8.DecodeRune(data[index:offset])
			if size == 0 {
				size = 1
			}
			index += size
			column++
			byteColumn += size
		}
	}
	return line, column, byteColumn
}

func findRune(data []byte, target rune) (int, bool) {
	for offset, current := range string(data) {
		if current == target {
			return offset, true
		}
	}
	return 0, false
}

func isCodePageEncodingError(err error) bool {
	encodingErr, ok := errors.AsType[*ovba.CodePageEncodingError](err)
	return ok && encodingErr != nil
}
