package lspserver

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	formevents "github.com/harumiWeb/xlflow/internal/vba/userforms/events"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const (
	userFormEventsMethod    = "xlflow/userFormEvents"
	userFormEventEditMethod = "xlflow/userFormEventEdit"
)

type userFormEventsParams struct {
	URI      string `json:"uri"`
	Version  int    `json:"version"`
	Text     string `json:"text"`
	TargetID string `json:"targetId,omitempty"`
}

type userFormEventsResult struct {
	Version    int                     `json:"version"`
	TargetID   string                  `json:"targetId,omitempty"`
	ObjectName string                  `json:"objectName"`
	Events     []formevents.Descriptor `json:"events"`
	Error      string                  `json:"error,omitempty"`
}

type userFormEventEditParams struct {
	userFormEventsParams
	EventName   string `json:"eventName"`
	CodeURI     string `json:"codeURI"`
	CodeVersion int    `json:"codeVersion"`
	CodeText    string `json:"codeText"`
	Indent      string `json:"indent"`
	Newline     string `json:"newline"`
	CodeSource  string `json:"codeSource"`
}

type userFormEventEditResult struct {
	Version     int                 `json:"version"`
	CodeVersion int                 `json:"codeVersion"`
	Edits       []protocol.TextEdit `json:"edits"`
	Selection   protocol.Range      `json:"selection"`
	Created     bool                `json:"created"`
	Error       string              `json:"error,omitempty"`
}

func (s *Server) userFormEvents(params userFormEventsParams) userFormEventsResult {
	result := userFormEventsResult{
		Version: params.Version, TargetID: params.TargetID, Events: []formevents.Descriptor{},
	}
	input, pathError := s.userFormSpecInput(params.URI)
	if pathError != nil {
		result.Error = pathError.Message
		return result
	}
	document, err := forms.ParseFormSpec(input, []byte(params.Text))
	if err != nil {
		result.Error = err.Error()
		return result
	}
	if err := validateUserFormSpecIdentity(input, document); err != nil {
		result.Error = err.Error()
		return result
	}
	target, err := formevents.ResolveTarget(document, params.TargetID)
	if err != nil {
		result.Error = err.Error()
		return result
	}
	result.ObjectName = target.ObjectName
	result.Events, _ = formevents.ForType(target.Type, target.ObjectName)
	return result
}

func (s *Server) userFormEventEdit(params userFormEventEditParams) userFormEventEditResult {
	result := userFormEventEditResult{
		Version: params.Version, CodeVersion: params.CodeVersion, Edits: []protocol.TextEdit{},
	}
	input, pathError := s.userFormSpecInput(params.URI)
	if pathError != nil {
		result.Error = pathError.Message
		return result
	}
	form, err := forms.ParseFormSpec(input, []byte(params.Text))
	if err != nil {
		result.Error = err.Error()
		return result
	}
	if err := validateUserFormSpecIdentity(input, form); err != nil {
		result.Error = err.Error()
		return result
	}
	target, err := formevents.ResolveTarget(form, params.TargetID)
	if err != nil {
		result.Error = err.Error()
		return result
	}
	if params.Newline != "\n" && params.Newline != "\r\n" {
		result.Error = "newline must be LF or CRLF"
		return result
	}
	if strings.Trim(params.Indent, " \t") != "" {
		result.Error = "indent may contain only spaces and tabs"
		return result
	}
	if params.CodeSource != s.opts.Config.UserForm.CodeSource {
		result.Error = "codeSource does not match the configured UserForm source mode"
		return result
	}
	if params.CodeSource != "sidecar" && params.CodeSource != "frm" {
		result.Error = "codeSource must be sidecar or frm"
		return result
	}
	codePath, err := s.validateUserFormCodeAuthority(form, params.CodeURI, params.CodeSource)
	if err != nil {
		result.Error = err.Error()
		return result
	}
	codeBase, err := userFormCodeBase(params.CodeText, params.CodeSource, form.Form.Name)
	if err != nil {
		result.Error = err.Error()
		return result
	}
	code := []byte(params.CodeText)
	codeSection := code[codeBase:]
	parsed, err := vbaast.ParseDocument(codePath, codeSection)
	if err != nil {
		result.Error = fmt.Sprintf("parse UserForm code: %v", err)
		return result
	}
	defer parsed.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{
		Path: codePath, ModuleName: form.Form.Name, ModuleKind: "class",
	}, parsed)
	if err != nil {
		result.Error = fmt.Sprintf("parse UserForm procedures: %v", err)
		return result
	}
	if ir.Parse.HasError || ir.Parse.HasMissing {
		result.Error = "UserForm code has syntax errors; event creation is unavailable until they are fixed"
		return result
	}
	var eventDescriptor *formevents.Descriptor
	descriptors, _ := formevents.ForType(target.Type, target.ObjectName)
	for _, descriptor := range descriptors {
		if strings.EqualFold(descriptor.Name, params.EventName) {
			eventDescriptor = &descriptor
			break
		}
	}
	if eventDescriptor == nil {
		result.Error = fmt.Sprintf("event %q is not supported for %s", params.EventName, target.Type)
		return result
	}
	procedureName := target.ObjectName + "_" + eventDescriptor.Name
	for _, declaration := range ir.Declarations {
		if (declaration.Scope != procedureir.ScopeModule && declaration.Scope != procedureir.ScopeProject) ||
			!strings.EqualFold(declaration.Name, procedureName) {
			continue
		}
		if declaration.Recovered || len(declaration.ConditionalBranches) != 0 {
			result.Error = fmt.Sprintf("module declaration %q has an ambiguous or recovered declaration", procedureName)
		} else {
			result.Error = fmt.Sprintf("module declaration %q conflicts with the event procedure", procedureName)
		}
		return result
	}
	matching := make([]procedureir.ProcedureIR, 0, 1)
	for _, procedure := range ir.Procedures {
		if strings.EqualFold(procedure.Symbol.Name, procedureName) {
			matching = append(matching, procedure)
		}
	}
	if len(matching) > 1 {
		result.Error = fmt.Sprintf("procedure %q is declared more than once", procedureName)
		return result
	}
	if len(matching) == 1 {
		procedure := matching[0].Symbol
		if procedure.Recovered || len(procedure.ConditionalBranches) != 0 {
			result.Error = fmt.Sprintf("procedure %q has an ambiguous or recovered declaration", procedureName)
			return result
		}
		if procedure.Kind != procedureir.ProcedureSub {
			result.Error = fmt.Sprintf("procedure %q exists but is not a Sub", procedureName)
			return result
		}
		selectionOffset := codeBase + procedure.BodyRange.StartByte
		position := sourcePositionAtUTF16(code, selectionOffset)
		result.Selection = protocol.Range{Start: position, End: position}
		return result
	}
	insertText, selectionOffset := userFormEventInsertion(code, params, eventDescriptor.Signature)
	position := sourcePositionAtUTF16(code, len(code))
	result.Edits = append(result.Edits, protocol.TextEdit{
		Range: protocol.Range{Start: position, End: position}, NewText: insertText,
	})
	selectionOffset += len(code)
	selection := sourcePositionAtUTF16(append(bytes.Clone(code), []byte(insertText)...), selectionOffset)
	result.Selection = protocol.Range{Start: selection, End: selection}
	result.Created = true
	return result
}

func validateUserFormSpecIdentity(input forms.SpecInput, form forms.FormSpec) error {
	filename := strings.TrimSuffix(filepath.Base(input.DisplayPath), filepath.Ext(input.DisplayPath))
	if !strings.EqualFold(filename, form.Form.Name) {
		return fmt.Errorf("FormSpec filename %q must match form.name %q", filename, form.Form.Name)
	}
	return nil
}

func (s *Server) validateUserFormCodeAuthority(form forms.FormSpec, codeURI, codeSource string) (string, error) {
	codePath, err := fileURIToPath(codeURI)
	if err != nil {
		return "", fmt.Errorf("invalid codeURI: %w", err)
	}
	codePath, err = filepath.Abs(codePath)
	if err != nil {
		return "", fmt.Errorf("resolve codeURI: %w", err)
	}
	formsRoot := s.opts.Config.Src.Forms
	if !filepath.IsAbs(formsRoot) {
		formsRoot = filepath.Join(s.opts.RootDir, filepath.FromSlash(strings.ReplaceAll(formsRoot, "\\", "/")))
	}
	formsRoot, err = filepath.Abs(formsRoot)
	if err != nil {
		return "", fmt.Errorf("resolve configured forms root: %w", err)
	}
	if !pathIsWithin(formsRoot, codePath) {
		return "", fmt.Errorf("codeURI must be inside the configured UserForm forms root")
	}
	if err := validateUserFormEventCodePath(s.opts.RootDir, codePath); err != nil {
		return "", err
	}

	if codeSource == "sidecar" {
		if !sameAbsolutePath(filepath.Join(formsRoot, "code"), filepath.Dir(codePath)) ||
			!strings.EqualFold(filepath.Ext(codePath), ".bas") ||
			!strings.EqualFold(strings.TrimSuffix(filepath.Base(codePath), filepath.Ext(codePath)), form.Form.Name) {
			return "", fmt.Errorf("sidecar codeURI must resolve to code/%s.bas under the configured forms root", form.Form.Name)
		}
		return codePath, nil
	}
	if !strings.EqualFold(filepath.Ext(codePath), ".frm") || !strings.EqualFold(strings.TrimSuffix(filepath.Base(codePath), filepath.Ext(codePath)), form.Form.Name) {
		return "", fmt.Errorf("frm codeURI basename must match UserForm %q", form.Form.Name)
	}
	relative, err := filepath.Rel(formsRoot, codePath)
	if err != nil || isReservedFormCodePath(relative) {
		return "", fmt.Errorf("frm codeURI must identify a .frm artifact under the configured forms root")
	}
	return codePath, nil
}

// Check that existing components, including the file itself, resolve to their
// expected location. Missing suffixes are valid for a new sidecar.
func validateUserFormEventCodePath(root, codePath string) error {
	root, err := filepath.Abs(root)
	if err != nil {
		return fmt.Errorf("resolve workspace root: %w", err)
	}
	if !pathIsWithin(root, codePath) {
		return fmt.Errorf("UserForm code must be inside the workspace")
	}
	relative, err := filepath.Rel(root, codePath)
	if err != nil {
		return err
	}
	resolvedRoot, err := filepath.EvalSymlinks(root)
	if err != nil {
		return fmt.Errorf("resolve workspace root: %w", err)
	}
	current := root
	expected := resolvedRoot
	for _, component := range strings.Split(relative, string(filepath.Separator)) {
		current = filepath.Join(current, component)
		expected = filepath.Join(expected, component)
		_, err := os.Lstat(current)
		if errors.Is(err, fs.ErrNotExist) {
			return nil
		}
		if err != nil {
			return fmt.Errorf("inspect UserForm code path: %w", err)
		}
		resolved, err := filepath.EvalSymlinks(current)
		if err != nil {
			return fmt.Errorf("resolve UserForm code path: %w", err)
		}
		if !sameAbsolutePath(expected, resolved) {
			return fmt.Errorf("UserForm event code paths must not contain symbolic links or junctions")
		}
	}
	return nil
}

func pathIsWithin(root, path string) bool {
	if runtime.GOOS == "windows" {
		root, path = strings.ToLower(root), strings.ToLower(path)
	}
	rel, err := filepath.Rel(root, path)
	if err != nil || filepath.IsAbs(rel) || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return false
	}
	return true
}

func sameAbsolutePath(left, right string) bool {
	leftAbs, leftErr := filepath.Abs(left)
	rightAbs, rightErr := filepath.Abs(right)
	if leftErr != nil || rightErr != nil {
		return false
	}
	leftAbs, rightAbs = filepath.Clean(leftAbs), filepath.Clean(rightAbs)
	if runtime.GOOS == "windows" {
		return strings.EqualFold(leftAbs, rightAbs)
	}
	return leftAbs == rightAbs
}

func isReservedFormCodePath(relative string) bool {
	first, _, _ := strings.Cut(filepath.ToSlash(relative), "/")
	return strings.EqualFold(first, "code") || strings.EqualFold(first, "specs")
}

var vbNameAttribute = regexp.MustCompile(`(?i)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$`)

func userFormCodeBase(text, codeSource, formName string) (int, error) {
	if codeSource != "frm" {
		for lineNumber, line := range strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n") {
			if strings.HasPrefix(strings.ToLower(strings.TrimSpace(line)), "attribute vb_") {
				return 0, fmt.Errorf("sidecar code must not contain Attribute VB_* header lines (line %d)", lineNumber+1)
			}
		}
		return 0, nil
	}
	base := 0
	lastAttributeEnd := -1
	matchedName := ""
	nameAttributeCount := 0
	for base < len(text) {
		lineEnd := strings.IndexByte(text[base:], '\n')
		end := len(text)
		next := len(text)
		if lineEnd >= 0 {
			end = base + lineEnd
			next = end + 1
		}
		line := strings.TrimSuffix(text[base:end], "\r")
		if match := vbNameAttribute.FindStringSubmatch(line); match != nil {
			matchedName = match[1]
			nameAttributeCount++
		}
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(line)), "attribute vb_") {
			lastAttributeEnd = next
		}
		base = next
	}
	if nameAttributeCount != 1 {
		return 0, fmt.Errorf(".frm must contain exactly one Attribute VB_Name")
	}
	if matchedName == "" || !strings.EqualFold(matchedName, formName) {
		return 0, fmt.Errorf(".frm Attribute VB_Name must match UserForm %q", formName)
	}
	if lastAttributeEnd < 0 {
		return 0, fmt.Errorf(".frm is missing its Attribute VB_* header")
	}
	codeBase := lastAttributeEnd
	for codeBase < len(text) {
		end := codeBase
		for end < len(text) && text[end] != '\r' && text[end] != '\n' {
			end++
		}
		if strings.TrimSpace(text[codeBase:end]) != "" {
			break
		}
		if end < len(text) && text[end] == '\r' {
			end++
		}
		if end < len(text) && text[end] == '\n' {
			end++
		}
		if end == codeBase {
			break
		}
		codeBase = end
	}
	return codeBase, nil
}

func userFormEventInsertion(source []byte, params userFormEventEditParams, signature string) (string, int) {
	newline := params.Newline
	var insertion strings.Builder
	if len(source) > 0 {
		trailing := trailingLineBreaks(source)
		switch trailing {
		case 0:
			insertion.WriteString(newline)
			insertion.WriteString(newline)
		case 1:
			insertion.WriteString(newline)
		}
	}
	if params.CodeSource == "sidecar" && strings.TrimSpace(params.CodeText) == "" {
		insertion.WriteString("Option Explicit")
		insertion.WriteString(newline)
		insertion.WriteString(newline)
	}
	insertion.WriteString(signature)
	insertion.WriteString(newline)
	insertion.WriteString(params.Indent)
	selectionOffset := insertion.Len()
	insertion.WriteString(newline)
	insertion.WriteString("End Sub")
	insertion.WriteString(newline)
	return insertion.String(), selectionOffset
}

func trailingLineBreaks(source []byte) int {
	count := 0
	for index := len(source) - 1; index >= 0; {
		switch source[index] {
		case '\n':
			count++
			index--
			if index >= 0 && source[index] == '\r' {
				index--
			}
		case '\r':
			count++
			index--
		default:
			return count
		}
	}
	return count
}

func parseUserFormEventsParams(raw []byte) (userFormEventsParams, error) {
	fields, err := strictJSONObject(raw)
	if err != nil {
		return userFormEventsParams{}, err
	}
	if err := validateJSONFieldSet(fields, []string{"uri", "version", "text"}, []string{"targetId"}); err != nil {
		return userFormEventsParams{}, err
	}
	return parseUserFormEventsParamsFromFields(fields)
}

func parseUserFormEventEditParams(raw []byte) (userFormEventEditParams, error) {
	fields, err := strictJSONObject(raw)
	if err != nil {
		return userFormEventEditParams{}, err
	}
	if err := validateJSONFieldSet(fields,
		[]string{"uri", "version", "text", "eventName", "codeURI", "codeVersion", "codeText", "indent", "newline", "codeSource"},
		[]string{"targetId"}); err != nil {
		return userFormEventEditParams{}, err
	}
	common, err := parseUserFormEventsParamsFromFields(fields)
	if err != nil {
		return userFormEventEditParams{}, err
	}
	params := userFormEventEditParams{userFormEventsParams: common}
	if params.EventName, err = decodeRequiredJSONString(fields, "eventName"); err != nil {
		return params, err
	}
	if params.CodeURI, err = decodeRequiredJSONString(fields, "codeURI"); err != nil {
		return params, err
	}
	if params.CodeVersion, err = decodeRequiredJSONInt(fields, "codeVersion"); err != nil || params.CodeVersion < 0 {
		return params, fmt.Errorf("codeVersion must be a non-negative integer")
	}
	if params.CodeText, err = decodeRequiredJSONText(fields, "codeText"); err != nil {
		return params, err
	}
	if params.Indent, err = decodeRequiredJSONText(fields, "indent"); err != nil {
		return params, err
	}
	if params.Newline, err = decodeRequiredJSONString(fields, "newline"); err != nil {
		return params, err
	}
	if params.CodeSource, err = decodeRequiredJSONString(fields, "codeSource"); err != nil {
		return params, err
	}
	return params, nil
}

func parseUserFormEventsParamsFromFields(fields map[string]json.RawMessage) (userFormEventsParams, error) {
	params := userFormEventsParams{}
	var err error
	if params.URI, err = decodeRequiredJSONString(fields, "uri"); err != nil {
		return params, err
	}
	if params.Version, err = decodeRequiredJSONInt(fields, "version"); err != nil || params.Version < 0 {
		return params, fmt.Errorf("version must be a non-negative integer")
	}
	if params.Text, err = decodeRequiredJSONText(fields, "text"); err != nil {
		return params, err
	}
	if rawTarget, exists := fields["targetId"]; exists {
		if bytes.Equal(bytes.TrimSpace(rawTarget), []byte("null")) || json.Unmarshal(rawTarget, &params.TargetID) != nil || params.TargetID == "" {
			return params, fmt.Errorf("targetId must be a non-empty string when provided")
		}
	}
	return params, nil
}

func strictJSONObject(raw []byte) (map[string]json.RawMessage, error) {
	if err := rejectDuplicateJSONFields(raw); err != nil {
		return nil, err
	}
	return decodeJSONObject(raw)
}

func decodeRequiredJSONText(fields map[string]json.RawMessage, key string) (string, error) {
	raw, exists := fields[key]
	var value string
	if !exists || bytes.Equal(bytes.TrimSpace(raw), []byte("null")) || json.Unmarshal(raw, &value) != nil {
		return "", fmt.Errorf("%s must be a string", key)
	}
	return value, nil
}

func (s *Server) dispatchUserFormEvents(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	if req.Notif {
		return
	}
	if !s.handler.IsInitialized() {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: -32002, Message: "server not initialized"})
		return
	}
	var raw []byte
	if req.Params != nil {
		raw = *req.Params
	}
	params, err := parseUserFormEventsParams(raw)
	if err != nil {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid UserForm events parameters: " + err.Error()})
		return
	}
	_ = conn.Reply(ctx, req.ID, s.userFormEvents(params))
}

func (s *Server) dispatchUserFormEventEdit(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	if req.Notif {
		return
	}
	if !s.handler.IsInitialized() {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: -32002, Message: "server not initialized"})
		return
	}
	var raw []byte
	if req.Params != nil {
		raw = *req.Params
	}
	params, err := parseUserFormEventEditParams(raw)
	if err != nil {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid UserForm event edit parameters: " + err.Error()})
		return
	}
	_ = conn.Reply(ctx, req.ID, s.userFormEventEdit(params))
}
