package vbaproject

import (
	"bytes"
	"crypto/rand"
	"encoding/binary"
	"encoding/hex"
	"fmt"
	"maps"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

const (
	UserFormGenerationInvalid            = "userform_generation_invalid"
	UserFormGenerationConflict           = "userform_generation_conflict"
	UserFormFormsReferenceRequired       = "forms_reference_required"
	UserFormModuleFormsReferenceRequired = UserFormFormsReferenceRequired
	userFormMSFormsReferenceGUID         = "0D452EE1-E08F-101A-852E-02608C4D0BB4"
)

var userFormMSFormsReferenceGUIDWire = [...]byte{
	0xE1, 0x2E, 0x45, 0x0D,
	0x8F, 0xE0,
	0x1A, 0x10,
	0x85, 0x2E,
	0x02, 0x60, 0x8C, 0x4D, 0x0B, 0xB4,
}

// UserFormGenerationError reports a rejected UserForm creation or addition.
// Code is stable; Form and Reason provide context for callers and logs.
type UserFormGenerationError struct {
	Code   string
	Form   string
	Reason string
	Cause  error
}

func (e *UserFormGenerationError) Error() string {
	if e.Form == "" {
		return fmt.Sprintf("%s: %s", e.Code, e.Reason)
	}
	return fmt.Sprintf("%s: %s: %s", e.Code, e.Form, e.Reason)
}

func (e *UserFormGenerationError) Unwrap() error { return e.Cause }

// UserFormGUIDGenerator supplies one canonical UUID string for a generated
// UserForm component identity.
type UserFormGUIDGenerator func() (string, error)

// UserFormModuleOptions controls nondeterministic parts of NewUserFormModule.
// The zero value uses crypto/rand.
type UserFormModuleOptions struct {
	GUIDGenerator UserFormGUIDGenerator
}

// NewUserFormModule constructs the in-bin UserForm code-behind module. The
// Designer's fixed MS-OFORMS VBFrame identity is deliberately not reused as a
// component identity; VB_Base receives two fresh UUIDs instead.
func NewUserFormModule(designer *oforms.Form, codeOnly string, codePage uint16, options ...UserFormModuleOptions) (Module, error) {
	if designer == nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, "", "designer is nil", nil)
	}
	if len(options) > 1 {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "more than one options value was supplied", nil)
	}
	if err := validateUserFormName(designer.Name, codePage); err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "invalid form identity", err)
	}
	if _, err := oforms.SerializeForm(designer, codePage); err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "designer is not a valid lossless form", err)
	}
	if hasUserFormAttributeLine(codeOnly) {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "codeOnly must not contain Attribute VB_ lines", nil)
	}
	if _, err := ovba.EncodeMBCS(codeOnly, codePage); err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "codeOnly is not representable in the project code page", err)
	}

	generator := UserFormGUIDGenerator(randomUserFormGUID)
	if len(options) == 1 && options[0].GUIDGenerator != nil {
		generator = options[0].GUIDGenerator
	}
	first, err := generatedUserFormGUID(generator)
	if err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "generate first component GUID", err)
	}
	second, err := generatedUserFormGUID(generator)
	if err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "generate second component GUID", err)
	}
	if first == second {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "component GUIDs must be unique", nil)
	}

	source := userFormModuleHeader(designer.Name, first, second) + userFormCodeBody(codeOnly)
	if _, err := ovba.EncodeMBCS(source, codePage); err != nil {
		return Module{}, userFormGenerationError(UserFormGenerationInvalid, designer.Name, "generated module source is not representable in the project code page", err)
	}
	return Module{Name: designer.Name, StreamName: designer.Name, Type: ModuleForm, Source: source}, nil
}

// WithNewUserForm returns a project containing one new UserForm Designer and
// its code-behind module. All validation happens before the input project or
// form is touched. The project must already contain the MSForms reference;
// reference mutation belongs to the separate reference-mutation boundary.
func WithNewUserForm(p *Project, form *oforms.Form, module Module) (*Project, error) {
	if p == nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, "", "project is nil", nil)
	}
	if form == nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, "", "form is nil", nil)
	}
	if p.Protection.IsProtected {
		return nil, userFormGenerationError(UserFormGenerationInvalid, form.Name, "protected projects cannot add UserForms", nil)
	}
	if !hasMSFormsReference(p) {
		return nil, userFormGenerationError(UserFormFormsReferenceRequired, form.Name, "project does not contain an MSForms reference", nil)
	}
	if err := validateNewUserFormInputs(p, form, module); err != nil {
		return nil, err
	}
	serialized, err := oforms.SerializeForm(form, p.Props.CodePage)
	if err != nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, form.Name, "designer is not a valid lossless form", err)
	}
	projectStream, err := addUserFormProjectComponent(p.ProjectStreamRaw, form.Name, p.Props.CodePage)
	if err != nil {
		return nil, err
	}
	clonedForm, err := cloneSerializedUserForm(serialized, form.Name, p.Props.CodePage)
	if err != nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, form.Name, "clone Designer for project result", err)
	}

	result := cloneProject(p)
	result.ProjectStreamRaw = projectStream
	result.Modules = append(result.Modules, module)
	result.Forms = append(result.Forms, clonedForm)
	result.rebuildProjectWM = true
	return result, nil
}

func validateNewUserFormInputs(p *Project, form *oforms.Form, module Module) error {
	if p.Props.CodePage == 0 {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "project code page is zero", nil)
	}
	if err := validateUserFormName(form.Name, p.Props.CodePage); err != nil {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "invalid form identity", err)
	}
	if module.Type != ModuleForm {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "module type is not ModuleForm", nil)
	}
	if module.Name != form.Name || module.StreamName != form.Name {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "module and Designer identities do not match", nil)
	}
	if err := ValidateWritableComponentIdentity(module.Name, module.StreamName, p.Props.CodePage); err != nil {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "invalid module identity", err)
	}
	if _, err := ovba.EncodeMBCS(module.Source, p.Props.CodePage); err != nil {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "module source is not representable in the project code page", err)
	}
	if err := validateUserFormModuleSource(module, form.Name); err != nil {
		return userFormGenerationError(UserFormGenerationInvalid, form.Name, "module attributes do not match a new UserForm", err)
	}
	if moduleNameConflict(p, module.Name) || moduleStreamConflict(p, module.StreamName) || formStorageConflict(p, form.Name) {
		return userFormGenerationError(UserFormGenerationConflict, form.Name, "module or Designer identity collides with the project", nil)
	}
	return nil
}

func validateUserFormName(name string, codePage uint16) error {
	if name == "" {
		return fmt.Errorf("form name is empty")
	}
	if strings.ContainsAny(name, "/\x00\r\n\"{}") {
		return fmt.Errorf("form name %q contains a path, quote, brace, NUL, or newline", name)
	}
	switch strings.ToLower(name) {
	case "vba", "project", "projectwm", "root entry":
		return fmt.Errorf("form name %q is reserved at the CFB root", name)
	}
	return ValidateWritableComponentIdentity(name, name, codePage)
}

func validateUserFormModuleSource(module Module, formName string) error {
	lines := strings.Split(toCRLF(module.Source), "\r\n")
	const attributeCount = 8
	if len(lines) < attributeCount {
		return fmt.Errorf("module source has fewer than %d required attributes", attributeCount)
	}
	if lines[0] != `Attribute VB_Name = "`+formName+`"` {
		return fmt.Errorf("attribute VB_Name does not match %q", formName)
	}
	if !validUserFormBaseAttribute(lines[1]) {
		return fmt.Errorf("attribute VB_Base is malformed")
	}
	want := []string{
		"Attribute VB_GlobalNameSpace = False",
		"Attribute VB_Creatable = False",
		"Attribute VB_PredeclaredId = True",
		"Attribute VB_Exposed = False",
		"Attribute VB_TemplateDerived = False",
		"Attribute VB_Customizable = False",
	}
	for i, expected := range want {
		if lines[i+2] != expected {
			return fmt.Errorf("attribute %q is missing or mismatched", expected)
		}
	}
	for _, line := range lines[attributeCount:] {
		if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
			return fmt.Errorf("module source contains an extra Attribute VB_ line")
		}
	}
	return nil
}

func validUserFormBaseAttribute(line string) bool {
	const prefix = `Attribute VB_Base = "0{`
	if !strings.HasPrefix(line, prefix) || !strings.HasSuffix(line, `}"`) {
		return false
	}
	value := strings.TrimSuffix(strings.TrimPrefix(line, prefix), `}"`)
	close := strings.IndexByte(value, '}')
	if close != 36 || len(value) != 74 || value[close+1] != '{' {
		return false
	}
	return validCanonicalGUID(value[:close]) && validCanonicalGUID(value[close+2:])
}

func validCanonicalGUID(value string) bool {
	if len(value) != 36 || value[8] != '-' || value[13] != '-' || value[18] != '-' || value[23] != '-' {
		return false
	}
	for i, b := range []byte(value) {
		if i == 8 || i == 13 || i == 18 || i == 23 {
			continue
		}
		if !isHexDigit(b) {
			return false
		}
	}
	return true
}

func isHexDigit(b byte) bool {
	return b >= '0' && b <= '9' || b >= 'a' && b <= 'f' || b >= 'A' && b <= 'F'
}

func hasUserFormAttributeLine(codeOnly string) bool {
	for line := range strings.SplitSeq(toCRLF(codeOnly), "\r\n") {
		if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
			return true
		}
	}
	return false
}

func userFormModuleHeader(name, firstGUID, secondGUID string) string {
	return strings.Join([]string{
		`Attribute VB_Name = "` + name + `"`,
		`Attribute VB_Base = "0{` + firstGUID + `}{` + secondGUID + `}"`,
		"Attribute VB_GlobalNameSpace = False",
		"Attribute VB_Creatable = False",
		"Attribute VB_PredeclaredId = True",
		"Attribute VB_Exposed = False",
		"Attribute VB_TemplateDerived = False",
		"Attribute VB_Customizable = False",
	}, "\r\n") + "\r\n"
}

func userFormCodeBody(codeOnly string) string {
	body := toCRLF(codeOnly)
	if body != "" && !strings.HasSuffix(body, "\r\n") {
		body += "\r\n"
	}
	return body
}

func randomUserFormGUID() (string, error) {
	var raw [16]byte
	if _, err := rand.Read(raw[:]); err != nil {
		return "", err
	}
	// RFC 4122 version 4 and variant bits make the persisted identity a
	// conventional UUID while the entropy still comes only from crypto/rand.
	raw[6] = raw[6]&0x0f | 0x40
	raw[8] = raw[8]&0x3f | 0x80
	return formatUserFormGUID(raw), nil
}

func formatUserFormGUID(raw [16]byte) string {
	var encoded [36]byte
	hex.Encode(encoded[0:8], raw[0:4])
	encoded[8] = '-'
	hex.Encode(encoded[9:13], raw[4:6])
	encoded[13] = '-'
	hex.Encode(encoded[14:18], raw[6:8])
	encoded[18] = '-'
	hex.Encode(encoded[19:23], raw[8:10])
	encoded[23] = '-'
	hex.Encode(encoded[24:36], raw[10:16])
	return strings.ToUpper(string(encoded[:]))
}

func generatedUserFormGUID(generator UserFormGUIDGenerator) (string, error) {
	value, err := generator()
	if err != nil {
		return "", err
	}
	if !validCanonicalGUID(value) {
		return "", fmt.Errorf("GUID %q is not canonical", value)
	}
	return strings.ToUpper(value), nil
}

func hasMSFormsReference(p *Project) bool {
	if p == nil {
		return false
	}
	return hasMSFormsReferenceRaw(p.ReferencesRaw)
}

func hasMSFormsReferenceRaw(raw []byte) bool {
	if len(raw) == 0 {
		return false
	}

	// References is display metadata only. The dir reference records are the
	// authoritative source. Excel-authored Forms references use either a
	// canonical GUID in the original-reference record or a REFERENCECONTROL
	// record whose twiddled LIBID proves the Forms reference.
	msFormsName := false
	for offset := 0; offset+6 <= len(raw); {
		id := binary.LittleEndian.Uint16(raw[offset:])
		size := binary.LittleEndian.Uint32(raw[offset+2:])
		if size > uint32(len(raw)-offset-6) {
			return false
		}
		end := offset + 6 + int(size)
		payload := raw[offset+6 : end]
		switch id {
		case 0x0016: // REFERENCENAME
			msFormsName = bytes.EqualFold(payload, []byte("MSForms"))
		case 0x0033: // REFERENCEORIGINAL
			if hasMSFormsGUID(payload) {
				return true
			}
		case 0x002F, 0x0030: // REFERENCECONTROL variants
			if hasMSFormsGUID(payload) || msFormsName && hasTwiddledFormsLIBID(payload) {
				return true
			}
		}
		offset = end
	}
	return false
}

func hasMSFormsGUID(payload []byte) bool {
	rawUpper := bytes.ToUpper(payload)
	return bytes.Contains(rawUpper, []byte(userFormMSFormsReferenceGUID)) ||
		bytes.Contains(payload, userFormMSFormsReferenceGUIDWire[:])
}

func hasTwiddledFormsLIBID(payload []byte) bool {
	return bytes.Contains(payload, []byte(`*\G{`))
}

func moduleNameConflict(p *Project, name string) bool {
	key := strings.ToLower(name)
	for _, module := range p.Modules {
		if strings.ToLower(module.Name) == key {
			return true
		}
	}
	parsed, err := ovba.ParseProjectText(p.ProjectStreamRaw, p.Props.CodePage)
	if err != nil {
		return false
	}
	for _, component := range parsed.Components {
		if strings.EqualFold(component.Name, name) {
			return true
		}
	}
	return false
}

func moduleStreamConflict(p *Project, streamName string) bool {
	key := cfb.DirectoryNameKey(streamName)
	for _, module := range p.Modules {
		if cfb.DirectoryNameKey(module.StreamName) == key {
			return true
		}
	}
	for path := range p.RawStreams {
		parts := strings.Split(path, "/")
		if len(parts) == 2 && cfb.DirectoryNameKey(parts[0]) == cfb.DirectoryNameKey("VBA") && cfb.DirectoryNameKey(parts[1]) == key {
			return true
		}
	}
	return key == cfb.DirectoryNameKey("dir") || key == cfb.DirectoryNameKey("_VBA_PROJECT")
}

func formStorageConflict(p *Project, name string) bool {
	key := cfb.DirectoryNameKey(name)
	for _, form := range p.Forms {
		if form != nil && cfb.DirectoryNameKey(form.Name) == key {
			return true
		}
	}
	for path := range p.StorageMetadata {
		if path != "" && cfb.DirectoryNameKey(firstSegment(path)) == key {
			return true
		}
	}
	for path := range p.RawStreams {
		if cfb.DirectoryNameKey(firstSegment(path)) == key {
			return true
		}
	}
	return false
}

func addUserFormProjectComponent(raw []byte, name string, codePage uint16) ([]byte, error) {
	parsed, err := ovba.ParseProjectText(raw, codePage)
	if err != nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, name, "PROJECT stream cannot be parsed", err)
	}
	for _, component := range parsed.Components {
		if strings.EqualFold(component.Name, name) {
			return nil, userFormGenerationError(UserFormGenerationConflict, name, "PROJECT component already exists", nil)
		}
	}
	encoded, err := ovba.EncodeMBCS(name, codePage)
	if err != nil {
		return nil, userFormGenerationError(UserFormGenerationInvalid, name, "PROJECT component is not representable in the project code page", err)
	}

	lines := splitProjectLines(raw)
	insertAfter := -1
	lastComponent := -1
	eol := []byte("\r\n")
	for i, line := range lines {
		if len(line.eol) != 0 {
			eol = line.eol
		}
		if isProjectComponentLine(line.body) {
			lastComponent = i
		}
		if bytes.HasPrefix(line.body, []byte("ID=")) {
			insertAfter = i
		}
	}
	if lastComponent >= 0 {
		insertAfter = lastComponent
	}
	addition := append([]byte("BaseClass="), encoded...)
	addition = append(addition, eol...)
	if insertAfter < 0 {
		return append(addition, raw...), nil
	}

	var out bytes.Buffer
	for i, line := range lines {
		out.Write(line.body)
		out.Write(line.eol)
		if i != insertAfter {
			continue
		}
		if len(line.eol) == 0 {
			out.Write(eol)
		}
		out.Write(addition)
	}
	return out.Bytes(), nil
}

type projectLine struct {
	body []byte
	eol  []byte
}

func splitProjectLines(raw []byte) []projectLine {
	var lines []projectLine
	for len(raw) > 0 {
		index := bytes.IndexByte(raw, '\n')
		if index < 0 {
			lines = append(lines, projectLine{body: raw})
			break
		}
		body, eol := raw[:index], []byte("\n")
		if bytes.HasSuffix(body, []byte("\r")) {
			body = body[:len(body)-1]
			eol = []byte("\r\n")
		}
		lines = append(lines, projectLine{body: body, eol: eol})
		raw = raw[index+1:]
	}
	return lines
}

func isProjectComponentLine(line []byte) bool {
	for _, prefix := range []string{"Module=", "Class=", "Document=", "BaseClass="} {
		if bytes.HasPrefix(line, []byte(prefix)) {
			return true
		}
	}
	return false
}

func cloneSerializedUserForm(serialized *oforms.SerializedForm, root string, codePage uint16) (*oforms.Form, error) {
	writer := cfb.NewWriter()
	for _, path := range slices.Sorted(maps.Keys(serialized.Storages)) {
		parts := []string(nil)
		if path != "" {
			parts = strings.Split(path, "/")
		}
		writer.AddStorage(parts, serialized.Storages[path])
	}
	for _, path := range slices.Sorted(maps.Keys(serialized.Streams)) {
		writer.AddStream(strings.Split(path, "/"), serialized.Streams[path])
	}
	body, err := writer.Bytes()
	if err != nil {
		return nil, err
	}
	container, err := cfb.Open(body)
	if err != nil {
		return nil, err
	}
	return oforms.ReadForm(container, root, codePage)
}

func cloneProject(p *Project) *Project {
	result := *p
	result.Modules = slices.Clone(p.Modules)
	for i := range result.Modules {
		result.Modules[i].Extra = slices.Clone(result.Modules[i].Extra)
		for j := range result.Modules[i].Extra {
			result.Modules[i].Extra[j].Payload = bytes.Clone(result.Modules[i].Extra[j].Payload)
		}
	}
	result.References = slices.Clone(p.References)
	result.ReferencesRaw = bytes.Clone(p.ReferencesRaw)
	result.ProjectInfoRaw = bytes.Clone(p.ProjectInfoRaw)
	result.ProjectStreamRaw = bytes.Clone(p.ProjectStreamRaw)
	result.Forms = slices.Clone(p.Forms)
	result.RawStreams = cloneByteMap(p.RawStreams)
	result.StorageMetadata = maps.Clone(p.StorageMetadata)
	return &result
}

func cloneByteMap(input map[string][]byte) map[string][]byte {
	if input == nil {
		return nil
	}
	result := make(map[string][]byte, len(input))
	for key, value := range input {
		result[key] = bytes.Clone(value)
	}
	return result
}

func userFormGenerationError(code, form, reason string, cause error) *UserFormGenerationError {
	return &UserFormGenerationError{Code: code, Form: form, Reason: reason, Cause: cause}
}
