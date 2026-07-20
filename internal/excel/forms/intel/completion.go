package intel

import (
	"fmt"
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/excel/forms"
)

// CompletionKind is intentionally independent of LSP completion kinds. The
// LSP adapter is responsible for translating these stable editor concepts.
type CompletionKind string

const (
	CompletionKindProperty  CompletionKind = "property"
	CompletionKindValue     CompletionKind = "value"
	CompletionKindReference CompletionKind = "reference"
	CompletionKindSnippet   CompletionKind = "snippet"
)

// CompletionItem is a completion suggestion for a UserForm YAML document.
// Replace is the source range that should be replaced by InsertText. Snippet
// indicates that InsertText uses LSP snippet syntax.
type CompletionItem struct {
	Label         string
	Detail        string
	Documentation string
	InsertText    string
	Replace       Range
	Kind          CompletionKind
	Snippet       bool
	SortText      string
}

// CompleteYAML returns context-aware completions for a UserForm YAML source
// buffer. It accepts incomplete YAML; ParseYAML's cursor fallback supplies the
// structural context when a syntax tree cannot be constructed.
func CompleteYAML(source string, pos Position) []CompletionItem {
	doc := ParseYAML(source)
	return CompleteDocument(doc, pos)
}

// CompleteDocument returns completion suggestions using a previously parsed
// YAML document. Reusing Document avoids reparsing in editor request paths.
func CompleteDocument(doc *Document, pos Position) []CompletionItem {
	if doc == nil {
		return nil
	}
	if strings.TrimSpace(doc.Source) == "" {
		return []CompletionItem{documentSkeleton(pos)}
	}

	context := doc.CursorContextAt(pos)
	line := doc.line(pos.Line)
	mode := completionLineContext(line, pos)
	if !mode.value && strings.TrimSpace(mode.prefix) != "" {
		// A syntactically valid YAML scalar can still be structurally incomplete
		// for this contract (`form:\n  ob`). The node-based lookup then treats
		// the cursor as the scalar value of form; the indentation fallback retains
		// the mapping context needed for property completion.
		fallback := fallbackCursorContext(doc.lines, pos)
		if fallback.ParentPath != "" || context.ParentPath == "" {
			context = fallback
		}
	}
	if mode.value {
		return valueCompletions(doc, context, mode, pos)
	}

	parent := context.ParentPath
	if mode.fieldPath != "" {
		parent = parentPath(mode.fieldPath)
	}
	items := propertyCompletions(doc, parent, context, mode, pos)
	if isControlPath(parent) && len(context.ExistingKeys) == 0 {
		items = append(items, controlSnippets(line, mode.Replace)...)
	}
	if isControlsSequencePath(parent) {
		items = append(items, controlSnippets(line, mode.Replace)...)
	}
	return filterPrefixAndSort(items, mode.prefix)
}

type lineCompletionContext struct {
	key       string
	prefix    string
	fieldPath string
	value     bool
	Replace   Range
}

func completionLineContext(line string, pos Position) lineCompletionContext {
	lineIndex := pos.Line
	bytePos := byteOffsetForUTF16(line, pos.Character)
	bytePos = min(bytePos, len(line))
	start := len(line) - len(strings.TrimLeft(line, " \t"))
	if start < len(line) && line[start] == '-' {
		start++
		for start < len(line) && (line[start] == ' ' || line[start] == '\t') {
			start++
		}
	}
	colon := strings.IndexByte(line[start:], ':')
	if colon >= 0 {
		colon += start
		key := strings.TrimSpace(line[start:colon])
		keyStart := start
		for keyStart < colon && (line[keyStart] == ' ' || line[keyStart] == '\t') {
			keyStart++
		}
		if bytePos <= colon {
			return lineCompletionContext{key: key, prefix: strings.TrimSpace(line[keyStart:bytePos]), Replace: Range{Start: positionForByte(line, lineIndex, keyStart), End: positionForByte(line, lineIndex, bytePos)}}
		}
		valueStart := colon + 1
		for valueStart < bytePos && (line[valueStart] == ' ' || line[valueStart] == '\t') {
			valueStart++
		}
		if valueStart < len(line) && (line[valueStart] == '\'' || line[valueStart] == '"') {
			// Keep YAML quotes outside the replacement range. This lets a
			// completion filter against the scalar contents and preserves a
			// closing quote when completion is requested inside a quoted value.
			quoteStart := valueStart
			quote := line[quoteStart]
			valueStart++
			valueEnd := bytePos
			if quotedEnd := quotedEnd(line, quoteStart); quotedEnd <= bytePos && quotedEnd > quoteStart+1 && line[quotedEnd-1] == quote {
				valueEnd = quotedEnd - 1
			}
			return lineCompletionContext{key: key, prefix: strings.TrimSpace(line[valueStart:valueEnd]), value: true, Replace: Range{Start: positionForByte(line, lineIndex, valueStart), End: positionForByte(line, lineIndex, valueEnd)}}
		}
		return lineCompletionContext{key: key, prefix: strings.TrimSpace(line[valueStart:bytePos]), value: true, Replace: Range{Start: positionForByte(line, lineIndex, valueStart), End: positionForByte(line, lineIndex, bytePos)}}
	}
	if bytePos < start {
		start = bytePos
	}
	return lineCompletionContext{prefix: strings.TrimSpace(line[start:bytePos]), Replace: Range{Start: positionForByte(line, lineIndex, start), End: positionForByte(line, lineIndex, bytePos)}}
}

func propertyCompletions(doc *Document, parent string, context CursorContext, line lineCompletionContext, pos Position) []CompletionItem {
	contract := forms.UserFormContract()
	var properties map[string]forms.PropertyContract
	switch {
	case parent == "":
		properties = contract.DocumentProperties
	case parent == "form":
		properties = contract.FormProperties
	case isControlPath(parent):
		properties = controlProperties(contract, controlAtPath(doc.Source, parent))
	default:
		return nil
	}

	existing := make(map[string]bool, len(context.ExistingKeys))
	for _, key := range context.ExistingKeys {
		existing[strings.ToLower(key)] = true
	}
	items := make([]CompletionItem, 0, len(properties))
	for name, property := range properties {
		if existing[strings.ToLower(name)] && !strings.EqualFold(name, line.key) {
			continue
		}
		if !offerProperty(name, property, parent, line.prefix) {
			continue
		}
		items = append(items, CompletionItem{
			Label: name, InsertText: name + ": ", Replace: line.Replace, Kind: CompletionKindProperty,
			Detail: propertyDetail(property), Documentation: propertyDocumentation(name, property), SortText: propertySortText(property),
		})
	}
	return items
}

func controlProperties(contract forms.Contract, control controlInfo) map[string]forms.PropertyContract {
	properties := make(map[string]forms.PropertyContract, len(contract.CommonControlProperties)+4)
	for name, property := range contract.CommonControlProperties {
		properties[name] = property
	}
	if control.isKnownType() && !control.isCustomProgID() {
		if typed, ok := contract.Controls[strings.ToLower(control.Type)]; ok {
			for name, property := range typed.Properties {
				properties[name] = property
			}
		}
	}
	return properties
}

func offerProperty(name string, property forms.PropertyContract, parent, prefix string) bool {
	if property.IncludeInAuthoring {
		return true
	}
	// Structural and explicitly authorable contract fields are normal candidates.
	if parent == "" && name != "warnings" || parent == "form" && name == "build" || isControlPath(parent) && name == "progId" {
		return true
	}
	// Snapshot and unchecked escape-hatch fields remain discoverable only after a
	// user begins typing their name, where their SortText ranks them last.
	return strings.TrimSpace(prefix) != ""
}

func valueCompletions(doc *Document, context CursorContext, line lineCompletionContext, pos Position) []CompletionItem {
	parent := context.ParentPath
	if context.FieldPath != "" && strings.EqualFold(lastPathSegment(context.FieldPath), line.key) {
		parent = parentPath(context.FieldPath)
	} else if strings.EqualFold(lastPathSegment(parent), line.key) {
		// The tolerant fallback represents a blank `key:` value as a child
		// mapping. For scalar completion its parent is the mapping above it.
		parent = parentPath(parent)
	}
	property, ok := propertyAt(doc.Source, parent, line.key)
	if !ok {
		return nil
	}
	if isControlPath(parent) {
		control := controlAtPath(doc.Source, parent)
		switch strings.ToLower(line.key) {
		case "type":
			return filterPrefixAndSort(controlTypeValues(line.Replace), line.prefix)
		case "progid":
			return filterPrefixAndSort(progIDValues(control, line.Replace), line.prefix)
		case "parentid":
			return filterPrefixAndSort(parentIDValues(doc.Source, control, line.Replace), line.prefix)
		}
	}
	if len(property.AllowedValues) != 0 {
		items := make([]CompletionItem, 0, len(property.AllowedValues))
		for _, value := range property.AllowedValues {
			items = append(items, scalarValue(value, property.Description, line.Replace))
		}
		return filterPrefixAndSort(items, line.prefix)
	}
	if property.ValueType == forms.ValueTypeBoolean {
		return filterPrefixAndSort([]CompletionItem{scalarValue("true", "Boolean true.", line.Replace), scalarValue("false", "Boolean false.", line.Replace)}, line.prefix)
	}
	return nil
}

func propertyAt(source, parent, name string) (forms.PropertyContract, bool) {
	contract := forms.UserFormContract()
	if parent == "" {
		return lookupProperty(contract.DocumentProperties, name)
	}
	if parent == "form" {
		return lookupProperty(contract.FormProperties, name)
	}
	if !isControlPath(parent) {
		return forms.PropertyContract{}, false
	}
	control := controlAtPath(source, parent)
	if property, ok := lookupProperty(contract.CommonControlProperties, name); ok {
		return property, true
	}
	if control.isKnownType() && !control.isCustomProgID() {
		return forms.LookupControlProperty(control.Type, name)
	}
	return forms.PropertyContract{}, false
}

func lookupProperty(properties map[string]forms.PropertyContract, name string) (forms.PropertyContract, bool) {
	for key, property := range properties {
		if strings.EqualFold(key, name) {
			return property, true
		}
	}
	return forms.PropertyContract{}, false
}

func controlTypeValues(replace Range) []CompletionItem {
	contract := forms.UserFormContract()
	controls := make([]forms.ControlContract, 0, len(contract.Controls))
	for _, control := range contract.Controls {
		controls = append(controls, control)
	}
	sort.Slice(controls, func(i, j int) bool { return controls[i].Type < controls[j].Type })
	items := make([]CompletionItem, 0, len(controls))
	for _, control := range controls {
		items = append(items, CompletionItem{Label: control.Type, InsertText: control.Type, Replace: replace, Kind: CompletionKindValue, Detail: control.ProgID, Documentation: fmt.Sprintf("**%s** control.  \nProgID: `%s`.", control.Type, control.ProgID)})
	}
	return items
}

func progIDValues(control controlInfo, replace Range) []CompletionItem {
	if control.isKnownType() {
		if progID, ok := forms.BuiltInControlProgID(control.Type); ok {
			return []CompletionItem{{Label: progID, InsertText: progID, Replace: replace, Kind: CompletionKindValue, Detail: control.Type, Documentation: fmt.Sprintf("Built-in ProgID for **%s**.", control.Type)}}
		}
	}
	contract := forms.UserFormContract()
	return controlProgIDs(contract, replace)
}

func controlProgIDs(contract forms.Contract, replace Range) []CompletionItem {
	controls := make([]forms.ControlContract, 0, len(contract.Controls))
	for _, control := range contract.Controls {
		controls = append(controls, control)
	}
	sort.Slice(controls, func(i, j int) bool { return controls[i].Type < controls[j].Type })
	items := make([]CompletionItem, 0, len(controls))
	for _, control := range controls {
		items = append(items, CompletionItem{Label: control.ProgID, InsertText: control.ProgID, Replace: replace, Kind: CompletionKindValue, Detail: control.Type, Documentation: fmt.Sprintf("Built-in ProgID for **%s**.", control.Type)})
	}
	return items
}

func scalarValue(value, documentation string, replace Range) CompletionItem {
	return CompletionItem{Label: value, InsertText: value, Replace: replace, Kind: CompletionKindValue, Documentation: documentation}
}

func propertyDetail(property forms.PropertyContract) string {
	detail := string(property.ValueType)
	if property.Required {
		detail += " · required"
	}
	if property.SupportLevel != forms.SupportLevelSupported {
		detail += " · " + string(property.SupportLevel)
	}
	return detail
}

func propertyDocumentation(name string, property forms.PropertyContract) string {
	parts := []string{fmt.Sprintf("**%s** (`%s`)", name, property.ValueType)}
	if property.Required {
		parts = append(parts, "Required.")
	}
	if len(property.ApplicableControls) != 0 {
		parts = append(parts, "Applies to: "+strings.Join(property.ApplicableControls, ", ")+".")
	}
	parts = append(parts, "Support: **"+string(property.SupportLevel)+"**.", property.Description)
	if property.SupportLevel != forms.SupportLevelSupported {
		parts = append(parts, "This field is not a fully supported authoring property.")
	}
	return strings.Join(parts, "  \n")
}

func propertySortText(property forms.PropertyContract) string {
	if property.SupportLevel == forms.SupportLevelSnapshotOnly || property.SupportLevel == forms.SupportLevelCustomUnchecked {
		return "z"
	}
	return "a"
}

func filterPrefixAndSort(items []CompletionItem, prefix string) []CompletionItem {
	prefix = strings.ToLower(strings.TrimSpace(prefix))
	filtered := items[:0]
	for _, item := range items {
		if prefix == "" || strings.HasPrefix(strings.ToLower(item.Label), prefix) {
			filtered = append(filtered, item)
		}
	}
	sort.SliceStable(filtered, func(i, j int) bool {
		if filtered[i].SortText != filtered[j].SortText {
			return filtered[i].SortText < filtered[j].SortText
		}
		return strings.ToLower(filtered[i].Label) < strings.ToLower(filtered[j].Label)
	})
	return filtered
}

func documentSkeleton(pos Position) CompletionItem {
	return CompletionItem{Label: "UserForm document", Detail: "UserForm YAML skeleton", Kind: CompletionKindSnippet, Snippet: true,
		InsertText: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\ncoordinateSystem: points\nform:\n  name: ${1:UserFormName}\ncontrols:\n  ${0}", Replace: Range{Start: pos, End: pos}}
}

func controlSnippets(line string, replace Range) []CompletionItem {
	prefix := "- "
	if strings.HasPrefix(strings.TrimLeft(line, " \t"), "-") {
		// The sequence dash already exists in the document, so inserting another
		// one would turn the current item into a nested sequence.
		prefix = ""
	}
	items := []CompletionItem{
		controlSnippet("Basic control", prefix+"id: ${1:control_id}\n  name: ${2:ControlName}\n  type: ${3|Label,TextBox,ComboBox,ListBox,CommandButton,CheckBox,OptionButton,Frame|}\n  left: ${4:0}\n  top: ${5:0}\n  width: ${6:100}\n  height: ${7:20}", replace),
		controlSnippet("Frame", prefix+"id: ${1:frame_id}\n  name: ${2:FrameName}\n  type: Frame\n  caption: ${3:Frame}\n  left: ${4:0}\n  top: ${5:0}\n  width: ${6:100}\n  height: ${7:80}", replace),
		controlSnippet("Label", prefix+"id: ${1:label_id}\n  name: ${2:LabelName}\n  type: Label\n  caption: ${3:Label}\n  left: ${4:0}\n  top: ${5:0}\n  width: ${6:100}\n  height: ${7:20}", replace),
		controlSnippet("TextBox", prefix+"id: ${1:textbox_id}\n  name: ${2:TextBoxName}\n  type: TextBox\n  text: ${3:}\n  left: ${4:0}\n  top: ${5:0}\n  width: ${6:100}\n  height: ${7:20}", replace),
		controlSnippet("CommandButton", prefix+"id: ${1:button_id}\n  name: ${2:CommandButtonName}\n  type: CommandButton\n  caption: ${3:OK}\n  left: ${4:0}\n  top: ${5:0}\n  width: ${6:100}\n  height: ${7:20}", replace),
	}
	indent := line[:len(line)-len(strings.TrimLeft(line, " \t"))]
	for i := range items {
		items[i].InsertText = strings.ReplaceAll(items[i].InsertText, "\n  ", "\n"+indent+"  ")
	}
	return items
}

func controlSnippet(label, insertText string, replace Range) CompletionItem {
	return CompletionItem{Label: label, Detail: "UserForm control snippet", InsertText: insertText, Replace: replace, Kind: CompletionKindSnippet, Snippet: true, SortText: "b"}
}

func isControlPath(path string) bool { return controlPathPattern.MatchString(path) }

func isControlsSequencePath(path string) bool {
	return path == "controls" || strings.HasSuffix(path, ".controls")
}

var controlPathPattern = regexp.MustCompile(`^controls\[\d+\](?:\.controls\[\d+\])*$`)

func lastPathSegment(path string) string {
	if index := strings.LastIndex(path, "."); index >= 0 {
		return path[index+1:]
	}
	return path
}

type controlInfo struct {
	path     string
	id       string
	name     string
	Type     string
	progID   string
	parentID string
	order    int
}

func (control controlInfo) isKnownType() bool {
	_, ok := forms.LookupControlContract(control.Type)
	return ok
}

func (control controlInfo) isCustomProgID() bool {
	if strings.TrimSpace(control.progID) == "" {
		return false
	}
	_, known := forms.LookupControlContractByProgID(control.progID)
	return !known
}

// controlsInSource deliberately scans source text rather than decoded structs.
// That keeps reference completion useful while a key or scalar is half typed.
func controlsInSource(source string) []controlInfo {
	type block struct {
		indent int
		path   string
		next   int
	}
	type frame struct {
		indent int
		path   string
	}
	lines := splitLines(source)
	blocks := []block{}
	frames := []frame{}
	controls := []controlInfo{}
	byPath := map[string]int{}
	for _, raw := range lines {
		indent := len(raw) - len(strings.TrimLeft(raw, " \t"))
		trimmed := strings.TrimSpace(raw)
		if trimmed == "" || strings.HasPrefix(trimmed, "#") {
			continue
		}
		for len(frames) > 0 && frames[len(frames)-1].indent >= indent && !strings.HasPrefix(trimmed, "-") {
			frames = frames[:len(frames)-1]
		}
		if strings.HasPrefix(trimmed, "-") && (len(trimmed) == 1 || trimmed[1] == ' ') {
			for len(frames) > 0 && frames[len(frames)-1].indent >= indent {
				frames = frames[:len(frames)-1]
			}
			blockIndex := -1
			for i := len(blocks) - 1; i >= 0; i-- {
				// YAML permits an indentationless sequence directly below its
				// mapping key (`controls:\n- ...`), so the owner may be at the
				// same indentation as this sequence entry.
				if blocks[i].indent <= indent {
					blockIndex = i
					break
				}
			}
			if blockIndex >= 0 {
				itemPath := fmt.Sprintf("%s[%d]", blocks[blockIndex].path, blocks[blockIndex].next)
				blocks[blockIndex].next++
				controls = append(controls, controlInfo{path: itemPath, order: len(controls)})
				byPath[itemPath] = len(controls) - 1
				frames = append(frames, frame{indent: indent, path: itemPath})
				trimmed = strings.TrimSpace(strings.TrimPrefix(trimmed, "-"))
			}
		}
		key, hasColon := yamlKey(trimmed)
		if !hasColon {
			continue
		}
		value := strings.TrimSpace(strings.TrimPrefix(trimmed[len(key):], ":"))
		if key == "controls" && value == "" {
			owner := ""
			if len(frames) > 0 && frames[len(frames)-1].indent < indent {
				owner = frames[len(frames)-1].path
			}
			path := "controls"
			if owner != "" {
				path = owner + ".controls"
			}
			blocks = append(blocks, block{indent: indent, path: path})
			continue
		}
		if len(frames) == 0 {
			continue
		}
		current := frames[len(frames)-1]
		if current.indent > indent {
			continue
		}
		index, ok := byPath[current.path]
		if !ok {
			continue
		}
		switch strings.ToLower(key) {
		case "id":
			controls[index].id = unquoteYAMLScalar(value)
		case "name":
			controls[index].name = unquoteYAMLScalar(value)
		case "type":
			controls[index].Type = unquoteYAMLScalar(value)
		case "progid":
			controls[index].progID = unquoteYAMLScalar(value)
		case "parentid":
			controls[index].parentID = unquoteYAMLScalar(value)
		}
	}
	return controls
}

func unquoteYAMLScalar(value string) string {
	value = strings.TrimSpace(value)
	if len(value) >= 2 && ((value[0] == '"' && value[len(value)-1] == '"') || (value[0] == '\'' && value[len(value)-1] == '\'')) {
		return value[1 : len(value)-1]
	}
	if comment := strings.IndexByte(value, '#'); comment >= 0 {
		value = strings.TrimSpace(value[:comment])
	}
	return value
}

func controlAtPath(source, path string) controlInfo {
	for _, control := range controlsInSource(source) {
		if control.path == path {
			return control
		}
	}
	return controlInfo{path: path}
}

func parentIDValues(source string, current controlInfo, replace Range) []CompletionItem {
	controls := controlsInSource(source)
	idCounts := map[string]int{}
	byID := map[string]controlInfo{}
	for _, candidate := range controls {
		if candidate.id == "" {
			continue
		}
		idCounts[candidate.id]++
		byID[candidate.id] = candidate
	}
	descendants := descendantsOf(current.id, byID)
	items := make([]CompletionItem, 0, len(controls))
	for _, candidate := range controls {
		if candidate.id == "" || idCounts[candidate.id] != 1 || candidate.path == current.path || candidate.id == current.id || descendants[candidate.id] {
			continue
		}
		canContain, known := containerEligibility(candidate)
		if known && !canContain {
			continue
		}
		if !known && strings.TrimSpace(candidate.progID) == "" {
			// A type-only unknown control is invalid under the shared contract;
			// only explicit custom ProgIDs retain unchecked parent compatibility.
			continue
		}
		detailType := candidate.Type
		if detailType == "" {
			detailType = candidate.progID
		}
		detail := strings.TrimSpace(candidate.name)
		if detail == "" {
			detail = candidate.id
		}
		detail += " (" + detailType + ")"
		sortText := "b"
		documentation := "Parent control reference."
		if !known {
			detail += " · custom/unchecked"
			documentation += " Container capability is custom/unchecked."
			sortText = "c"
		} else if canContain {
			sortText = "a"
		}
		items = append(items, CompletionItem{Label: candidate.id, InsertText: candidate.id, Replace: replace, Kind: CompletionKindReference, Detail: detail, Documentation: documentation, SortText: sortText})
	}
	return items
}

func containerEligibility(control controlInfo) (bool, bool) {
	if control.progID != "" {
		if known, ok := forms.LookupControlContractByProgID(control.progID); ok {
			return known.CanContainChildren, true
		}
		return false, false
	}
	if known, ok := forms.LookupControlContract(control.Type); ok {
		return known.CanContainChildren, true
	}
	return false, false
}

func descendantsOf(id string, byID map[string]controlInfo) map[string]bool {
	descendants := map[string]bool{}
	if id == "" {
		return descendants
	}
	changed := true
	for changed {
		changed = false
		for candidateID, candidate := range byID {
			if candidate.parentID == id || descendants[candidate.parentID] {
				if !descendants[candidateID] {
					descendants[candidateID] = true
					changed = true
				}
			}
		}
	}
	return descendants
}
