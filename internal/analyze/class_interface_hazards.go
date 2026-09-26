package analyze

import (
	"cmp"
	"fmt"
	"slices"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// This file implements the Issue #826 class/interface public-API hazard rules:
// VBA278 public member names containing "_", VBA279 public Enum declarations
// in document modules, VBA280 write-only Property APIs, VBA281 members that
// collide with implemented-interface member names or expose event handlers as
// Public, and VBA282 suspicious self-name access to a predeclared default
// instance. Every rule is opt-in (default disabled) and file-local so it runs
// on both the batch and realtime/LSP analysis surfaces.

// interfaceMemberIndex records the public member names of every class module
// in the analyzed file set plus the per-name module counts used to detect
// ambiguous interface names. A zero value (populated == false) means the
// index was never built, which makes every prefix match unresolved and the
// hazard rules fail open.
type interfaceMemberIndex struct {
	members      map[string]map[string]struct{}
	moduleCounts map[string]int
	populated    bool
}

// buildInterfaceMemberIndex collects the public member names of every class
// module in the file set. Interface members must be public; private members
// are not part of the interface contract. Recovered or conditionally compiled
// symbols are skipped because their IR projection may be incomplete, and a
// class that only contributes such members still counts as resolved: its
// observable interface is what remains after those exclusions.
func buildInterfaceMemberIndex(files []parsedFile) interfaceMemberIndex {
	index := interfaceMemberIndex{
		members:      map[string]map[string]struct{}{},
		moduleCounts: map[string]int{},
		populated:    true,
	}
	for i := range files {
		if module := strings.ToLower(cleanIdentifier(files[i].Module)); module != "" {
			index.moduleCounts[module]++
		}
	}
	for i := range files {
		file := files[i]
		if !strings.EqualFold(strings.TrimSpace(file.ModuleKind), "class") {
			continue
		}
		module := strings.ToLower(cleanIdentifier(file.Module))
		if module == "" {
			continue
		}
		members := index.members[module]
		if members == nil {
			members = map[string]struct{}{}
			index.members[module] = members
		}
		for p := range file.IR.Procedures {
			symbol := file.IR.Procedures[p].Symbol
			if !memberKind(symbol.Kind) || symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
				continue
			}
			if !isPublicMember(symbol) {
				continue
			}
			if name := strings.ToLower(cleanIdentifier(symbol.Name)); name != "" {
				members[name] = struct{}{}
			}
		}
	}
	return index
}

// classInterfaceHazardFindings dispatches the file-local class/interface rules
// (VBA278, VBA279, VBA280, VBA281). Each rule remains individually gated so a
// disabled configuration costs one branch per file.
func (a Analyzer) classInterfaceHazardFindings(file parsedFile, interfaces interfaceMemberIndex) []Finding {
	cfg := a.Config.Analyze
	if !cfg.DetectPublicMemberUnderscoreNames && !cfg.DetectDocumentModulePublicEnum &&
		!cfg.DetectWriteOnlyProperty && !cfg.DetectPublicInterfaceEventMembers {
		return nil
	}
	// VBA220 only pre-computes designer controls for _change/_click-shaped
	// members. The underscore and interface-event rules need the control set
	// for every form member, so resolve the designer artifact lazily when the
	// pipeline did not already populate it.
	if (cfg.DetectPublicMemberUnderscoreNames || cfg.DetectPublicInterfaceEventMembers) &&
		!file.userFormControlsChecked && strings.EqualFold(strings.TrimSpace(file.ModuleKind), "form") {
		file.UserFormControlNames, file.UserFormControlsKnown = a.userFormControlNames(file)
		file.userFormControlsChecked = true
	}
	var findings []Finding
	if cfg.DetectPublicMemberUnderscoreNames {
		findings = append(findings, a.publicMemberUnderscoreFindings(file, interfaces)...)
	}
	if cfg.DetectDocumentModulePublicEnum {
		findings = append(findings, a.documentModuleEnumFindings(file)...)
	}
	if cfg.DetectWriteOnlyProperty {
		findings = append(findings, a.writeOnlyPropertyFindings(file)...)
	}
	if cfg.DetectPublicInterfaceEventMembers {
		findings = append(findings, a.publicInterfaceEventMemberFindings(file, interfaces)...)
	}
	return findings
}

// objectModuleKind reports whether the module kind participates in the
// class/interface naming contract. Implements and event-style member names are
// legal in class, form, and document modules; standard modules ignore the
// conventions entirely.
func objectModuleKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "class", "form", "document":
		return true
	}
	return false
}

// memberKind reports whether kind is a member-facing procedure kind
// (Sub/Function/Property). Event statements and Declare APIs are excluded.
func memberKind(kind procedureir.ProcedureKind) bool {
	switch kind {
	case procedureir.ProcedureSub, procedureir.ProcedureFunction,
		procedureir.ProcedurePropertyGet, procedureir.ProcedurePropertyLet,
		procedureir.ProcedurePropertySet, procedureir.ProcedureProperty:
		return true
	}
	return false
}

func isPublicMember(symbol procedureir.ProcedureSymbol) bool {
	visibility := strings.TrimSpace(symbol.Visibility)
	return visibility == "" || strings.EqualFold(visibility, "public")
}

func isExplicitPublicMember(symbol procedureir.ProcedureSymbol) bool {
	return strings.EqualFold(strings.TrimSpace(symbol.Visibility), "public")
}

// isPublicOrFriendMember reports whether a member is part of the
// project-visible API surface: implicit (empty) visibility, explicit Public,
// or Friend. Friend members are callable across the project, so a write-only
// Friend property is still a public-facing API hazard.
func isPublicOrFriendMember(symbol procedureir.ProcedureSymbol) bool {
	visibility := strings.TrimSpace(symbol.Visibility)
	return visibility == "" || strings.EqualFold(visibility, "public") || strings.EqualFold(visibility, "friend")
}

// modulePredeclaredID reports whether the file's own module name resolves to a
// predeclared default instance: document modules (sheet/workbook classes) and
// UserForms always carry the predeclared id, while class modules opt in with
// Attribute VB_PredeclaredId = True.
func modulePredeclaredID(file parsedFile) bool {
	switch strings.ToLower(strings.TrimSpace(file.ModuleKind)) {
	case "document", "form":
		return true
	case "class":
		for _, attribute := range file.IR.ModuleAttributes {
			if !strings.EqualFold(attribute.Name, "VB_PredeclaredId") {
				continue
			}
			value := strings.Trim(strings.TrimSpace(attribute.Value), "\"")
			return strings.EqualFold(value, "true") || value == "1" || value == "-1"
		}
	}
	return false
}

// interfaceBindingMatch reports whether name is a verified <Interface>_<Member>
// binding: the suffix after the interface prefix must name a public member of
// a resolved, unambiguous interface class. When a prefix-matched interface
// cannot be resolved to exactly one class module the binding stays
// unresolved and the caller fails open.
func interfaceBindingMatch(name string, interfaces map[string]struct{}, index interfaceMemberIndex) (verified, unresolved bool) {
	lower := strings.ToLower(cleanIdentifier(name))
	for iface := range interfaces {
		if !strings.HasPrefix(lower, iface+"_") {
			continue
		}
		member := strings.TrimPrefix(lower, iface+"_")
		members, resolved := index.members[iface]
		if resolved && index.moduleCounts[iface] != 1 {
			resolved = false
		}
		if resolved {
			if _, found := members[member]; found {
				verified = true
			}
			continue
		}
		unresolved = true
	}
	return verified, unresolved
}

// userFormIntrinsicEvents lists the intrinsic UserForm events that need no
// control prefix. Everything else on a form with the <name>_<event> shape must
// name a designer control before it counts as an event handler.
var userFormIntrinsicEvents = map[string]struct{}{
	"activate": {}, "addcontrol": {}, "beforedragover": {}, "beforedroporpaste": {},
	"click": {}, "dblclick": {}, "deactivate": {}, "error": {}, "initialize": {},
	"keydown": {}, "keypress": {}, "keyup": {}, "layout": {}, "mousedown": {},
	"mousemove": {}, "mouseup": {}, "queryclose": {}, "removecontrol": {},
	"resize": {}, "scroll": {}, "terminate": {}, "zoom": {},
}

// recognizedEventMember reports whether an underscored member name is a real
// event binding rather than an ordinary helper. Document-module names reuse
// the IR's IsEventHandler classification because the host event names are
// fixed; form events must additionally resolve their control prefix against
// the designer control set when that artifact is available, and fall back to
// the intrinsic UserForm_* names when controls are unknown. Class modules
// keep the IR classification (Class_Initialize/Class_Terminate are handled by
// the dedicated lifecycle name check).
func recognizedEventMember(file parsedFile, symbol procedureir.ProcedureSymbol) bool {
	if !symbol.IsEventHandler {
		return false
	}
	if !strings.EqualFold(strings.TrimSpace(file.ModuleKind), "form") {
		return true
	}
	name := strings.ToLower(cleanIdentifier(symbol.Name))
	index := strings.LastIndex(name, "_")
	if index <= 0 || index == len(name)-1 {
		return false
	}
	prefix, suffix := name[:index], name[index+1:]
	if prefix == "userform" {
		_, intrinsic := userFormIntrinsicEvents[suffix]
		return intrinsic
	}
	if file.UserFormControlsKnown {
		_, ok := file.UserFormControlNames[prefix]
		return ok
	}
	return false
}

// publicMemberUnderscoreFindings implements VBA278: public Sub/Function/
// Property members in object modules whose names contain "_" collide with
// VBA's <Interface>_<Member> and <Object>_<Event> conventions. Recognized
// event handlers and verified implemented-interface members (VBA281) are
// skipped so each member produces one focused finding; unresolved interface
// information also fails open because the binding cannot be disproved.
func (a Analyzer) publicMemberUnderscoreFindings(file parsedFile, interfaces interfaceMemberIndex) []Finding {
	if !objectModuleKind(file.ModuleKind) {
		return nil
	}
	declared := implementedInterfaceNames(file.IR)
	isClassModule := strings.EqualFold(strings.TrimSpace(file.ModuleKind), "class")
	var findings []Finding
	for i := range file.IR.Procedures {
		symbol := file.IR.Procedures[i].Symbol
		if !memberKind(symbol.Kind) || symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
			continue
		}
		if !isPublicMember(symbol) || recognizedEventMember(file, symbol) {
			continue
		}
		name := cleanIdentifier(symbol.Name)
		if !strings.Contains(name, "_") {
			continue
		}
		if isClassModule && classLifecycleMemberName(name) {
			continue
		}
		verified, unresolved := interfaceBindingMatch(name, declared, interfaces)
		if verified || unresolved {
			continue
		}
		finding := a.simpleFinding(
			file, sourceProcedure{Name: symbol.Name}, symbol.DeclarationRange.StartLine, "VBA278", "warning",
			fmt.Sprintf("Public %s '%s' in %s module '%s' contains an underscore.",
				classHazardMemberLabel(symbol.Kind), name, strings.ToLower(file.ModuleKind), file.Module),
			"Underscored public member names collide with VBA's <Interface>_<Member> implementation convention and <Object>_<Event> handler names; a matching Implements member or control event becomes ambiguous.",
			"Rename the member without underscores, or keep it Private if it is an internal helper.",
		)
		applyDeclarationRange(&finding, symbol.DeclarationRange)
		finding.ScopeEndLine = symbol.BodyRange.EndLine
		findings = append(findings, finding)
	}
	return findings
}

// classLifecycleMemberName reports whether name is a fixed VBA class lifecycle
// handler (Class_Initialize or Class_Terminate). VBA mandates those exact
// names, so they cannot be renamed to satisfy the underscore convention and
// are not flagged.
func classLifecycleMemberName(name string) bool {
	lower := strings.ToLower(name)
	return lower == "class_initialize" || lower == "class_terminate"
}

// documentModuleEnumFindings implements VBA279: a Public Enum inside a
// worksheet/document module is copied into every worksheet copy produced at
// runtime, duplicating the public name and producing an ambiguous-name
// compile error. Standard and class modules are unaffected.
func (a Analyzer) documentModuleEnumFindings(file parsedFile) []Finding {
	if !strings.EqualFold(strings.TrimSpace(file.ModuleKind), "document") {
		return nil
	}
	var findings []Finding
	for _, declaration := range file.IR.Declarations {
		if declaration.Kind != "enum" || declaration.Recovered {
			continue
		}
		if visibility := strings.TrimSpace(declaration.Visibility); strings.EqualFold(visibility, "private") {
			continue
		}
		name := cleanIdentifier(declaration.Name)
		finding := a.simpleFinding(
			file, sourceProcedure{}, declaration.Range.StartLine, "VBA279", "warning",
			fmt.Sprintf("Public Enum '%s' is declared in document module '%s'.", name, file.Module),
			"Public Enum declarations in worksheet or workbook code-behind are duplicated when the host copies the sheet, producing an ambiguous-name compile error.",
			"Move the Enum to a standard module, or declare it Private if worksheet copies must keep working.",
		)
		applyDeclarationRange(&finding, declaration.Range)
		findings = append(findings, finding)
	}
	return findings
}

// writeOnlyPropertyFindings implements VBA280: a Property Let/Set writer with
// no matching Property Get on the same name is a write-only API. Private
// writers are module-internal conveniences and stay silent; Public and Friend
// writers (and implicit-public writers) are flagged. An accessor marked
// recovered or declared only under a conditional-compilation branch makes the
// whole property name uncertain, so the finding fails open: a conditional
// getter may be absent from the analyzed IR for other build configurations.
func (a Analyzer) writeOnlyPropertyFindings(file parsedFile) []Finding {
	type propertyFacts struct {
		hasGet    bool
		uncertain bool
		writers   []procedureir.ProcedureIR
	}
	byName := map[string]*propertyFacts{}
	var order []string
	for i := range file.IR.Procedures {
		procedure := file.IR.Procedures[i]
		symbol := procedure.Symbol
		isWriter := false
		switch symbol.Kind {
		case procedureir.ProcedurePropertyGet:
		case procedureir.ProcedurePropertyLet, procedureir.ProcedurePropertySet:
			isWriter = true
		default:
			continue
		}
		key := strings.ToLower(cleanIdentifier(symbol.Name))
		if key == "" {
			continue
		}
		entry := byName[key]
		if entry == nil {
			entry = &propertyFacts{}
			byName[key] = entry
			order = append(order, key)
		}
		if symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
			entry.uncertain = true
			continue
		}
		if symbol.Kind == procedureir.ProcedurePropertyGet {
			entry.hasGet = true
		}
		if isWriter {
			entry.writers = append(entry.writers, procedure)
		}
	}
	var findings []Finding
	for _, key := range order {
		entry := byName[key]
		if entry.hasGet || entry.uncertain {
			continue
		}
		var publicWriters []procedureir.ProcedureIR
		var kinds []string
		seenKind := map[string]bool{}
		for _, writer := range entry.writers {
			if !isPublicOrFriendMember(writer.Symbol) {
				continue
			}
			publicWriters = append(publicWriters, writer)
			label := classHazardMemberLabel(writer.Symbol.Kind)
			if !seenKind[label] {
				seenKind[label] = true
				kinds = append(kinds, label)
			}
		}
		if len(publicWriters) == 0 {
			continue
		}
		first := publicWriters[0].Symbol
		name := cleanIdentifier(first.Name)
		finding := a.simpleFinding(
			file, sourceProcedure{Name: first.Name}, first.DeclarationRange.StartLine, "VBA280", "warning",
			fmt.Sprintf("Property '%s' is write-only: it declares %s but no Property Get.", name, strings.Join(kinds, " and ")),
			"A public property that can be assigned but never read usually signals an API-design bug; callers cannot observe the value they stored.",
			"Add a Property Get with the matching type, or convert the writer to a named Sub if the value is consumed internally only.",
		)
		applyDeclarationRange(&finding, first.DeclarationRange)
		finding.ScopeEndLine = first.BodyRange.EndLine
		findings = append(findings, finding)
	}
	return findings
}

// publicInterfaceEventMemberFindings implements VBA281: a member is flagged
// when (a) it is implicitly or explicitly public and its name is a verified
// <Interface>_<Member> binding - the corresponding private implementation
// already exists in VBA's member namespace, so a public member with the same
// name collides or leaks the interface implementation - or (b) a recognized
// event handler is declared explicitly Public, which exposes the handler on
// the default interface where callers can invoke it directly. Unresolved
// interface information fails open: a prefix-shaped name whose interface or
// member cannot be verified is left to VBA278 instead of claimed as a
// collision.
func (a Analyzer) publicInterfaceEventMemberFindings(file parsedFile, interfaces interfaceMemberIndex) []Finding {
	if !objectModuleKind(file.ModuleKind) {
		return nil
	}
	declared := implementedInterfaceNames(file.IR)
	var findings []Finding
	for i := range file.IR.Procedures {
		symbol := file.IR.Procedures[i].Symbol
		if !memberKind(symbol.Kind) || symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
			continue
		}
		if !isPublicMember(symbol) {
			continue
		}
		name := cleanIdentifier(symbol.Name)
		verified, _ := interfaceBindingMatch(name, declared, interfaces)
		interfaceCollision := verified
		publicEvent := recognizedEventMember(file, symbol) && isExplicitPublicMember(symbol)
		if !interfaceCollision && !publicEvent {
			continue
		}
		var message, reason string
		if interfaceCollision {
			message = fmt.Sprintf("Public %s '%s' matches the implemented interface binding %s and competes with the private implementation member.", classHazardMemberLabel(symbol.Kind), name, implementedInterfaceLabel(name, declared))
			reason = "Interface implementations in VBA live in the <Interface>_<Member> namespace; a public member with the same name produces an ambiguous-name conflict or unintentionally exposes the implementation."
		} else {
			message = fmt.Sprintf("Event handler '%s' is declared Public on the module's default interface.", name)
			reason = "Event procedures are raised by the host; exposing them as Public members lets any caller invoke the handler and obscures the event wiring."
		}
		finding := a.simpleFinding(
			file, sourceProcedure{Name: symbol.Name}, symbol.DeclarationRange.StartLine, "VBA281", "warning",
			message, reason,
			"Declare the member Private so it cannot be called through the default interface.",
		)
		applyDeclarationRange(&finding, symbol.DeclarationRange)
		finding.ScopeEndLine = symbol.BodyRange.EndLine
		findings = append(findings, finding)
	}
	return findings
}

func implementedInterfaceLabel(name string, interfaces map[string]struct{}) string {
	lower := strings.ToLower(cleanIdentifier(name))
	ordered := make([]string, 0, len(interfaces))
	for iface := range interfaces {
		ordered = append(ordered, iface)
	}
	slices.Sort(ordered)
	for _, iface := range ordered {
		if strings.HasPrefix(lower, iface+"_") {
			return "'" + iface + "_*'"
		}
	}
	return "'<Interface>_<Member>'"
}

// predeclaredInstanceFindings implements VBA282 for a single procedure: a
// reference to the containing module's own name inside a predeclared module
// (document, form, or Attribute VB_PredeclaredId = True class) resolves to
// the module's predeclared default instance, not to Me. Accesses in type
// positions (As clauses, Implements targets, new/typeof expression type
// operands) are already classified as metadata and never produce an access.
func (a Analyzer) predeclaredInstanceFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectPredeclaredInstanceAccess {
		return nil
	}
	module := strings.TrimSpace(file.Module)
	if module == "" || !modulePredeclaredID(file) || proc.IR == nil {
		return nil
	}
	if proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		return nil
	}
	var findings []Finding
	seen := map[vbaast.Range]bool{}
	report := func(access procedureir.VariableAccess) {
		if !strings.EqualFold(cleanIdentifier(access.Name), module) {
			return
		}
		if access.Range.StartByte < 0 || seen[access.Range] {
			return
		}
		if predeclaredAccessIsTypeOperand(proc.IR, access) {
			return
		}
		if access.Scope == procedureir.ScopeLocal || access.Scope == procedureir.ScopeParameter ||
			access.Scope == procedureir.ScopeModule {
			return
		}
		seen[access.Range] = true
		finding := a.simpleFinding(
			file, proc, access.Range.StartLine, "VBA282", "warning",
			fmt.Sprintf("'%s' inside %s '%s' refers to the predeclared default instance, not to the current instance (Me).", module, strings.ToLower(file.ModuleKind), module),
			"Unqualified self-name references on a predeclared class bind to the shared default instance, so state read or written through them silently diverges from Me.",
			"Use Me (or an unqualified member name) for the current instance; qualify with the class name only when the shared default instance is intended.",
		)
		finding.Column = access.Range.StartColumn
		finding.EndLine = access.Range.EndLine
		finding.EndColumn = access.Range.EndColumn
		finding.ScopeEndLine = proc.EndLine
		findings = append(findings, finding)
	}
	for access := range proc.Accesses.All() {
		report(access)
	}
	slices.SortFunc(findings, func(x, y Finding) int {
		return cmp.Or(cmp.Compare(x.Line, y.Line), cmp.Compare(x.Column, y.Column))
	})
	return findings
}

// predeclaredAccessIsTypeOperand reports whether the access denotes the type
// operand of a TypeOf ... Is <Name> expression or the type in a New <Name>
// expression rather than the default instance. Only the right-hand type
// operand of TypeOf is excluded: the left-hand object expression still reads
// the default instance. Declaration positions are excluded upstream because
// they never produce an Access.
func predeclaredAccessIsTypeOperand(ir *procedureir.ProcedureIR, access procedureir.VariableAccess) bool {
	if ir == nil || access.ExpressionID <= 0 {
		return false
	}
	id := access.ExpressionID
	for id > 0 && id <= len(ir.Expressions) {
		expression := ir.Expressions[id-1]
		kind := strings.ToLower(strings.TrimSpace(expression.SyntaxKind))
		if kind == "new_expression" || kind == "type_expression" {
			return true
		}
		parentID := expression.ParentID
		if parentID <= 0 || parentID > len(ir.Expressions) {
			return false
		}
		parent := ir.Expressions[parentID-1]
		if strings.Contains(strings.ToLower(strings.TrimSpace(parent.SyntaxKind)), "type_of") {
			// The type operand of "TypeOf <expr> Is <type>" is the last child
			// expression; accesses on the left-hand side still bind to the
			// default instance and must not be excluded.
			if len(parent.Children) > 0 && parent.Children[len(parent.Children)-1] == id {
				return true
			}
		}
		id = parentID
	}
	return false
}

// classHazardMemberLabel renders the procedure kind used in finding messages.
func classHazardMemberLabel(kind procedureir.ProcedureKind) string {
	switch kind {
	case procedureir.ProcedureSub:
		return "Sub"
	case procedureir.ProcedureFunction:
		return "Function"
	case procedureir.ProcedurePropertyGet:
		return "Property Get"
	case procedureir.ProcedurePropertyLet:
		return "Property Let"
	case procedureir.ProcedurePropertySet:
		return "Property Set"
	case procedureir.ProcedureProperty:
		return "Property"
	}
	return "member"
}

func applyDeclarationRange(finding *Finding, declared vbaast.Range) {
	finding.Column = declared.StartColumn
	finding.EndLine = declared.EndLine
	finding.EndColumn = declared.EndColumn
}
