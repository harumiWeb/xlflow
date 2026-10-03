// Package modulestate collects protocol-neutral facts about module-scope VBA
// state from already-resolved procedure IR.
package modulestate

import (
	"cmp"
	"context"
	"maps"
	"path/filepath"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// Options supplies path context needed when resolved candidates use
// project-relative paths while the caller's source files use absolute paths.
type Options struct {
	RootDir string
}

// Facts contains indexed module fields and their resolved procedure uses.
// IDs are stable within the supplied document set and are independent of any
// diagnostic or output protocol.
type Facts struct {
	Fields     []Field
	Procedures []Procedure
}

// Field describes one module-scope variable or constant.
type Field struct {
	ID              string
	Path            string
	Module          string
	ModuleKind      string
	Name            string
	Type            string
	Visibility      string
	Kind            string
	Scope           procedureir.SymbolScope
	DeclarationLine int
	IsObject        bool
	IsCollection    bool
	IsExcel         bool
	Readers         []string
	Writers         []string
	Mutators        []string
}

// Procedure describes a procedure and its indexed module-state relationships.
// Callees contains only uniquely resolved project-local procedure IDs.
type Procedure struct {
	ID              string
	Path            string
	Module          string
	ModuleKind      string
	Name            string
	QualifiedName   string
	Kind            procedureir.ProcedureKind
	DeclarationLine int
	ParameterCount  int
	Visibility      string
	EventHandler    bool
	Reads           []string
	Writes          []string
	Mutators        []string
	Callees         []string
}

type fieldBuilder struct {
	Field
	resolvedPath string
	readers      map[string]struct{}
	writers      map[string]struct{}
	mutators     map[string]struct{}
}

type procedureBuilder struct {
	Procedure
	resolvedPath string
	reads        map[string]struct{}
	writes       map[string]struct{}
	mutators     map[string]struct{}
	callees      map[string]struct{}
}

// Collect builds module-state facts from resolved documents without emitting
// diagnostics. Documents and their nested IR slices are read-only inputs.
func Collect(documents []procedureir.DocumentIR) Facts {
	return CollectWithOptions(documents, Options{})
}

// CollectWithOptions builds module-state facts with optional project path
// context. It is useful when IR paths and resolved candidate paths have
// different absolute/relative forms.
func CollectWithOptions(documents []procedureir.DocumentIR, options Options) Facts {
	facts, _ := CollectContext(context.Background(), documents, options)
	return facts
}

// CollectContext is the cancellable form of CollectWithOptions. It checks the
// context between documents and procedures and returns no partial facts when
// cancellation is observed.
func CollectContext(ctx context.Context, documents []procedureir.DocumentIR, options Options) (Facts, error) {
	if err := ctx.Err(); err != nil {
		return Facts{}, err
	}

	procedures := make([]*procedureBuilder, 0)
	procedureByID := make(map[string]*procedureBuilder)
	procedureByCandidate := make(map[string]string)
	fields := make([]*fieldBuilder, 0)
	fieldsByID := make(map[string]*fieldBuilder)
	fieldsByName := make(map[string][]*fieldBuilder)

	for docIndex := range documents {
		if err := ctx.Err(); err != nil {
			return Facts{}, err
		}
		document := &documents[docIndex]
		for procedureIndex := range document.Procedures {
			if err := ctx.Err(); err != nil {
				return Facts{}, err
			}
			ir := &document.Procedures[procedureIndex]
			id := procedureID(document.Path, document.ModuleName, ir.Symbol)
			procedure := &procedureBuilder{
				Procedure: Procedure{
					ID: id, Path: document.Path, Module: document.ModuleName,
					ModuleKind: document.ModuleKind, Name: ir.Symbol.Name,
					QualifiedName: ir.Symbol.QualifiedName, Kind: ir.Symbol.Kind,
					DeclarationLine: ir.Symbol.DeclarationRange.StartLine,
					ParameterCount:  len(ir.Symbol.Parameters), Visibility: ir.Symbol.Visibility,
					EventHandler: ir.Symbol.IsEventHandler,
				},
				resolvedPath: resolvePath(options.RootDir, document.Path),
				reads:        make(map[string]struct{}),
				writes:       make(map[string]struct{}),
				mutators:     make(map[string]struct{}),
				callees:      make(map[string]struct{}),
			}
			procedures = append(procedures, procedure)
			procedureByID[id] = procedure
			procedureByCandidate[candidateIdentity(document.Path, ir.Symbol.QualifiedName, ir.Symbol.Kind, ir.Symbol.DeclarationRange.StartLine)] = id
		}

		for declarationIndex := range document.Declarations {
			declaration := &document.Declarations[declarationIndex]
			if declaration.Scope != procedureir.ScopeModule && declaration.Scope != procedureir.ScopeProject {
				continue
			}
			if declaration.Kind != "variable" && declaration.Kind != "const" {
				continue
			}
			id := fieldID(document.Path, document.ModuleName, declaration.Name)
			field := &fieldBuilder{
				Field: Field{
					ID: id, Path: document.Path, Module: document.ModuleName,
					ModuleKind: document.ModuleKind, Name: declaration.Name,
					Type: declaration.Type, Visibility: declaration.Visibility,
					Kind: declaration.Kind, Scope: declaration.Scope,
					DeclarationLine: declaration.Range.StartLine,
					IsObject:        declaration.IsObject,
					IsCollection:    collectionType(declaration.Type),
					IsExcel:         declaration.IsObject && excelType(declaration.Type),
				},
				resolvedPath: resolvePath(options.RootDir, document.Path),
				readers:      make(map[string]struct{}),
				writers:      make(map[string]struct{}),
				mutators:     make(map[string]struct{}),
			}
			fields = append(fields, field)
			fieldsByID[id] = field
			nameKey := strings.ToLower(strings.TrimSpace(declaration.Name))
			fieldsByName[nameKey] = append(fieldsByName[nameKey], field)
		}
	}

	// Detect late-bound Collection/Dictionary initialization before attributing
	// mutators so source order does not change whether a field is classified.
	for docIndex := range documents {
		if err := ctx.Err(); err != nil {
			return Facts{}, err
		}
		document := &documents[docIndex]
		for procedureIndex := range document.Procedures {
			if err := ctx.Err(); err != nil {
				return Facts{}, err
			}
			ir := &document.Procedures[procedureIndex]
			procedure := procedureByID[procedureID(document.Path, document.ModuleName, ir.Symbol)]
			for accessIndex := range ir.Accesses {
				access := &ir.Accesses[accessIndex]
				field := resolveField(options.RootDir, *access, procedure, fieldsByID, fieldsByName)
				if field != nil && !field.IsCollection && collectionInitializer(*ir, *access, field.Name) {
					field.IsCollection = true
				}
			}
		}
	}

	for docIndex := range documents {
		if err := ctx.Err(); err != nil {
			return Facts{}, err
		}
		document := &documents[docIndex]
		for procedureIndex := range document.Procedures {
			if err := ctx.Err(); err != nil {
				return Facts{}, err
			}
			ir := &document.Procedures[procedureIndex]
			procedure := procedureByID[procedureID(document.Path, document.ModuleName, ir.Symbol)]
			for accessIndex := range ir.Accesses {
				access := &ir.Accesses[accessIndex]
				field := resolveField(options.RootDir, *access, procedure, fieldsByID, fieldsByName)
				if field == nil {
					continue
				}
				if access.Mode == procedureir.AccessRead || access.Mode == procedureir.AccessReadWrite {
					field.readers[procedure.ID] = struct{}{}
					procedure.reads[field.ID] = struct{}{}
				}
				if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
					field.writers[procedure.ID] = struct{}{}
					procedure.writes[field.ID] = struct{}{}
				}
			}

			for callIndex := range ir.Calls {
				call := &ir.Calls[callIndex]
				if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
					candidate := call.Resolution.Candidates[0]
					callee := procedureByCandidate[candidateIdentity(candidate.File, candidate.QualifiedName, procedureir.ProcedureKind(candidate.Kind), candidate.Line)]
					if callee != "" {
						procedure.callees[callee] = struct{}{}
					}
				}

				member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
				if !collectionMutator(*ir, *call, member) {
					continue
				}
				name := receiverName(*call)
				if name == "" || procedureDeclaresName(*ir, name) {
					continue
				}
				field := resolveName(name, procedure, fieldsByID, fieldsByName)
				if field == nil || !field.IsCollection {
					continue
				}
				field.mutators[procedure.ID] = struct{}{}
				field.writers[procedure.ID] = struct{}{}
				procedure.mutators[field.ID] = struct{}{}
				procedure.writes[field.ID] = struct{}{}
			}
		}
	}

	result := Facts{
		Fields:     make([]Field, 0, len(fields)),
		Procedures: make([]Procedure, 0, len(procedures)),
	}
	for _, field := range fields {
		field.Readers = slices.Sorted(maps.Keys(field.readers))
		field.Writers = slices.Sorted(maps.Keys(field.writers))
		field.Mutators = slices.Sorted(maps.Keys(field.mutators))
		result.Fields = append(result.Fields, field.Field)
	}
	for _, procedure := range procedures {
		procedure.Reads = slices.Sorted(maps.Keys(procedure.reads))
		procedure.Writes = slices.Sorted(maps.Keys(procedure.writes))
		procedure.Mutators = slices.Sorted(maps.Keys(procedure.mutators))
		procedure.Callees = slices.Sorted(maps.Keys(procedure.callees))
		result.Procedures = append(result.Procedures, procedure.Procedure)
	}
	slices.SortFunc(result.Fields, func(left, right Field) int {
		if order := cmp.Compare(strings.ToLower(left.Path), strings.ToLower(right.Path)); order != 0 {
			return order
		}
		if order := cmp.Compare(left.DeclarationLine, right.DeclarationLine); order != 0 {
			return order
		}
		if order := cmp.Compare(strings.ToLower(left.Name), strings.ToLower(right.Name)); order != 0 {
			return order
		}
		return cmp.Compare(left.ID, right.ID)
	})
	slices.SortFunc(result.Procedures, func(left, right Procedure) int {
		return cmp.Compare(left.ID, right.ID)
	})
	return result, nil
}

func resolveField(rootDir string, access procedureir.VariableAccess, procedure *procedureBuilder, byID map[string]*fieldBuilder, byName map[string][]*fieldBuilder) *fieldBuilder {
	if access.Scope != procedureir.ScopeModule && access.Scope != procedureir.ScopeProject {
		return nil
	}
	if access.Resolution.Scope == procedureir.ScopeProject && len(access.Resolution.Candidates) == 1 {
		candidate := access.Resolution.Candidates[0]
		candidatePath := candidate.File
		name := strings.ToLower(strings.TrimSpace(access.Name))
		for _, field := range byName[name] {
			if samePath(field.resolvedPath, candidatePath) && field.DeclarationLine == candidate.Line {
				return field
			}
		}
		if candidatePath != "" && !filepath.IsAbs(candidatePath) && rootDir != "" {
			candidatePath = resolvePath(rootDir, candidatePath)
			for _, field := range byName[name] {
				if samePath(field.resolvedPath, candidatePath) && field.DeclarationLine == candidate.Line {
					return field
				}
			}
		}
	}
	return resolveName(access.Name, procedure, byID, byName)
}

func resolveName(name string, procedure *procedureBuilder, byID map[string]*fieldBuilder, byName map[string][]*fieldBuilder) *fieldBuilder {
	name = strings.ToLower(strings.TrimSpace(name))
	if name == "" {
		return nil
	}
	if field := byID[fieldID(procedure.Path, procedure.Module, name)]; field != nil {
		return field
	}
	for _, field := range byName[name] {
		if samePath(field.resolvedPath, procedure.resolvedPath) && strings.EqualFold(field.Module, procedure.Module) {
			return field
		}
	}
	for _, field := range byName[name] {
		if field.Scope == procedureir.ScopeProject && strings.EqualFold(field.ModuleKind, "standard") {
			return field
		}
	}
	return nil
}

func collectionMutator(procedure procedureir.ProcedureIR, call procedureir.CallSite, member string) bool {
	switch member {
	case "add", "remove", "removeall", "clear", "delete":
		return true
	case "item", "comparemode":
		for statementIndex := range procedure.Statements {
			statement := &procedure.Statements[statementIndex]
			if statement.ID != call.StatementID {
				continue
			}
			if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
				return false
			}
			return expressionContains(procedure.Expressions, statement.TargetID, call.ExpressionID)
		}
	}
	return false
}

func expressionContains(expressions []procedureir.Expression, rootID, childID int) bool {
	if rootID <= 0 || childID <= 0 {
		return false
	}
	for current := childID; current > 0 && current <= len(expressions); {
		if current == rootID {
			return true
		}
		current = expressions[current-1].ParentID
	}
	return false
}

func receiverName(call procedureir.CallSite) string {
	if call.Callee.Receiver != nil {
		return cleanIdentifier(*call.Callee.Receiver)
	}
	text := strings.TrimSpace(call.Callee.Text)
	if dot := strings.LastIndex(text, "."); dot >= 0 {
		return cleanIdentifier(text[:dot])
	}
	return ""
}

func cleanIdentifier(text string) string {
	text = strings.TrimSpace(strings.Trim(text, "[]"))
	return strings.TrimRight(text, "$%&#@^!")
}

func collectionInitializer(procedure procedureir.ProcedureIR, access procedureir.VariableAccess, fieldName string) bool {
	if !strings.EqualFold(strings.TrimSpace(access.Name), strings.TrimSpace(fieldName)) {
		return false
	}
	for statementIndex := range procedure.Statements {
		statement := &procedure.Statements[statementIndex]
		if access.StatementID > 0 && statement.ID != access.StatementID {
			continue
		}
		text := strings.ToLower(statement.Text)
		if fieldName != "" && !strings.Contains(text, strings.ToLower(strings.TrimSpace(fieldName))) {
			continue
		}
		if strings.Contains(text, "new collection") || strings.Contains(text, "scripting.dictionary") {
			return true
		}
	}
	return false
}

func procedureDeclaresName(procedure procedureir.ProcedureIR, name string) bool {
	name = strings.TrimSpace(name)
	if name == "" {
		return false
	}
	for _, parameter := range procedure.Symbol.Parameters {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), name) {
			return true
		}
	}
	for declarationIndex := range procedure.Declarations {
		declaration := &procedure.Declarations[declarationIndex]
		if !strings.EqualFold(strings.TrimSpace(declaration.Name), name) {
			continue
		}
		if declaration.Scope == procedureir.ScopeLocal || declaration.Scope == procedureir.ScopeParameter {
			return true
		}
	}
	return false
}

func collectionType(typ string) bool {
	switch unqualifiedType(typ) {
	case "collection", "dictionary":
		return true
	default:
		return false
	}
}

// IsCollectionType reports whether typ names a built-in VBA Collection or
// Scripting.Dictionary, optionally with a recognized library qualifier.
func IsCollectionType(typ string) bool {
	return collectionType(typ)
}

func excelType(typ string) bool {
	switch unqualifiedType(typ) {
	case "application", "workbook", "worksheet", "range", "chart", "pivottable", "listobject", "window":
		return true
	default:
		return false
	}
}

// IsExcelType reports whether typ names one of the Excel object types tracked
// as cached module state by the analyzer.
func IsExcelType(typ string) bool {
	return excelType(typ)
}

func unqualifiedType(typ string) string {
	lower := strings.ToLower(strings.TrimSpace(typ))
	for _, qualifier := range []string{"excel.", "vba.", "scripting."} {
		if unqualified, ok := strings.CutPrefix(lower, qualifier); ok {
			return strings.TrimSpace(unqualified)
		}
	}
	return lower
}

func procedureID(path, module string, symbol procedureir.ProcedureSymbol) string {
	return candidateIdentity(path, module+"."+symbol.Name, symbol.Kind, symbol.DeclarationRange.StartLine)
}

func candidateIdentity(path, qualified string, kind procedureir.ProcedureKind, line int) string {
	return strings.Join([]string{
		strings.ToLower(filepath.ToSlash(filepath.Clean(path))),
		strings.ToLower(strings.TrimSpace(qualified)),
		strings.ToLower(string(kind)), strconv.Itoa(line),
	}, "\x00")
}

func fieldID(path, module, name string) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(path)) + "\x00" + strings.TrimSpace(module) + "\x00" + strings.TrimSpace(name))
}

func resolvePath(rootDir, path string) string {
	if !filepath.IsAbs(path) && rootDir != "" {
		path = filepath.Join(rootDir, filepath.FromSlash(path))
	}
	return filepath.Clean(path)
}

func samePath(left, right string) bool {
	return strings.EqualFold(pathKey(left), pathKey(right))
}

func pathKey(path string) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(path)))
}
