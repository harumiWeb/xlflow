package analyze

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

const publicAPITypeRule = "VBA222"

type apiTypeInfo struct {
	Name       string
	Module     string
	Kind       string
	Visibility string
	File       string
	Accessible bool
}

type apiTypeIndex struct {
	byName             map[string][]apiTypeInfo
	db                 *vbadb.DB
	resolutionComplete bool
}

func buildAPITypeIndex(files []parsedFile, db *vbadb.DB, resolutionComplete bool) *apiTypeIndex {
	index := &apiTypeIndex{
		byName:             map[string][]apiTypeInfo{},
		db:                 db,
		resolutionComplete: resolutionComplete,
	}
	for _, file := range files {
		module := strings.TrimSpace(file.IR.ModuleName)
		if module == "" {
			module = file.Module
		}
		if strings.EqualFold(file.IR.ModuleKind, "class") {
			index.add(apiTypeInfo{
				Name: module, Module: module, Kind: "class", File: file.Path,
				Accessible: moduleIsExposed(file.IR),
			})
		}
		for _, declaration := range file.IR.Declarations {
			if declaration.Kind != "type" && declaration.Kind != "enum" {
				continue
			}
			accessible := !strings.EqualFold(declaration.Visibility, "Private")
			if !strings.EqualFold(file.IR.ModuleKind, "standard") && !moduleIsExposed(file.IR) {
				accessible = false
			}
			index.add(apiTypeInfo{
				Name: declaration.Name, Module: module, Kind: declaration.Kind,
				Visibility: declaration.Visibility, File: file.Path, Accessible: accessible,
			})
		}
	}
	return index
}

func (index *apiTypeIndex) add(info apiTypeInfo) {
	name := publicAPITypeName(info.Name)
	if name == "" {
		return
	}
	key := strings.ToLower(name)
	index.byName[key] = append(index.byName[key], info)
}

func moduleIsExposed(document procedureir.DocumentIR) bool {
	for _, attribute := range document.ModuleAttributes {
		if !strings.EqualFold(attribute.Name, "VB_Exposed") {
			continue
		}
		value := strings.Trim(strings.TrimSpace(attribute.Value), "\"")
		return strings.EqualFold(value, "true") || value == "1" || value == "-1"
	}
	return false
}

func (a Analyzer) publicAPITypeFindings(file parsedFile, index *apiTypeIndex) []Finding {
	if index == nil || !a.Config.Analyze.DetectPublicAPITypeSafety {
		return nil
	}
	if !publicAPIModule(file.IR) {
		return nil
	}
	procedures := sourceProceduresFromIR(file.IR)
	var findings []Finding
	for i, procedure := range file.IR.Procedures {
		if i >= len(procedures) || !publicAPIProcedure(procedure.Symbol) || procedure.Symbol.IsEventHandler {
			continue
		}
		proc := procedures[i]
		if procedure.Symbol.Kind == procedureir.ProcedureFunction ||
			procedure.Symbol.Kind == procedureir.ProcedurePropertyGet {
			findings = append(findings, a.checkPublicAPIType(file, proc, index, procedure.Symbol.ReturnType, proc.StartLine, "return type")...)
		}
		for _, parameter := range proc.Params {
			line := proc.StartLine
			if parameter.Range.StartLine > 0 {
				line = parameter.Range.StartLine
			}
			findings = append(findings, a.checkPublicAPIType(file, proc, index, parameter.Type, line, "parameter "+parameter.Name)...)
		}
	}
	findings = append(findings, a.propertyVisibilityFindings(file, index, procedures)...)
	for _, declaration := range file.IR.Declarations {
		if !isEventDeclaration(declaration) || !publicEventDeclaration(file.IR, declaration) {
			continue
		}
		proc := sourceProcedure{Name: declaration.Name, ModuleKind: file.IR.ModuleKind, StartLine: declaration.Range.StartLine, EndLine: declaration.Range.EndLine}
		for _, parameter := range declaration.Parameters {
			line := declaration.Range.StartLine
			if parameter.Range.StartLine > 0 {
				line = parameter.Range.StartLine
			}
			findings = append(findings, a.checkPublicAPIType(file, proc, index, parameter.Type, line, "event parameter "+parameter.Name)...)
		}
	}
	return findings
}

func publicAPIModule(document procedureir.DocumentIR) bool {
	if strings.EqualFold(document.ModuleKind, "standard") {
		return true
	}
	return strings.EqualFold(document.ModuleKind, "class") && moduleIsExposed(document)
}

func publicAPIProcedure(symbol procedureir.ProcedureSymbol) bool {
	return !strings.EqualFold(symbol.Visibility, "Private") && !strings.EqualFold(symbol.Visibility, "Friend")
}

func publicEventDeclaration(document procedureir.DocumentIR, declaration procedureir.Declaration) bool {
	if strings.EqualFold(declaration.Visibility, "Private") || strings.EqualFold(declaration.Visibility, "Friend") {
		return false
	}
	return publicAPIModule(document)
}

func isEventDeclaration(declaration procedureir.Declaration) bool {
	return strings.EqualFold(declaration.Kind, "event") || strings.EqualFold(declaration.Kind, "event_declaration")
}

func (a Analyzer) checkPublicAPIType(file parsedFile, proc sourceProcedure, index *apiTypeIndex, typeText string, line int, role string) []Finding {
	typeName := publicAPITypeName(typeText)
	if typeName == "" || isSafePublicAPIType(typeName) {
		return nil
	}
	status := index.resolve(typeName, file.IR.ModuleName)
	if status == apiTypeAllowed {
		return nil
	}
	// Absence from the TypeLib database is not evidence of an unresolved
	// external type while the project or generated TypeLib view is incomplete.
	// Project-local visibility and ambiguity findings remain actionable because
	// they do not depend on proving the complete external reference set.
	if status == apiTypeUnresolved && !index.resolutionComplete {
		return nil
	}
	var message, reason, suggestion string
	switch status {
	case apiTypeInaccessible:
		message = fmt.Sprintf("Public API %s exposes inaccessible type %s.", role, typeName)
		reason = fmt.Sprintf("The %s uses %s, but that project type is private or belongs to a non-exposed module and cannot safely be consumed through the public API.", role, typeName)
		suggestion = fmt.Sprintf("Make %s public and externally exposed, or replace %s with an accessible API type.", typeName, typeName)
	case apiTypeAmbiguous:
		message = fmt.Sprintf("Public API %s exposes ambiguous type %s.", role, typeName)
		reason = fmt.Sprintf("The %s uses %s, which resolves to multiple project types or references with conflicting visibility.", role, typeName)
		suggestion = fmt.Sprintf("Qualify %s with its module or remove the duplicate type definitions before exposing it publicly.", typeName)
	default:
		message = fmt.Sprintf("Public API %s references unresolved external type %s.", role, typeName)
		reason = fmt.Sprintf("The %s uses %s, but the type is not present in the project index or the available built-in/TypeLib database.", role, typeName)
		suggestion = fmt.Sprintf("Add the required reference or TypeLib metadata for %s, or change the public signature to a known accessible type.", typeName)
	}
	finding := a.simpleFinding(file, proc, line, publicAPITypeRule, "warning", message, reason, suggestion)
	return []Finding{finding}
}

type apiTypeStatus int

const (
	apiTypeAllowed apiTypeStatus = iota
	apiTypeInaccessible
	apiTypeAmbiguous
	apiTypeUnresolved
)

func (index *apiTypeIndex) resolve(typeName, ownerModule string) apiTypeStatus {
	shortName := typeName
	qualifiedModule := ""
	if dot := strings.LastIndex(typeName, "."); dot >= 0 {
		qualifiedModule = strings.TrimSpace(typeName[:dot])
		shortName = strings.TrimSpace(typeName[dot+1:])
	}
	candidates := append([]apiTypeInfo(nil), index.byName[strings.ToLower(shortName)]...)
	if qualifiedModule != "" {
		filtered := candidates[:0]
		for _, candidate := range candidates {
			if strings.EqualFold(candidate.Module, qualifiedModule) || strings.EqualFold(candidate.Name, typeName) {
				filtered = append(filtered, candidate)
			}
		}
		candidates = filtered
	}
	if qualifiedModule == "" {
		local := candidates[:0]
		for _, candidate := range candidates {
			if strings.EqualFold(candidate.Module, ownerModule) {
				local = append(local, candidate)
			}
		}
		if len(local) > 0 {
			candidates = local
		}
	}
	if len(candidates) == 1 {
		if candidates[0].Accessible {
			return apiTypeAllowed
		}
		return apiTypeInaccessible
	}
	if len(candidates) > 1 {
		return apiTypeAmbiguous
	}
	if knownAPITypeInDB(index.db, typeName) {
		return apiTypeAllowed
	}
	if qualifiedModule != "" {
		if knownAPITypeInLibrary(index.db, shortName, qualifiedModule) {
			return apiTypeAllowed
		}
	}
	return apiTypeUnresolved
}

func knownAPITypeInLibrary(db *vbadb.DB, name, library string) bool {
	if db == nil {
		return false
	}
	if typ, ok := db.ResolveType(name); ok && strings.EqualFold(typ.Library, library) {
		return true
	}
	for _, constant := range db.Constants {
		if strings.EqualFold(constant.EnumGroup, name) && strings.EqualFold(constant.Library, library) {
			return true
		}
	}
	return false
}

func knownAPITypeInDB(db *vbadb.DB, name string) bool {
	if db == nil {
		return false
	}
	if _, ok := db.ResolveType(name); ok {
		return true
	}
	// Built-in and generated databases represent enum types through their
	// constants' enum_group when no standalone TypeInfo exists.
	for _, constant := range db.Constants {
		if strings.EqualFold(constant.EnumGroup, name) {
			return true
		}
	}
	return false
}

func publicAPITypeName(typeText string) string {
	name := strings.TrimSpace(typeText)
	if strings.HasPrefix(strings.ToLower(name), "as ") {
		name = strings.TrimSpace(name[3:])
	}
	if strings.HasPrefix(strings.ToLower(name), "new ") {
		name = strings.TrimSpace(name[4:])
	}
	name = strings.Trim(name, "[]")
	for strings.HasSuffix(name, "()") {
		name = strings.TrimSpace(strings.TrimSuffix(name, "()"))
	}
	if i := strings.Index(name, " *"); i > 0 {
		name = strings.TrimSpace(name[:i])
	}
	return name
}

func isSafePublicAPIType(name string) bool {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "", "any", "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "object", "single", "string", "variant":
		return true
	default:
		return false
	}
}

func (a Analyzer) propertyVisibilityFindings(file parsedFile, index *apiTypeIndex, procedures []sourceProcedure) []Finding {
	type propertyAccessors struct {
		get   []int
		write []int
	}
	groups := map[string]*propertyAccessors{}
	for i, procedure := range file.IR.Procedures {
		if i >= len(procedures) || procedure.Symbol.IsEventHandler || !publicAPIModule(file.IR) {
			continue
		}
		name := strings.ToLower(procedure.Symbol.Name)
		if name == "" {
			continue
		}
		group := groups[name]
		if group == nil {
			group = &propertyAccessors{}
			groups[name] = group
		}
		switch procedure.Symbol.Kind {
		case procedureir.ProcedurePropertyGet:
			group.get = append(group.get, i)
		case procedureir.ProcedurePropertyLet, procedureir.ProcedurePropertySet:
			group.write = append(group.write, i)
		}
	}
	var findings []Finding
	for _, group := range groups {
		if len(group.get) == 0 || len(group.write) == 0 {
			continue
		}
		getter := file.IR.Procedures[group.get[0]].Symbol
		writer := file.IR.Procedures[group.write[0]].Symbol
		if !publicAPIProcedure(getter) && !publicAPIProcedure(writer) {
			continue
		}
		if procedureVisibilityRank(getter.Visibility) == procedureVisibilityRank(writer.Visibility) {
			continue
		}
		line := procedures[group.get[0]].StartLine
		if procedureVisibilityRank(getter.Visibility) < procedureVisibilityRank(writer.Visibility) {
			line = procedures[group.write[0]].StartLine
		}
		typeName := publicAPITypeName(getter.ReturnType)
		if typeName == "" {
			params := writer.Parameters
			if len(params) > 0 {
				typeName = publicAPITypeName(params[len(params)-1].Type)
			}
		}
		if typeName == "" {
			typeName = "Variant"
		}
		message := fmt.Sprintf("Property %s has mismatched getter/setter visibility for type %s (%s getter, %s setter).", getter.Name, typeName, effectiveVisibility(getter.Visibility), effectiveVisibility(writer.Visibility))
		reason := fmt.Sprintf("The public property %s exposes type %s through accessors with different effective visibility, so callers cannot rely on one stable public contract.", getter.Name, typeName)
		suggestion := fmt.Sprintf("Give both accessors the same visibility and keep type %s consistently exposed.", typeName)
		finding := a.simpleFinding(file, procedures[group.get[0]], line, publicAPITypeRule, "warning", message, reason, suggestion)
		findings = append(findings, finding)
	}
	return findings
}

func procedureVisibilityRank(visibility string) int {
	switch {
	case strings.EqualFold(visibility, "Private"):
		return 0
	case strings.EqualFold(visibility, "Friend"):
		return 1
	default:
		return 2
	}
}

func effectiveVisibility(visibility string) string {
	if strings.TrimSpace(visibility) == "" {
		return "Public"
	}
	return visibility
}
