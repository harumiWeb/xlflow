package lint

import (
	"context"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

// constantValuesFromSource projects the module-level Const and Enum values
// into the shared evaluator's immutable value map. It deliberately records
// only Known values; an unresolved or conditional declaration remains visible
// through the name-only map and therefore cannot accidentally become a value.
func constantValuesFromSource(source string, ir *procedureir.DocumentIR, base map[string]constexpr.Value) map[string]constexpr.Value {
	values := make(map[string]constexpr.Value, len(base)+8)
	for name, value := range base {
		values[name] = value
	}
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	insideProcedure := make([]bool, len(lines)+1)
	if ir != nil {
		delta := make([]int, len(lines)+2)
		for _, procedure := range ir.Procedures {
			start := procedure.Symbol.DeclarationRange.StartLine
			end := procedure.Symbol.DeclarationRange.EndLine
			if start < 1 {
				start = 1
			}
			if end > len(lines) {
				end = len(lines)
			}
			if start <= end {
				delta[start]++
				delta[end+1]--
			}
		}
		active := 0
		for line := 1; line <= len(lines); line++ {
			active += delta[line]
			insideProcedure[line] = active > 0
		}
	} else {
		markProcedureLines(lines, insideProcedure)
	}
	type pending struct {
		name, expression string
		qualified        []string
		enumGroup        string
		implicit         bool
	}
	var pendingConsts []pending
	inEnum := false
	enumName := ""
	conditionalDepth := 0
	for lineNo, line := range lines {
		trim := strings.TrimSpace(strings.SplitN(line, "'", 2)[0])
		lower := strings.ToLower(trim)
		if strings.HasPrefix(lower, "#if ") {
			conditionalDepth++
			continue
		}
		if strings.HasPrefix(lower, "#elseif ") || strings.HasPrefix(lower, "#else") {
			continue
		}
		if strings.HasPrefix(lower, "#end if") {
			if conditionalDepth > 0 {
				conditionalDepth--
			}
			continue
		}
		if conditionalDepth > 0 || insideProcedure[lineNo+1] {
			continue
		}
		switch {
		case strings.HasPrefix(lower, "enum ") || strings.HasPrefix(lower, "public enum ") || strings.HasPrefix(lower, "private enum ") || strings.HasPrefix(lower, "friend enum "):
			inEnum = true
			enumName = ""
			fields := strings.Fields(trim)
			for i, field := range fields {
				if strings.EqualFold(field, "enum") && i+1 < len(fields) {
					enumName = cleanIdentifier(fields[i+1])
					break
				}
			}
			continue
		case inEnum && strings.HasPrefix(lower, "end enum"):
			inEnum = false
			enumName = ""
			continue
		}
		if inEnum && trim != "" {
			name, expression := enumMemberConstant(trim, nil)
			if name == "" {
				continue
			}
			item := pending{
				name: name, expression: expression, enumGroup: enumName,
				implicit: !strings.Contains(trim, "="),
			}
			if enumName != "" {
				item.qualified = []string{enumName + "." + name}
			}
			pendingConsts = append(pendingConsts, item)
			continue
		}
		if name, expression, ok := arrayShapeConstLine(trim); ok {
			pendingConsts = append(pendingConsts, pending{name: cleanIdentifier(name), expression: expression})
		}
	}
	// Resolve forward references without making map iteration order observable.
	for pass := 0; pass < len(pendingConsts); pass++ {
		changed := false
		environment := constexpr.NewValues(values)
		enumNext := make(map[string]int64)
		enumSeen := make(map[string]bool)
		for _, item := range pendingConsts {
			if item.name == "" {
				continue
			}
			expression := item.expression
			if item.enumGroup != "" {
				group := strings.ToLower(item.enumGroup)
				if item.implicit {
					if !enumSeen[group] {
						expression = "0"
					} else if next, ok := enumNext[group]; ok {
						expression = strconv.FormatInt(next, 10)
					} else {
						expression = "MissingEnumValue"
					}
				}
				enumSeen[group] = true
			}
			result := constexpr.Evaluate(expression, environment)
			if result.Kind != constexpr.Known {
				if item.enumGroup != "" {
					delete(enumNext, strings.ToLower(item.enumGroup))
				}
				continue
			}
			if item.enumGroup != "" && !isIntegralConstant(result.Typed) {
				delete(enumNext, strings.ToLower(item.enumGroup))
				continue
			}
			key := strings.ToLower(item.name)
			if prior, exists := values[key]; !exists || prior != result.Typed {
				values[key] = result.Typed
				environment[key] = result.Typed
				for _, qualified := range item.qualified {
					qualifiedKey := strings.ToLower(qualified)
					values[qualifiedKey] = result.Typed
					environment[qualifiedKey] = result.Typed
				}
				if ir != nil && ir.ModuleName != "" {
					moduleKey := strings.ToLower(cleanIdentifier(ir.ModuleName)) + "." + key
					values[moduleKey] = result.Typed
					environment[moduleKey] = result.Typed
				}
				changed = true
			}
			if item.enumGroup != "" {
				if result.Typed.Integer == math.MaxInt64 {
					delete(enumNext, strings.ToLower(item.enumGroup))
				} else {
					enumNext[strings.ToLower(item.enumGroup)] = result.Typed.Integer + 1
				}
			}
		}
		if !changed {
			break
		}
	}
	return values
}

// markProcedureLines conservatively identifies procedure bodies when callers
// do not have a parsed DocumentIR. It keeps the source-only projection useful
// while ensuring procedure-local Const declarations cannot leak into the
// module-level environment.
func markProcedureLines(lines []string, inside []bool) {
	depth := 0
	for lineNo, line := range lines {
		trim := strings.TrimSpace(strings.SplitN(line, "'", 2)[0])
		lower := strings.ToLower(trim)
		fields := strings.Fields(lower)
		if depth > 0 {
			inside[lineNo+1] = true
			if len(fields) >= 2 && fields[0] == "end" && (fields[1] == "sub" || fields[1] == "function" || fields[1] == "property") {
				depth = 0
			}
			continue
		}
		if len(fields) == 0 || fields[0] == "declare" || fields[0] == "end" {
			continue
		}
		for _, field := range fields {
			if field == "sub" || field == "function" || field == "property" {
				depth = 1
				inside[lineNo+1] = true
				break
			}
		}
	}
}

func isIntegralConstant(value constexpr.Value) bool {
	switch value.Kind {
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong:
		return true
	default:
		return false
	}
}

// ConstantValuesFromSource exposes the deterministic module projection for
// batch analyzers that already own the parsed DocumentIR.
func ConstantValuesFromSource(source string, ir *procedureir.DocumentIR, base map[string]constexpr.Value) map[string]constexpr.Value {
	return constantValuesFromSource(source, ir, base)
}

// ConstantValueDocument is the immutable source/IR pair used to build a
// project-wide constant environment.
type ConstantValueDocument struct {
	Source string
	IR     *procedureir.DocumentIR
}

// TypeLibConstantValue converts a static TypeLib record without invoking VBA
// or Excel. Unsupported and malformed records remain unknown to callers.
func TypeLibConstantValue(constant vbadb.ConstantInfo) (constexpr.Value, bool) {
	switch strings.ToLower(strings.TrimSpace(constant.Type)) {
	case "string":
		return constexpr.Value{Kind: constexpr.ValueString, String: constant.Value}, true
	case "boolean":
		if strings.EqualFold(strings.TrimSpace(constant.Value), "true") {
			return constexpr.Value{Kind: constexpr.ValueBoolean, Boolean: true}, true
		}
		if strings.EqualFold(strings.TrimSpace(constant.Value), "false") {
			return constexpr.Value{Kind: constexpr.ValueBoolean}, true
		}
	}
	result := constexpr.Evaluate(constant.Value, nil)
	if result.Kind != constexpr.Known {
		return constexpr.Value{}, false
	}
	return result.Typed, true
}

// ProjectConstantValues builds a deterministic, ambiguity-free project
// environment. Qualified source and TypeLib names are retained only when
// their qualifier is unique; unqualified names are retained only when there
// is one visible declaration. Repeated passes resolve cross-module forward
// references without making map iteration order observable.
func ProjectConstantValues(documents []ConstantValueDocument, typeDB *vbadb.DB) map[string]constexpr.Value {
	type candidate struct {
		identity string
		value    constexpr.Value
		known    bool
	}
	add := func(byKey map[string][]candidate, key, identity string, value constexpr.Value, known bool) {
		key = normalizeConstantValueKey(key)
		if key == "" {
			return
		}
		for _, prior := range byKey[key] {
			if prior.identity == identity {
				return
			}
		}
		byKey[key] = append(byKey[key], candidate{identity: identity, value: value, known: known})
	}
	unique := func(byKey map[string][]candidate) map[string]constexpr.Value {
		values := make(map[string]constexpr.Value, len(byKey))
		for key, candidates := range byKey {
			if len(candidates) == 1 && candidates[0].known {
				values[key] = candidates[0].value
			}
		}
		return values
	}
	equal := func(left, right map[string]constexpr.Value) bool {
		if len(left) != len(right) {
			return false
		}
		for key, value := range left {
			if other, ok := right[key]; !ok || other != value {
				return false
			}
		}
		return true
	}
	environment := map[string]constexpr.Value{}
	passes := len(documents) + 2
	if passes < 2 {
		passes = 2
	}
	for pass := 0; pass < passes; pass++ {
		byKey := make(map[string][]candidate)
		if typeDB != nil {
			for index, constant := range typeDB.AllConstantsList() {
				value, ok := TypeLibConstantValue(constant)
				identity := "typelib:" + strconv.Itoa(index)
				name := cleanIdentifier(constant.Name)
				add(byKey, name, identity, value, ok)
				add(byKey, cleanIdentifier(constant.EnumGroup)+"."+name, identity, value, ok)
				add(byKey, cleanIdentifier(constant.Library)+"."+name, identity, value, ok)
			}
		}
		for documentIndex, document := range documents {
			if document.IR == nil {
				continue
			}
			fileValues := constantValuesFromSource(document.Source, document.IR, environment)
			standard := strings.EqualFold(document.IR.ModuleKind, "standard")
			for declarationIndex, declaration := range document.IR.Declarations {
				if !declaration.IsConst && !procedureir.IsConstKind(declaration.Kind) {
					continue
				}
				if !strings.EqualFold(declaration.Visibility, "public") && !strings.EqualFold(declaration.Visibility, "friend") {
					continue
				}
				name := cleanIdentifier(declaration.Name)
				value, ok := fileValues[normalizeConstantValueKey(name)]
				if name == "" {
					continue
				}
				identity := "source:" + strconv.Itoa(documentIndex) + ":" + strconv.Itoa(declarationIndex)
				if standard {
					add(byKey, name, identity, value, ok)
				}
				add(byKey, cleanIdentifier(document.IR.ModuleName)+"."+name, identity, value, ok)
				add(byKey, cleanIdentifier(declaration.Parent)+"."+name, identity, value, ok)
			}
		}
		next := unique(byKey)
		if equal(environment, next) {
			return next
		}
		environment = next
	}
	return environment
}

func normalizeConstantValueKey(text string) string {
	parts := strings.Split(text, ".")
	for index, part := range parts {
		parts[index] = strings.ToLower(cleanIdentifier(part))
		if parts[index] == "" {
			return ""
		}
	}
	return strings.Join(parts, ".")
}

func (l Linter) projectConstantValuesContext(ctx context.Context, files []string) map[string]constexpr.Value {
	var documents []ConstantValueDocument
	for _, path := range files {
		if err := ctx.Err(); err != nil {
			return ProjectConstantValues(documents, nil)
		}
		source, err := os.ReadFile(path)
		if err != nil {
			continue
		}
		parsed, err := ast.ParseDocument(path, source)
		if err != nil {
			continue
		}
		ir, err := procedureir.BuildParsedContext(ctx, procedureir.BuildOptions{
			RootDir: l.RootDir, Path: path, ModuleKind: l.moduleKindForPath(path),
		}, parsed)
		parsed.Close()
		if err != nil {
			continue
		}
		documents = append(documents, ConstantValueDocument{Source: string(source), IR: &ir})
	}
	var typeDB *vbadb.DB
	if result, err := typedb.LoadForRuntime(""); err == nil {
		typeDB = result.DB
	}
	return ProjectConstantValues(documents, typeDB)
}
