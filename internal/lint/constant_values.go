package lint

import (
	"context"
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
			result := constexpr.Evaluate(expression, constexpr.Values(values))
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
				for _, qualified := range item.qualified {
					values[strings.ToLower(qualified)] = result.Typed
				}
				if ir != nil && ir.ModuleName != "" {
					values[strings.ToLower(cleanIdentifier(ir.ModuleName))+"."+key] = result.Typed
				}
				changed = true
			}
			if item.enumGroup != "" {
				if result.Typed.Integer == int64(^uint64(0)>>1) {
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

func (l Linter) projectConstantValuesContext(ctx context.Context, files []string) map[string]constexpr.Value {
	values := make(map[string]constexpr.Value)
	for _, path := range files {
		if err := ctx.Err(); err != nil {
			return values
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
		fileValues := constantValuesFromSource(string(source), &ir, nil)
		standard := strings.EqualFold(ir.ModuleKind, "standard")
		for _, declaration := range ir.Declarations {
			if !declaration.IsConst && !procedureir.IsConstKind(declaration.Kind) {
				continue
			}
			if !strings.EqualFold(declaration.Visibility, "public") && !strings.EqualFold(declaration.Visibility, "friend") {
				continue
			}
			value, ok := fileValues[strings.ToLower(cleanIdentifier(declaration.Name))]
			if !ok {
				continue
			}
			if standard {
				values[strings.ToLower(cleanIdentifier(declaration.Name))] = value
			}
			if ir.ModuleName != "" {
				values[strings.ToLower(cleanIdentifier(ir.ModuleName))+"."+strings.ToLower(cleanIdentifier(declaration.Name))] = value
			}
			if declaration.Parent != "" {
				values[strings.ToLower(cleanIdentifier(declaration.Parent))+"."+strings.ToLower(cleanIdentifier(declaration.Name))] = value
			}
		}
	}
	if result, err := typedb.LoadForRuntime(""); err == nil && result.DB != nil {
		constants := result.DB.AllConstantsList()
		counts := make(map[string]int)
		for _, constant := range constants {
			name := strings.ToLower(cleanIdentifier(constant.Name))
			if name != "" {
				counts[name]++
			}
		}
		for _, constant := range constants {
			value, ok := staticTypeLibConstantValue(constant)
			if !ok {
				continue
			}
			name := strings.ToLower(cleanIdentifier(constant.Name))
			if name != "" && counts[name] == 1 {
				values[name] = value
			}
			if group := strings.ToLower(cleanIdentifier(constant.EnumGroup)); group != "" && name != "" {
				values[group+"."+name] = value
			}
			if library := strings.ToLower(cleanIdentifier(constant.Library)); library != "" && name != "" {
				values[library+"."+name] = value
			}
		}
	}
	return values
}

func staticTypeLibConstantValue(constant vbadb.ConstantInfo) (constexpr.Value, bool) {
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
