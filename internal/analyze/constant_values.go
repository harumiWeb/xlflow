package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func projectConstantValues(files []parsedFile, typeDB *vbadb.DB) map[string]constexpr.Value {
	values := make(map[string]constexpr.Value)
	for _, file := range files {
		fileValues := file.ConstantValues
		if fileValues == nil {
			fileValues = lint.ConstantValuesFromSource(string(file.Source), &file.IR, nil)
		}
		standard := strings.EqualFold(file.ModuleKind, "standard")
		for _, declaration := range file.IR.Declarations {
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
			name := strings.ToLower(cleanIdentifier(declaration.Name))
			if standard {
				values[name] = value
			}
			if file.IR.ModuleName != "" {
				values[strings.ToLower(cleanIdentifier(file.IR.ModuleName))+"."+name] = value
			}
			if declaration.Parent != "" {
				values[strings.ToLower(cleanIdentifier(declaration.Parent))+"."+name] = value
			}
		}
	}
	if typeDB == nil {
		return values
	}
	all := typeDB.AllConstantsList()
	counts := make(map[string]int)
	for _, constant := range all {
		name := strings.ToLower(cleanIdentifier(constant.Name))
		if name != "" {
			counts[name]++
		}
	}
	for _, constant := range all {
		value, ok := typeLibConstantValue(constant)
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
	return values
}

func typeLibConstantValue(constant vbadb.ConstantInfo) (constexpr.Value, bool) {
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
