package analyze

import (
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

var projectDefaultMemberAttributeRE = regexp.MustCompile(`(?im)^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\.VB_UserMemId\s*=\s*0\s*$`)

func defaultMemberAnalysisEnabled(cfg config.AnalyzeConfig) bool {
	return cfg.DetectImplicitDefaultMemberAccess || cfg.DetectUnboundDefaultMemberAccess || cfg.DetectBangNotation
}

// projectDefaultMemberIndex derives value-producing project defaults from the
// attributes preserved in exported class-like modules. The index belongs to
// the analyzed revision, so unsaved realtime siblings and in-memory projects
// use the same source of truth as batch analysis.
func projectDefaultMemberIndex(files []parsedFile) (map[string]vbadb.MemberInfo, map[string]bool) {
	members := make(map[string]vbadb.MemberInfo)
	types := make(map[string]bool)
	ambiguous := make(map[string]bool)
	for _, file := range files {
		if strings.EqualFold(strings.TrimSpace(file.IR.ModuleKind), "standard") {
			continue
		}
		typeName := strings.ToLower(strings.TrimSpace(file.IR.ModuleName))
		if typeName == "" {
			continue
		}
		types[typeName] = true
		matches := projectDefaultMemberAttributeRE.FindAllStringSubmatch(string(file.Source), -1)
		for _, match := range matches {
			if len(match) != 2 {
				continue
			}
			candidate, ok := projectValueProcedure(file.IR.Procedures, match[1])
			if !ok {
				continue
			}
			if existing, exists := members[typeName]; exists && !strings.EqualFold(existing.Name, candidate.Name) {
				ambiguous[typeName] = true
				delete(members, typeName)
				continue
			}
			if !ambiguous[typeName] {
				members[typeName] = candidate
			}
		}
	}
	return members, types
}

func projectValueProcedure(procedures []procedureir.ProcedureIR, name string) (vbadb.MemberInfo, bool) {
	for _, procedure := range procedures {
		symbol := procedure.Symbol
		if symbol.Recovered || !strings.EqualFold(symbol.Name, name) {
			continue
		}
		if symbol.Kind != procedureir.ProcedureFunction && symbol.Kind != procedureir.ProcedurePropertyGet {
			continue
		}
		parameters := make([]vbadb.ParamInfo, 0, len(symbol.Parameters))
		for _, parameter := range symbol.Parameters {
			parameters = append(parameters, vbadb.ParamInfo{
				Name: parameter.Name, Type: parameter.Type,
				Optional: parameter.Optional, ParamArray: parameter.ParamArray,
			})
		}
		return vbadb.MemberInfo{
			Name: symbol.Name, ReturnType: symbol.ReturnType,
			Parameters: parameters, Default: true,
		}, true
	}
	return vbadb.MemberInfo{}, false
}
