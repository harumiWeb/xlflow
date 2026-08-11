package corpus

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

// These rules depend on the Excel object model. The policy intentionally
// lives in the corpus adapter rather than the shared rule registry because it
// describes evidence selection for third-party fixtures, not analyzer meaning.
var nonExcelRuleIDs = map[string]struct{}{
	"VB002": {}, "VB003": {}, "VB027": {},
	"VBA104": {}, "VBA201": {}, "VBA203": {}, "VBA205": {},
	"VBA211": {}, "VBA215": {}, "VBA216": {}, "VBA217": {},
	"VBA218": {}, "VBA221": {}, "VBA225": {}, "VBA226": {}, "VBA238": {},
	"VBA242": {},
}

var configurableNonExcelLintRuleIDs = []string{"VB002", "VB003", "VB027"}

var configurableNonExcelAnalyzeRuleIDs = []string{
	"VBA201", "VBA203", "VBA205", "VBA215", "VBA216", "VBA217",
	"VBA218", "VBA221", "VBA225", "VBA226", "VBA238", "VBA242",
}

func applyProfilePolicy(cfg *config.Config, profile string) {
	if strings.EqualFold(profile, ProfileExcel) {
		// The Excel corpus explicitly opts into Excel-specific advisory rules;
		// production defaults remain unchanged.
		cfg.Analyze.DetectExpensiveFullRangeOperations = true
		return
	}
	cfg.Lint.DisabledRules = append([]string(nil), configurableNonExcelLintRuleIDs...)
	cfg.Analyze.DisabledRules = append([]string(nil), configurableNonExcelAnalyzeRuleIDs...)
}

func profileExcludes(profile, code string) bool {
	if strings.EqualFold(profile, ProfileExcel) {
		return false
	}
	_, ok := nonExcelRuleIDs[strings.ToUpper(strings.TrimSpace(code))]
	return ok
}
