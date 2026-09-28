package corpus

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
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
	"VBA243": {},
	"VBA251": {},
	"VBA252": {},
	"VBA260": {},
	"VBA261": {},
	"VBA262": {},
}

var configurableNonExcelLintRuleIDs = []string{"VB002", "VB003", "VB027"}

var configurableNonExcelAnalyzeRuleIDs = []string{
	"VBA201", "VBA203", "VBA205", "VBA215", "VBA216", "VBA217",
	"VBA218", "VBA221", "VBA225", "VBA226", "VBA238", "VBA242", "VBA243", "VBA251", "VBA252", "VBA260", "VBA261", "VBA262",
}

func applyProfilePolicy(cfg *config.Config, profile string) {
	if strings.EqualFold(profile, ProfileExcel) {
		// The Excel corpus explicitly opts into Excel-specific advisory rules;
		// production defaults remain unchanged.
		cfg.Analyze.DetectExpensiveFullRangeOperations = true
		cfg.Analyze.DetectValue2PerformanceOpportunities = true
		cfg.Analyze.DetectApplicationWorksheetFunction = true
		cfg.Analyze.DetectHostBracketExpressions = true
		return
	}
	cfg.Lint.DisabledRules = append([]string(nil), configurableNonExcelLintRuleIDs...)
	cfg.Analyze.DisabledRules = append([]string(nil), configurableNonExcelAnalyzeRuleIDs...)
}

// applyCorpusReviewPolicy is the corpus review evaluation profile. It enables
// the new opt-in rule family (VB067-VB092 lint rules and the VBA253-VBA283
// analyzer wave, plus previously reviewed opt-in rules kept enabled for
// continuity) so the committed corpus exercises them for promotion review.
// Rules in corpusReviewExcludedRules stay disabled, and host-profile
// exclusions still apply, so non-Excel projects keep their documented rule
// omissions. This policy belongs to corpus evidence only and never alters
// production defaults.
func applyCorpusReviewPolicy(cfg *config.Config, profile string) {
	for _, id := range corpusReviewRuleIDs {
		if profileExcludes(profile, id) {
			continue
		}
		rule, ok := staticrules.Lookup(id)
		if !ok {
			continue
		}
		switch rule.Family {
		case staticrules.FamilyLint:
			config.SetLintRuleEnabled(&cfg.Lint, id, true)
		case staticrules.FamilyAnalyze:
			config.SetAnalyzeRuleEnabled(&cfg.Analyze, id, true)
		}
	}
}

// corpusReviewRuleIDs is the corpus review evaluation profile: the opt-in
// rules the generated third-party workspace config enables so real-world
// evidence exists for promotion review. Host-profile exclusions are applied
// per project. Entries must resolve to configurable, default-disabled rules.
var corpusReviewRuleIDs = []string{
	// New lint family VB067-VB092. VB085 is excluded: config validation makes
	// require_explicit_step mutually exclusive with VB084
	// (detect_redundant_step_one), so it relies on focused fixtures.
	"VB067", "VB068", "VB069", "VB070", "VB071", "VB072", "VB073", "VB074",
	"VB075", "VB076", "VB077", "VB078", "VB079", "VB080", "VB081", "VB082",
	"VB083", "VB084", "VB086", "VB087", "VB088", "VB089", "VB090", "VB091",
	"VB092",
	// Opt-in analyzer rules, including the VBA253-VBA283 wave. VBA277 is
	// excluded: detect_redundant_byref_modifiers is mutually exclusive with
	// VBA273 (detect_implicit_byref_parameters), so it relies on focused
	// fixtures. VBA260/VBA261/VBA262 fire only for the excel host profile.
	"VBA240",
	"VBA253", "VBA254", "VBA255", "VBA256", "VBA257", "VBA258", "VBA259",
	"VBA260", "VBA261", "VBA262",
	"VBA265", "VBA266", "VBA267", "VBA268", "VBA269", "VBA270", "VBA271",
	"VBA272", "VBA273", "VBA274", "VBA275", "VBA278", "VBA279",
	"VBA280", "VBA281", "VBA282",
}

// corpusReviewExcludedRules documents opt-in rules that intentionally stay
// out of the review profile, with the reason.
var corpusReviewExcludedRules = map[string]string{
	"VB018":  "pre-existing opt-in outside the reviewed family; per-project convention",
	"VB021":  "pre-existing opt-in outside the reviewed family",
	"VB044":  "requires a configured procedure-name constant convention",
	"VB085":  "mutually exclusive with VB084 under config validation",
	"VBA207": "pre-existing opt-in outside the reviewed family",
	"VBA210": "pre-existing opt-in outside the reviewed family",
	"VBA213": "pre-existing opt-in outside the reviewed family",
	"VBA248": "pre-existing opt-in outside the reviewed family",
	"VBA277": "mutually exclusive with VBA273 under config validation",
}

func profileExcludes(profile, code string) bool {
	if strings.EqualFold(profile, ProfileExcel) {
		return false
	}
	_, ok := nonExcelRuleIDs[strings.ToUpper(strings.TrimSpace(code))]
	return ok
}
