package analyze

import (
	"cmp"
	"math"
	"regexp"
	"slices"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// SelectCaseUnreachableContext is the stable, machine-readable evidence attached
// to a VBA259 finding. The registry metadata remains the source of truth for
// severity; this context only identifies why a Case item or Case Else branch can
// never be reached.
type SelectCaseUnreachableContext struct {
	// Kind is "duplicate", "covered", "empty_range", "impossible", or
	// "else_covered".
	Kind string `json:"kind"`
	// Item is the Case item source text; it is empty for else_covered.
	Item string `json:"item,omitempty"`
	// CoveredByLine is the earliest earlier Case item line that participates
	// in the coverage proof for duplicate and covered findings.
	CoveredByLine int `json:"covered_by_line,omitempty"`
}

// selectNumInterval is a closed numeric interval of selector values. ±Inf marks
// an open end; strict bounds are normalized into the closed form with
// math.Nextafter so comparisons stay exact.
type selectNumInterval struct {
	lo, hi float64
	line   int
}

// selectStrInterval is a string interval under Option Compare Binary code-unit
// order. hasLo/hasHi mark unbounded ends; loExcl/hiExcl mark strict bounds
// because UTF-16 order has no Nextafter equivalent.
type selectStrInterval struct {
	lo, hi         string
	hasLo, hasHi   bool
	loExcl, hiExcl bool
	line           int
}

// selectItemSet is the set of selector values a single Case item can match.
// opaque means the item cannot be modeled deterministically and must be skipped
// silently; empty means the item is provably empty (an inverted To range).
type selectItemSet struct {
	num    []selectNumInterval
	str    []selectStrInterval
	line   int
	opaque bool
	empty  bool
}

// selectCaseDomain describes the set of values a Select Case selector can
// produce. elseCoverable is true only when the domain is finite and enumerable,
// so a Case Else can be proven unreachable.
type selectCaseDomain struct {
	numeric, str, integral, elseCoverable, opaque bool
	numUniverse                                   []selectNumInterval // nil = unbounded
	strUniverse                                   []selectStrInterval // nil = unbounded
	typeName                                      string
}

// selectCoverage is the normalized disjoint union of earlier Case item sets,
// kept sorted so containment checks only ever test a single union member.
// integral mirrors the selector domain: on integral domains every filtered
// bound is an exact integer, so members separated by less than one whole step
// still merge, while non-integral domains fall back to float64-ulp adjacency.
type selectCoverage struct {
	num      []selectNumInterval
	str      []selectStrInterval
	integral bool
}

var (
	selectCaseIsRe       = regexp.MustCompile(`^\s*[Ii][Ss]\s*(<=|>=|<>|=|<|>)`)
	selectCaseIsWordRe   = regexp.MustCompile(`(?i)^\s*is\b`)
	selectCaseElseWordRe = regexp.MustCompile(`(?i)^\s*case\s+else`)
	selectCaseIdentRe    = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)
)

// unreachableSelectCaseFindings reports VBA259 findings for each Select Case
// statement in the procedure. The analysis is deliberately fail-open: any
// recovered syntax, conditional-compilation branch, or value the shared
// constant evaluator cannot prove leaves the affected Select Case silent.
func (a Analyzer) unreachableSelectCaseFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) []Finding {
	if enabled, known := config.AnalyzeRuleEnabled(a.Config.Analyze, "VBA259"); !known || !enabled {
		return nil
	}
	facts := proc.analysisFacts()
	env := a.procedureConstantEnvironment(file, proc)
	compare := fileOptionCompare(file)

	children := make(map[int][]procedureir.Statement)
	var selects []procedureir.Statement
	for statement := range proc.Statements.All() {
		if statement.Kind == procedureir.StatementSelect && statement.SyntaxKind == "select_statement" {
			selects = append(selects, statement)
		}
		if statement.ParentID != 0 {
			children[statement.ParentID] = append(children[statement.ParentID], statement)
		}
	}
	if len(selects) == 0 {
		return nil
	}
	slices.SortFunc(selects, selectStatementCompare)

	var findings []Finding
	for _, sel := range selects {
		findings = append(findings, a.selectCaseStatementFindings(file, proc, sel, children, facts, env, compare, moduleDecls)...)
	}
	return findings
}

func selectStatementCompare(a, b procedureir.Statement) int {
	return cmp.Or(cmp.Compare(a.Range.StartByte, b.Range.StartByte), cmp.Compare(a.ID, b.ID))
}

// selectCaseChildren returns the direct children of parent that carry the given
// case syntax kind, sorted in source order. Nested Select Case statements have
// different parents, so this cannot leak items across Select boundaries.
func selectCaseChildren(statements []procedureir.Statement, syntaxKind string) []procedureir.Statement {
	var out []procedureir.Statement
	for _, statement := range statements {
		if statement.Kind == procedureir.StatementCase && statement.SyntaxKind == syntaxKind {
			out = append(out, statement)
		}
	}
	slices.SortFunc(out, selectStatementCompare)
	return out
}

func (a Analyzer) selectCaseStatementFindings(file parsedFile, proc sourceProcedure, sel procedureir.Statement, children map[int][]procedureir.Statement, facts *procedureAnalysisFacts, env constexpr.Environment, compare string, moduleDecls map[string]sourceDeclaration) []Finding {
	if sel.Recovered || len(sel.ConditionalBranches) > 0 {
		return nil
	}
	selector := sel.Value
	if selector == nil && len(sel.ExpressionIDs) > 0 {
		if expression, ok := facts.Expression(sel.ExpressionIDs[0]); ok {
			selector = &expression
		}
	}
	if selector == nil || selector.Recovered {
		return nil
	}
	clauses := selectCaseChildren(children[sel.ID], "case_clause")
	// The clause list is the coverage proof's backbone: when any clause or item
	// is recovered or sits inside a #If branch the list may be incomplete, so
	// fail open for the whole Select rather than risk a wrong shadowing claim.
	itemsByClause := make(map[int][]procedureir.Statement, len(clauses))
	for _, clause := range clauses {
		if clause.Recovered || len(clause.ConditionalBranches) > 0 {
			return nil
		}
		items := selectCaseChildren(children[clause.ID], "case_expression")
		for _, item := range items {
			if item.Recovered || len(item.ConditionalBranches) > 0 {
				return nil
			}
		}
		itemsByClause[clause.ID] = items
	}
	domain := selectCaseSelectorDomain(*selector, proc, env, compare, moduleDecls)
	if domain.opaque {
		return nil
	}

	coverage := selectCoverage{integral: domain.integral}
	earlier := make([]selectItemSet, 0, 4)
	var findings []Finding
	for _, clause := range clauses {
		if clause.Control != nil && clause.Control.CaseElse {
			if domain.elseCoverable && coverage.coversUniverse(domain) {
				findings = append(findings, a.selectCaseElseFinding(file, proc, clause, domain))
			}
			// VB063 owns everything after Case Else, including a second Case
			// Else; stop rather than double-report invalid syntax.
			break
		}
		for _, item := range itemsByClause[clause.ID] {
			set := classifySelectCaseItem(item, facts, env, compare)
			if set.empty {
				findings = append(findings, a.selectCaseItemFinding(file, proc, item, "empty_range", 0, domain))
				continue
			}
			if set.opaque {
				continue
			}
			eff := domain.filter(set)
			if eff.opaque {
				continue
			}
			if len(eff.num) == 0 && len(eff.str) == 0 {
				findings = append(findings, a.selectCaseItemFinding(file, proc, item, "impossible", 0, domain))
				continue
			}
			if coveredBy, duplicate := selectEarlierDuplicate(earlier, eff); duplicate {
				findings = append(findings, a.selectCaseItemFinding(file, proc, item, "duplicate", coveredBy, domain))
				continue
			}
			if witness, covered := coverage.contains(eff); covered {
				findings = append(findings, a.selectCaseItemFinding(file, proc, item, "covered", witness, domain))
				earlier = append(earlier, eff)
				continue
			}
			coverage.add(eff)
			earlier = append(earlier, eff)
		}
	}
	return findings
}

// classifySelectCaseItem converts one case_expression statement into the set of
// selector values it matches. The wrapper expression's children carry the
// operands: two for a `To` range, one for a point or an `Is` comparison.
func classifySelectCaseItem(item procedureir.Statement, facts *procedureAnalysisFacts, env constexpr.Environment, compare string) selectItemSet {
	set := selectItemSet{line: item.Range.StartLine}
	if len(item.ExpressionIDs) == 0 {
		set.opaque = true
		return set
	}
	wrapper, ok := facts.Expression(item.ExpressionIDs[0])
	if !ok || wrapper.Recovered || len(wrapper.Children) == 0 || len(wrapper.Children) > 2 {
		set.opaque = true
		return set
	}
	operands := make([]procedureir.Expression, 0, len(wrapper.Children))
	for _, id := range wrapper.Children {
		operand, ok := facts.Expression(id)
		if !ok || operand.Recovered {
			set.opaque = true
			return set
		}
		operands = append(operands, operand)
	}
	if match := selectCaseIsRe.FindStringSubmatch(item.Text); match != nil {
		if len(operands) != 1 {
			set.opaque = true
			return set
		}
		return selectCaseRaySet(match[1], operands[0].Text, env, compare, set.line)
	}
	if selectCaseIsWordRe.MatchString(item.Text) {
		// An `Is` keyword without a recognized comparison operator cannot be
		// modeled deterministically.
		set.opaque = true
		return set
	}
	if len(operands) == 2 {
		return selectCaseRangeSet(operands[0].Text, operands[1].Text, env, compare, set.line)
	}
	return selectCasePointSet(operands[0].Text, env, compare, set.line)
}

// selectConstValue is a statically known Case operand: either a float64 number
// or a string. Boolean inputs are normalized to VBA's -1/0 numeric form.
type selectConstValue struct {
	num   float64
	str   string
	isNum bool
	isStr bool
	ok    bool
}

func evalSelectConst(text string, env constexpr.Environment) selectConstValue {
	result := constexpr.Evaluate(text, env)
	if result.Kind != constexpr.Known {
		return selectConstValue{}
	}
	switch result.Typed.Kind {
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong:
		if result.Typed.Integer > 1<<53 || result.Typed.Integer < -(1<<53) {
			// float64 cannot represent these exactly; interval edges must stay
			// precise or coverage could be claimed for a value that differs.
			return selectConstValue{}
		}
		return selectConstValue{num: float64(result.Typed.Integer), isNum: true, ok: true}
	case constexpr.ValueSingle, constexpr.ValueDouble:
		if math.IsNaN(result.Typed.Float) || math.IsInf(result.Typed.Float, 0) {
			return selectConstValue{}
		}
		return selectConstValue{num: result.Typed.Float, isNum: true, ok: true}
	case constexpr.ValueCurrency:
		if result.Typed.Currency > 1<<53 || result.Typed.Currency < -(1<<53) {
			return selectConstValue{}
		}
		return selectConstValue{num: float64(result.Typed.Currency) / 10000, isNum: true, ok: true}
	case constexpr.ValueBoolean:
		if result.Typed.Boolean {
			return selectConstValue{num: -1, isNum: true, ok: true}
		}
		return selectConstValue{num: 0, isNum: true, ok: true}
	case constexpr.ValueString:
		return selectConstValue{str: result.Typed.String, isStr: true, ok: true}
	default:
		// Date, Empty, Null, and Nothing are not modeled.
		return selectConstValue{}
	}
}

func selectCasePointSet(text string, env constexpr.Environment, compare string, line int) selectItemSet {
	set := selectItemSet{line: line}
	value := evalSelectConst(text, env)
	switch {
	case !value.ok:
		set.opaque = true
	case value.isNum:
		set.num = []selectNumInterval{{lo: value.num, hi: value.num, line: line}}
	default:
		interval, ok := selectStrPointInterval(value.str, compare, line)
		if !ok {
			set.opaque = true
			return set
		}
		set.str = []selectStrInterval{interval}
	}
	return set
}

func selectCaseRangeSet(loText, hiText string, env constexpr.Environment, compare string, line int) selectItemSet {
	set := selectItemSet{line: line}
	lo, hi := evalSelectConst(loText, env), evalSelectConst(hiText, env)
	if !lo.ok || !hi.ok {
		set.opaque = true
		return set
	}
	if lo.isNum && hi.isNum {
		if lo.num > hi.num {
			set.empty = true
			return set
		}
		set.num = []selectNumInterval{{lo: lo.num, hi: hi.num, line: line}}
		return set
	}
	if lo.isStr && hi.isStr {
		// Option Compare Text/Database ordering cannot be reproduced with Go
		// string order, so only Binary ranges are modeled.
		if compare != "binary" || !selectStrSupported(lo.str) || !selectStrSupported(hi.str) {
			set.opaque = true
			return set
		}
		if lo.str > hi.str {
			set.empty = true
			return set
		}
		set.str = []selectStrInterval{{lo: lo.str, hi: hi.str, hasLo: true, hasHi: true, line: line}}
		return set
	}
	// Mixed-type ranges are left to VBA's runtime coercion rules.
	set.opaque = true
	return set
}

func selectCaseRaySet(operator, operandText string, env constexpr.Environment, compare string, line int) selectItemSet {
	set := selectItemSet{line: line}
	value := evalSelectConst(operandText, env)
	if !value.ok {
		set.opaque = true
		return set
	}
	if value.isNum {
		v := value.num
		switch operator {
		case "=":
			set.num = []selectNumInterval{{lo: v, hi: v, line: line}}
		case "<":
			set.num = []selectNumInterval{{lo: math.Inf(-1), hi: math.Nextafter(v, math.Inf(-1)), line: line}}
		case "<=":
			set.num = []selectNumInterval{{lo: math.Inf(-1), hi: v, line: line}}
		case ">":
			set.num = []selectNumInterval{{lo: math.Nextafter(v, math.Inf(1)), hi: math.Inf(1), line: line}}
		case ">=":
			set.num = []selectNumInterval{{lo: v, hi: math.Inf(1), line: line}}
		case "<>":
			set.num = []selectNumInterval{
				{lo: math.Inf(-1), hi: math.Nextafter(v, math.Inf(-1)), line: line},
				{lo: math.Nextafter(v, math.Inf(1)), hi: math.Inf(1), line: line},
			}
		default:
			set.opaque = true
		}
		return set
	}
	if compare != "binary" || !selectStrSupported(value.str) {
		set.opaque = true
		return set
	}
	s := value.str
	switch operator {
	case "=":
		set.str = []selectStrInterval{{lo: s, hi: s, hasLo: true, hasHi: true, line: line}}
	case "<":
		set.str = []selectStrInterval{{hasHi: true, hi: s, hiExcl: true, line: line}}
	case "<=":
		set.str = []selectStrInterval{{hasHi: true, hi: s, line: line}}
	case ">":
		set.str = []selectStrInterval{{hasLo: true, lo: s, loExcl: true, line: line}}
	case ">=":
		set.str = []selectStrInterval{{hasLo: true, lo: s, line: line}}
	case "<>":
		set.str = []selectStrInterval{
			{hasHi: true, hi: s, hiExcl: true, line: line},
			{hasLo: true, lo: s, loExcl: true, line: line},
		}
	default:
		set.opaque = true
	}
	return set
}

// selectStrPointInterval models a string equality item. Option Compare Text is
// approximated by folding ASCII-only literals to lower case; Database and
// non-ASCII/non-BMP inputs stay opaque because their ordering cannot be
// reproduced deterministically.
func selectStrPointInterval(s string, compare string, line int) (selectStrInterval, bool) {
	if !selectStrSupported(s) {
		return selectStrInterval{}, false
	}
	switch compare {
	case "binary":
		return selectStrInterval{lo: s, hi: s, hasLo: true, hasHi: true, line: line}, true
	case "text":
		if !selectStrFoldable(s) {
			return selectStrInterval{}, false
		}
		folded := strings.ToLower(s)
		return selectStrInterval{lo: folded, hi: folded, hasLo: true, hasHi: true, line: line}, true
	default:
		return selectStrInterval{}, false
	}
}

// selectStrSupported reports whether every rune is representable in one UTF-16
// code unit, where Go's byte-wise string order equals VBA's Binary order.
func selectStrSupported(s string) bool {
	for _, r := range s {
		if r > 0xFFFF || r == utf8.RuneError {
			return false
		}
	}
	return true
}

// selectStrFoldable reports whether ASCII case folding reproduces the VBA
// Option Compare Text outcome for this literal.
func selectStrFoldable(s string) bool {
	for i := 0; i < len(s); i++ {
		if s[i] >= 0x80 {
			return false
		}
	}
	return true
}

// filter restricts an item set to the values the selector can actually take.
// A cross-type item is made opaque rather than empty because VBA coerces types
// at runtime (for example `Case "5"` can match a numeric selector).
func (domain selectCaseDomain) filter(set selectItemSet) selectItemSet {
	if set.opaque || set.empty {
		return set
	}
	if len(set.num) > 0 && !domain.numeric {
		return selectItemSet{line: set.line, opaque: true}
	}
	if len(set.str) > 0 && !domain.str {
		return selectItemSet{line: set.line, opaque: true}
	}
	eff := selectItemSet{line: set.line}
	for _, iv := range set.num {
		for _, effIv := range selectNumIntersectUniverse(iv, domain.numUniverse) {
			if domain.integral {
				effIv.lo = math.Ceil(effIv.lo)
				effIv.hi = math.Floor(effIv.hi)
				if effIv.lo > effIv.hi {
					continue
				}
			}
			eff.num = append(eff.num, effIv)
		}
	}
	for _, iv := range set.str {
		eff.str = append(eff.str, selectStrIntersectUniverse(iv, domain.strUniverse)...)
	}
	return eff
}

func selectNumIntersectUniverse(iv selectNumInterval, universe []selectNumInterval) []selectNumInterval {
	if universe == nil {
		return []selectNumInterval{iv}
	}
	out := make([]selectNumInterval, 0, len(universe))
	for _, u := range universe {
		lo, hi := max(iv.lo, u.lo), min(iv.hi, u.hi)
		if lo <= hi {
			out = append(out, selectNumInterval{lo: lo, hi: hi, line: iv.line})
		}
	}
	return out
}

func selectStrIntersectUniverse(iv selectStrInterval, universe []selectStrInterval) []selectStrInterval {
	if universe == nil {
		return []selectStrInterval{iv}
	}
	out := make([]selectStrInterval, 0, len(universe))
	for _, u := range universe {
		if merged, ok := selectStrIntersect(iv, u); ok {
			out = append(out, merged)
		}
	}
	return out
}

// selectStrLoMax returns the more restrictive of two lower bounds.
func selectStrLoMax(a, b selectStrInterval) (string, bool, bool) {
	if !a.hasLo {
		return b.lo, b.hasLo, b.loExcl
	}
	if !b.hasLo {
		return a.lo, a.hasLo, a.loExcl
	}
	switch {
	case a.lo > b.lo:
		return a.lo, true, a.loExcl
	case b.lo > a.lo:
		return b.lo, true, b.loExcl
	default:
		return a.lo, true, a.loExcl || b.loExcl
	}
}

// selectStrHiMin returns the more restrictive of two upper bounds.
func selectStrHiMin(a, b selectStrInterval) (string, bool, bool) {
	if !a.hasHi {
		return b.hi, b.hasHi, b.hiExcl
	}
	if !b.hasHi {
		return a.hi, a.hasHi, a.hiExcl
	}
	switch {
	case a.hi < b.hi:
		return a.hi, true, a.hiExcl
	case b.hi < a.hi:
		return b.hi, true, b.hiExcl
	default:
		return a.hi, true, a.hiExcl || b.hiExcl
	}
}

func selectStrIntersect(a, b selectStrInterval) (selectStrInterval, bool) {
	lo, hasLo, loExcl := selectStrLoMax(a, b)
	hi, hasHi, hiExcl := selectStrHiMin(a, b)
	out := selectStrInterval{lo: lo, hasLo: hasLo, loExcl: loExcl, hi: hi, hasHi: hasHi, hiExcl: hiExcl, line: a.line}
	if hasLo && hasHi {
		if lo > hi || (lo == hi && (loExcl || hiExcl)) {
			return selectStrInterval{}, false
		}
	}
	return out, true
}

// selectEarlierDuplicate reports the first earlier item whose effective value
// set is identical, so the finding can distinguish an exact restatement from a
// merely covered one.
func selectEarlierDuplicate(earlier []selectItemSet, eff selectItemSet) (int, bool) {
	for _, set := range earlier {
		if selectItemSetEqual(set, eff) {
			return set.line, true
		}
	}
	return 0, false
}

func selectItemSetEqual(a, b selectItemSet) bool {
	if len(a.num) != len(b.num) || len(a.str) != len(b.str) {
		return false
	}
	for i := range a.num {
		if a.num[i].lo != b.num[i].lo || a.num[i].hi != b.num[i].hi {
			return false
		}
	}
	for i := range a.str {
		x, y := a.str[i], b.str[i]
		if x.lo != y.lo || x.hi != y.hi || x.hasLo != y.hasLo || x.hasHi != y.hasHi || x.loExcl != y.loExcl || x.hiExcl != y.hiExcl {
			return false
		}
	}
	return true
}

// contains reports whether every interval of eff is already inside a single
// union member; because the union is disjoint, spanning two members implies a
// real gap. The returned line is the earliest union member contributing to the
// coverage proof.
func (coverage selectCoverage) contains(eff selectItemSet) (int, bool) {
	witness := 0
	for _, iv := range eff.num {
		line, ok := selectNumContainingLine(coverage.num, iv)
		if !ok {
			return 0, false
		}
		if witness == 0 || line < witness {
			witness = line
		}
	}
	for _, iv := range eff.str {
		line, ok := selectStrContainingLine(coverage.str, iv)
		if !ok {
			return 0, false
		}
		if witness == 0 || line < witness {
			witness = line
		}
	}
	return witness, true
}

func (coverage *selectCoverage) add(eff selectItemSet) {
	for _, iv := range eff.num {
		coverage.num = selectNumUnionInsert(coverage.num, iv, coverage.integral)
	}
	for _, iv := range eff.str {
		coverage.str = selectStrUnionInsert(coverage.str, iv)
	}
}

func (coverage selectCoverage) coversUniverse(domain selectCaseDomain) bool {
	if len(domain.numUniverse) == 0 && len(domain.strUniverse) == 0 {
		return false
	}
	for _, iv := range domain.numUniverse {
		if _, ok := selectNumContainingLine(coverage.num, iv); !ok {
			return false
		}
	}
	for _, iv := range domain.strUniverse {
		if _, ok := selectStrContainingLine(coverage.str, iv); !ok {
			return false
		}
	}
	return true
}

func selectNumContainingLine(union []selectNumInterval, iv selectNumInterval) (int, bool) {
	for _, u := range union {
		if u.lo > iv.lo {
			break
		}
		if iv.hi <= u.hi {
			return u.line, true
		}
	}
	return 0, false
}

// selectNumUnionInsert merges iv into the sorted disjoint union. An existing
// member merges when it overlaps or is adjacent to iv — `u.hi` one float64 ulp
// below `iv.lo` still leaves no representable gap, so the closed-interval union
// remains contiguous. On integral domains adjacency is one whole step instead:
// filtered bounds are exact integers, so [0,127] and [128,255] merge while a
// real gap like [0,127] + [129,255] stays disjoint.
func selectNumUnionInsert(union []selectNumInterval, iv selectNumInterval, integral bool) []selectNumInterval {
	loAdj, hiAdj := math.Nextafter(iv.lo, math.Inf(-1)), math.Nextafter(iv.hi, math.Inf(1))
	if integral && iv.lo > -(1<<53) && iv.hi < 1<<53 {
		// Whole-step adjacency is only exact while every bound stays inside
		// float64's exact integer range; beyond ±2^53 (reachable only through
		// uncapped Double literals on unbounded integral domains such as
		// LongLong) the ±1 step is not representable, so ulp adjacency applies.
		loAdj, hiAdj = iv.lo-1, iv.hi+1
	}
	out := make([]selectNumInterval, 0, len(union)+1)
	i := 0
	for i < len(union) && union[i].hi < loAdj {
		out = append(out, union[i])
		i++
	}
	for i < len(union) && union[i].lo <= hiAdj {
		iv.lo = min(iv.lo, union[i].lo)
		iv.hi = max(iv.hi, union[i].hi)
		iv.line = min(iv.line, union[i].line)
		i++
	}
	out = append(out, iv)
	return append(out, union[i:]...)
}

// selectStrBefore reports whether a ends strictly before b starts, leaving at
// least the shared boundary value uncovered when both sides exclude it.
func selectStrBefore(a, b selectStrInterval) bool {
	if !a.hasHi || !b.hasLo {
		return false
	}
	if a.hi < b.lo {
		return true
	}
	return a.hi == b.lo && a.hiExcl && b.loExcl
}

// selectStrMerge keeps the more permissive bound on each side; an unbounded
// end or an inclusive shared bound wins.
func selectStrMerge(a, b selectStrInterval) selectStrInterval {
	merged := a
	if !b.hasLo || (a.hasLo && (b.lo < a.lo || (b.lo == a.lo && !b.loExcl))) {
		merged.lo, merged.hasLo, merged.loExcl = b.lo, b.hasLo, b.loExcl
	}
	if !b.hasHi || (a.hasHi && (b.hi > a.hi || (b.hi == a.hi && !b.hiExcl))) {
		merged.hi, merged.hasHi, merged.hiExcl = b.hi, b.hasHi, b.hiExcl
	}
	merged.line = min(a.line, b.line)
	return merged
}

func selectStrUnionInsert(union []selectStrInterval, iv selectStrInterval) []selectStrInterval {
	out := make([]selectStrInterval, 0, len(union)+1)
	i := 0
	for i < len(union) && selectStrBefore(union[i], iv) {
		out = append(out, union[i])
		i++
	}
	for i < len(union) && !selectStrBefore(iv, union[i]) {
		iv = selectStrMerge(union[i], iv)
		i++
	}
	out = append(out, iv)
	return append(out, union[i:]...)
}

func selectStrContainingLine(union []selectStrInterval, iv selectStrInterval) (int, bool) {
	for _, u := range union {
		if selectStrContains(u, iv) {
			return u.line, true
		}
	}
	return 0, false
}

func selectStrContains(u, iv selectStrInterval) bool {
	if u.hasLo {
		if !iv.hasLo {
			return false
		}
		if u.lo > iv.lo || (u.lo == iv.lo && u.loExcl && !iv.loExcl) {
			return false
		}
	}
	if u.hasHi {
		if !iv.hasHi {
			return false
		}
		if u.hi < iv.hi || (u.hi == iv.hi && u.hiExcl && !iv.hiExcl) {
			return false
		}
	}
	return true
}

// selectCaseSelectorDomain resolves the set of values the selector can take: a
// statically known constant gives a singleton domain, a declared identifier
// gives a type domain, and anything else stays open so coverage claims are
// never made for an unconstrained selector.
func selectCaseSelectorDomain(selector procedureir.Expression, proc sourceProcedure, env constexpr.Environment, compare string, moduleDecls map[string]sourceDeclaration) selectCaseDomain {
	if result := constexpr.Evaluate(selector.Text, env); result.Kind == constexpr.Known {
		return selectCaseConstantDomain(result.Typed, compare)
	}
	name, suffix, ok := selectCaseSelectorIdentifier(selector.Text)
	if !ok {
		return selectCaseOpenDomain()
	}
	typeName := selectCaseDeclaredType(name, suffix, proc, moduleDecls)
	if typeName == "" {
		typeName = selectCaseSuffixType(suffix)
	}
	return selectCaseTypeDomain(typeName)
}

func selectCaseConstantDomain(value constexpr.Value, compare string) selectCaseDomain {
	switch value.Kind {
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong:
		if value.Integer > 1<<53 || value.Integer < -(1<<53) {
			return selectCaseDomain{opaque: true}
		}
		return selectCaseSingletonNum(float64(value.Integer), true, selectCaseKindTypeName(value.Kind))
	case constexpr.ValueSingle, constexpr.ValueDouble:
		if math.IsNaN(value.Float) || math.IsInf(value.Float, 0) {
			return selectCaseDomain{opaque: true}
		}
		return selectCaseSingletonNum(value.Float, false, selectCaseKindTypeName(value.Kind))
	case constexpr.ValueCurrency:
		if value.Currency > 1<<53 || value.Currency < -(1<<53) {
			return selectCaseDomain{opaque: true}
		}
		return selectCaseSingletonNum(float64(value.Currency)/10000, false, "Currency")
	case constexpr.ValueBoolean:
		v := 0.0
		if value.Boolean {
			v = -1
		}
		return selectCaseSingletonNum(v, true, "Boolean")
	case constexpr.ValueString:
		s := value.String
		if !selectStrSupported(s) || compare == "database" {
			// Option Compare Database ordering depends on the Access locale and
			// cannot be reproduced deterministically, so the singleton universe
			// would risk a wrong impossible/else_covered claim.
			return selectCaseDomain{opaque: true}
		}
		if compare == "text" {
			if !selectStrFoldable(s) {
				// Non-ASCII text still case-folds under Option Compare Text
				// with mappings this model cannot reproduce (U+212A compares
				// equal to "k"), so a raw singleton universe would disagree
				// with the folded item intervals.
				return selectCaseDomain{opaque: true}
			}
			s = strings.ToLower(s)
		}
		return selectCaseDomain{
			str:           true,
			elseCoverable: true,
			strUniverse:   []selectStrInterval{{lo: s, hi: s, hasLo: true, hasHi: true}},
			typeName:      "String",
		}
	default:
		return selectCaseDomain{opaque: true}
	}
}

func selectCaseKindTypeName(kind constexpr.ValueKind) string {
	switch kind {
	case constexpr.ValueInteger:
		return "Integer"
	case constexpr.ValueLong:
		return "Long"
	case constexpr.ValueLongLong:
		return "LongLong"
	case constexpr.ValueSingle:
		return "Single"
	case constexpr.ValueDouble:
		return "Double"
	default:
		return "Variant"
	}
}

func selectCaseSingletonNum(v float64, integral bool, typeName string) selectCaseDomain {
	return selectCaseDomain{
		numeric:       true,
		integral:      integral,
		elseCoverable: true,
		numUniverse:   []selectNumInterval{{lo: v, hi: v}},
		typeName:      typeName,
	}
}

func selectCaseOpenDomain() selectCaseDomain {
	return selectCaseDomain{numeric: true, str: true, typeName: "Variant"}
}

// selectCaseSelectorIdentifier unwraps parentheses and brackets and reports the
// bare identifier plus any trailing VBA type suffix. Member access, calls, and
// other composite expressions are not identifiers.
func selectCaseSelectorIdentifier(text string) (name string, suffix byte, ok bool) {
	trimmed := strings.TrimSpace(text)
	for strings.HasPrefix(trimmed, "(") && strings.HasSuffix(trimmed, ")") {
		trimmed = strings.TrimSpace(trimmed[1 : len(trimmed)-1])
	}
	trimmed = strings.Trim(trimmed, "[]")
	if trimmed == "" {
		return "", 0, false
	}
	if last := trimmed[len(trimmed)-1]; strings.ContainsRune("$%&#@^!", rune(last)) {
		suffix = last
		trimmed = trimmed[:len(trimmed)-1]
	}
	if !selectCaseIdentRe.MatchString(trimmed) {
		return "", 0, false
	}
	return trimmed, suffix, true
}

func selectCaseDeclaredType(name string, suffix byte, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) string {
	for declaration := range proc.Declarations.All() {
		if strings.EqualFold(cleanIdentifier(declaration.Name), name) {
			return declaration.Type
		}
	}
	for parameter := range proc.Params.All() {
		if strings.EqualFold(cleanIdentifier(parameter.Name), name) {
			return parameter.Type
		}
	}
	lowered := strings.ToLower(name)
	if declaration, ok := moduleDecls[lowered]; ok {
		return declaration.Type
	}
	if suffix != 0 {
		if declaration, ok := moduleDecls[lowered+string(suffix)]; ok {
			return declaration.Type
		}
	}
	return ""
}

func selectCaseSuffixType(suffix byte) string {
	switch suffix {
	case '%':
		return "Integer"
	case '&':
		return "Long"
	case '^':
		return "LongLong"
	case '!':
		return "Single"
	case '#':
		return "Double"
	case '@':
		return "Currency"
	case '$':
		return "String"
	default:
		return ""
	}
}

// selectCaseTypeDomain maps a declared VBA type onto the selector domain.
// Unbounded types keep a nil universe so Case Else can never be proven covered;
// unknown and user types fail open.
func selectCaseTypeDomain(typeName string) selectCaseDomain {
	lowered := strings.ToLower(strings.TrimSpace(cleanIdentifier(typeName)))
	switch {
	case lowered == "byte":
		return selectCaseDomain{numeric: true, integral: true, elseCoverable: true, typeName: "Byte",
			numUniverse: []selectNumInterval{{lo: 0, hi: 255}}}
	case lowered == "integer":
		return selectCaseDomain{numeric: true, integral: true, elseCoverable: true, typeName: "Integer",
			numUniverse: []selectNumInterval{{lo: -32768, hi: 32767}}}
	case lowered == "long":
		return selectCaseDomain{numeric: true, integral: true, elseCoverable: true, typeName: "Long",
			numUniverse: []selectNumInterval{{lo: -2147483648, hi: 2147483647}}}
	case lowered == "longlong" || lowered == "longptr":
		return selectCaseDomain{numeric: true, integral: true, typeName: "LongLong"}
	case lowered == "single" || lowered == "double" || lowered == "currency" || lowered == "decimal":
		return selectCaseDomain{numeric: true, typeName: selectCaseDomainTypeTitle(lowered)}
	case lowered == "boolean":
		return selectCaseDomain{numeric: true, integral: true, elseCoverable: true, typeName: "Boolean",
			numUniverse: []selectNumInterval{{lo: -1, hi: -1}, {lo: 0, hi: 0}}}
	case lowered == "string" || strings.HasPrefix(lowered, "string *"):
		return selectCaseDomain{str: true, typeName: "String"}
	case lowered == "date":
		return selectCaseDomain{opaque: true}
	case lowered == "" || lowered == "variant":
		return selectCaseOpenDomain()
	case isObjectType(typeName):
		return selectCaseDomain{opaque: true}
	default:
		return selectCaseOpenDomain()
	}
}

func selectCaseDomainTypeTitle(lowered string) string {
	return strings.ToUpper(lowered[:1]) + lowered[1:]
}

func (a Analyzer) selectCaseItemFinding(file parsedFile, proc sourceProcedure, item procedureir.Statement, kind string, coveredByLine int, domain selectCaseDomain) Finding {
	itemText := strings.TrimSpace(item.Text)
	message := ""
	reason := "Select Case executes only the first matching clause in source order, so this item can never be reached."
	suggestion := ""
	switch kind {
	case "duplicate":
		message = "Case item \"" + itemText + "\" duplicates the Case item on line " + strconvItoa(coveredByLine) + "."
		suggestion = "Remove this item or merge it with the earlier item."
	case "covered":
		message = "Case item \"" + itemText + "\" is fully covered by earlier Case items (first covered on line " + strconvItoa(coveredByLine) + ")."
		suggestion = "Narrow the item to values not already handled, move it earlier, or remove it."
	case "empty_range":
		message = "Case item \"" + itemText + "\" describes an empty range."
		reason = "The lower bound of a To range exceeds the upper bound, so no value can match this item."
		suggestion = "Correct the range bounds or remove the item."
	case "impossible":
		message = "Case item \"" + itemText + "\" can never match a " + domain.typeName + " selector value."
		reason = "The selector's declared type cannot produce a value that satisfies this item."
		suggestion = "Adjust the item to the selector's domain or remove it."
	}
	finding := a.simpleFinding(file, proc, item.Range.StartLine, "VBA259", "warning", message, reason, suggestion)
	finding.Column = selectCaseUTF16Column(file.Lines, item.Range.StartLine, item.Range.StartColumn)
	finding.EndLine = item.Range.EndLine
	finding.EndColumn = selectCaseUTF16Column(file.Lines, item.Range.EndLine, item.Range.EndColumn)
	finding.SelectCaseUnreachable = &SelectCaseUnreachableContext{Kind: kind, Item: itemText, CoveredByLine: coveredByLine}
	return finding
}

func (a Analyzer) selectCaseElseFinding(file parsedFile, proc sourceProcedure, clause procedureir.Statement, domain selectCaseDomain) Finding {
	finding := a.simpleFinding(file, proc, clause.Range.StartLine, "VBA259", "warning",
		"Case Else is unreachable because earlier Case items already cover every "+domain.typeName+" value.",
		"Select Case runs Case Else only when no earlier Case item matches; every possible selector value is already matched.",
		"Remove the unreachable Case Else branch, or keep it deliberately as a defensive fallback.")
	finding.Column = selectCaseUTF16Column(file.Lines, clause.Range.StartLine, clause.Range.StartColumn)
	finding.EndLine = clause.Range.StartLine
	endColumn := clause.Range.StartColumn + len("Case Else")
	if match := selectCaseElseWordRe.FindString(clause.Text); match != "" {
		endColumn = clause.Range.StartColumn + len(match)
	}
	finding.EndColumn = selectCaseUTF16Column(file.Lines, clause.Range.StartLine, endColumn)
	finding.SelectCaseUnreachable = &SelectCaseUnreachableContext{Kind: "else_covered"}
	return finding
}

// selectCaseUTF16Column converts a 1-based vbaast byte column into the 1-based
// UTF-16 code-unit column that analyzer findings carry when EndLine is set.
// Multi-byte splits snap back to the containing rune boundary.
func selectCaseUTF16Column(lines []string, line, oneBasedByteColumn int) int {
	if line < 1 || line > len(lines) || oneBasedByteColumn < 1 {
		return 0
	}
	text := lines[line-1]
	col := oneBasedByteColumn - 1
	if col > len(text) {
		col = len(text)
	}
	for col > 0 && col < len(text) && !utf8.RuneStart(text[col]) {
		col--
	}
	return len(utf16.Encode([]rune(text[:col]))) + 1
}

// optionCompare mirrors optionBase: it scans the module header for the
// `Option Compare` mode and defaults to Binary, matching VBA's default.
func optionCompare(lines []string) string {
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(normalizedCodeLine(line)))
		if len(fields) == 3 && fields[0] == "option" && fields[1] == "compare" {
			switch fields[2] {
			case "binary", "text", "database":
				return fields[2]
			}
		}
	}
	return "binary"
}

func fileOptionCompare(file parsedFile) string {
	if file.OptionCompareSet {
		return file.OptionCompare
	}
	return optionCompare(file.Lines)
}
