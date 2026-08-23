package analyze

import (
	"iter"
	"sort"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// readOnlySpan is an internal, allocation-free view over immutable IR
// storage. It intentionally has no conversion to a mutable slice API; the
// named slice representation keeps existing in-package range/len callers
// source-compatible while new code can use Len, At, or All.
type readOnlySpan[T any] []T

func (span readOnlySpan[T]) Len() int { return len(span) }

func (span readOnlySpan[T]) At(index int) (T, bool) {
	if index < 0 || index >= len(span) {
		var zero T
		return zero, false
	}
	return span[index], true
}

func (span readOnlySpan[T]) All() iter.Seq[T] {
	return func(yield func(T) bool) {
		for _, value := range span {
			if !yield(value) {
				return
			}
		}
	}
}

// procedureAnalysisFacts is the immutable, procedure-local index shared by
// analysis rules. Production facts retain only the canonical ProcedureIR
// pointer and compact offsets; they never own a second declaration,
// statement, expression, call, or access collection. The legacy slices below
// exist only for small synthetic fixtures that do not have a ProcedureIR
// owner.
//
// The fields are deliberately private. Consumers should use the lookup and
// iteration helpers below instead of retaining or modifying the indexes.
// sourceProcedure itself is an immutable projection during an analysis run;
// procedure workers may safely read the same facts concurrently.
type procedureAnalysisFacts struct {
	features procedureFeatureSet
	// procedure is the canonical owned IR for production views. The legacy
	// slices below are populated only by focused synthetic tests and are kept
	// as a compatibility path while consumers migrate to the view API.
	procedure    *procedureir.ProcedureIR
	declarations []procedureir.Declaration
	statements   []procedureir.Statement
	expressions  []procedureir.Expression
	calls        []procedureir.CallSite
	accesses     []procedureir.VariableAccess

	declarationIndex map[int]int
	statementIndex   map[int]int
	expressionIndex  map[int]int
	callsByStatement
	accessesByStatement
	memberExpressionsByStatement map[int][]int
	// memberExpressionFallback is true only for hand-built/recovered IR where
	// at least one expression omits StatementID. Production IR assigns the ID
	// consistently, so MemberExpressionsForStatement can return its compact
	// grouped index without traversing expression trees in that common case.
	memberExpressionFallback bool
}

// newProcedureAnalysisFacts constructs facts from already-owned procedure IR
// projections for focused synthetic fixtures. It intentionally does not copy
// the supplied slices; production callers use newProcedureAnalysisFactsForProcedure
// so the canonical ProcedureIR is the explicit owner.
func newProcedureAnalysisFacts(
	statements []procedureir.Statement,
	expressions []procedureir.Expression,
	calls []procedureir.CallSite,
	accesses []procedureir.VariableAccess,
) *procedureAnalysisFacts {
	return newProcedureAnalysisFactsWithDeclarations(nil, statements, expressions, calls, accesses)
}

// newProcedureAnalysisFactsForProcedure builds compact indexes over the
// canonical ProcedureIR. No procedure collection is copied or retained as a
// second backing allocation.
func newProcedureAnalysisFactsForProcedure(procedure *procedureir.ProcedureIR) *procedureAnalysisFacts {
	if procedure == nil {
		return newProcedureAnalysisFacts(nil, nil, nil, nil)
	}
	facts := &procedureAnalysisFacts{procedure: procedure}
	return facts.initialize(procedure.Declarations, procedure.Statements, procedure.Expressions, procedure.Calls, procedure.Accesses)
}

func newProcedureAnalysisFactsWithDeclarations(
	declarations []procedureir.Declaration,
	statements []procedureir.Statement,
	expressions []procedureir.Expression,
	calls []procedureir.CallSite,
	accesses []procedureir.VariableAccess,
) *procedureAnalysisFacts {
	facts := &procedureAnalysisFacts{
		declarations: declarations,
		statements:   statements,
		expressions:  expressions,
		calls:        calls,
		accesses:     accesses,
	}
	return facts.initialize(declarations, statements, expressions, calls, accesses)
}

func (facts *procedureAnalysisFacts) initialize(
	declarations []procedureir.Declaration,
	statements []procedureir.Statement,
	expressions []procedureir.Expression,
	calls []procedureir.CallSite,
	accesses []procedureir.VariableAccess,
) *procedureAnalysisFacts {

	if len(declarations) > 0 {
		facts.declarationIndex = make(map[int]int, len(declarations))
		for index, declaration := range declarations {
			facts.features.observeDeclaration(declaration)
			facts.declarationIndex[declaration.ID] = index
		}
	}
	if len(statements) > 0 {
		facts.statementIndex = make(map[int]int, len(statements))
		for index, statement := range statements {
			facts.features.observeStatement(statement)
			facts.statementIndex[statement.ID] = index
		}
	}
	if len(expressions) > 0 {
		facts.expressionIndex = make(map[int]int, len(expressions))
		for index, expression := range expressions {
			facts.features.observeExpression(expression)
			facts.expressionIndex[expression.ID] = index
		}
		// The IR builder assigns StatementID to every expression it emits. Keep
		// this index compact by grouping those IDs directly; the compatibility
		// fallback in MemberExpressionsForStatement handles hand-built IR where
		// a reachable child has no StatementID.
		for index, expression := range expressions {
			if expression.StatementID == 0 {
				facts.memberExpressionFallback = true
			}
			if expression.StatementID == 0 || expression.Recovered || expression.Kind != procedureir.ExpressionMember {
				continue
			}
			facts.addMemberExpression(expression.StatementID, index)
		}
		for statementID, indexes := range facts.memberExpressionsByStatement {
			sort.SliceStable(indexes, func(i, j int) bool {
				return expressions[indexes[i]].Range.StartByte < expressions[indexes[j]].Range.StartByte
			})
			facts.memberExpressionsByStatement[statementID] = indexes
		}
	}
	if len(calls) > 0 {
		for _, call := range calls {
			facts.features.observeCall(call)
		}
		facts.callsByStatement = newCallFactsByStatement(calls)
	}
	if len(accesses) > 0 {
		facts.accessesByStatement = newAccessFactsByStatement(accesses)
	}
	return facts
}

func (facts *procedureAnalysisFacts) declarationsView() readOnlySpan[procedureir.Declaration] {
	if facts == nil {
		return nil
	}
	if facts.procedure != nil {
		return readOnlySpan[procedureir.Declaration](facts.procedure.Declarations)
	}
	return readOnlySpan[procedureir.Declaration](facts.declarations)
}

func (facts *procedureAnalysisFacts) statementsView() readOnlySpan[procedureir.Statement] {
	if facts == nil {
		return nil
	}
	if facts.procedure != nil {
		return readOnlySpan[procedureir.Statement](facts.procedure.Statements)
	}
	return readOnlySpan[procedureir.Statement](facts.statements)
}

func (facts *procedureAnalysisFacts) expressionsView() readOnlySpan[procedureir.Expression] {
	if facts == nil {
		return nil
	}
	if facts.procedure != nil {
		return readOnlySpan[procedureir.Expression](facts.procedure.Expressions)
	}
	return readOnlySpan[procedureir.Expression](facts.expressions)
}

func (facts *procedureAnalysisFacts) callsView() readOnlySpan[procedureir.CallSite] {
	if facts == nil {
		return nil
	}
	if facts.procedure != nil {
		return readOnlySpan[procedureir.CallSite](facts.procedure.Calls)
	}
	return readOnlySpan[procedureir.CallSite](facts.calls)
}

func (facts *procedureAnalysisFacts) accessesView() readOnlySpan[procedureir.VariableAccess] {
	if facts == nil {
		return nil
	}
	if facts.procedure != nil {
		return readOnlySpan[procedureir.VariableAccess](facts.procedure.Accesses)
	}
	return readOnlySpan[procedureir.VariableAccess](facts.accesses)
}

// Declarations returns a read-only view of procedure-local declarations in IR
// order. The view aliases the canonical ProcedureIR and does not allocate.
func (facts *procedureAnalysisFacts) Declarations() readOnlySpan[procedureir.Declaration] {
	if facts == nil {
		return nil
	}
	return facts.declarationsView()
}

// Declaration returns one procedure-local declaration by its stable IR ID.
func (facts *procedureAnalysisFacts) Declaration(id int) (procedureir.Declaration, bool) {
	if facts == nil {
		return procedureir.Declaration{}, false
	}
	index, ok := facts.declarationIndex[id]
	if !ok {
		return procedureir.Declaration{}, false
	}
	declarations := facts.declarationsView()
	return declarations[index], true
}

// Statements returns a read-only view of the procedure's source-order
// statements without allocating.
func (facts *procedureAnalysisFacts) Statements() readOnlySpan[procedureir.Statement] {
	if facts == nil {
		return nil
	}
	return facts.statementsView()
}

// Statement returns one statement by its stable IR ID.
func (facts *procedureAnalysisFacts) Statement(id int) (procedureir.Statement, bool) {
	if facts == nil {
		return procedureir.Statement{}, false
	}
	index, ok := facts.statementIndex[id]
	if !ok {
		return procedureir.Statement{}, false
	}
	statements := facts.statementsView()
	return statements[index], true
}

// Expressions returns a read-only view of the procedure's expressions in IR
// order without allocating.
func (facts *procedureAnalysisFacts) Expressions() readOnlySpan[procedureir.Expression] {
	if facts == nil {
		return nil
	}
	return facts.expressionsView()
}

// forEachExpression visits expressions without exposing the facts-owned slice
// or allocating a copy for read-only rule scans.
func (facts *procedureAnalysisFacts) forEachExpression(visit func(procedureir.Expression)) {
	if facts == nil || visit == nil {
		return
	}
	for _, expression := range facts.expressionsView() {
		visit(expression)
	}
}

// Expression returns one expression by its stable IR ID.
func (facts *procedureAnalysisFacts) Expression(id int) (procedureir.Expression, bool) {
	if facts == nil {
		return procedureir.Expression{}, false
	}
	index, ok := facts.expressionIndex[id]
	if !ok {
		return procedureir.Expression{}, false
	}
	expressions := facts.expressionsView()
	return expressions[index], true
}

// Calls returns a read-only view of the procedure's calls in IR order without
// allocating.
func (facts *procedureAnalysisFacts) Calls() readOnlySpan[procedureir.CallSite] {
	if facts == nil {
		return nil
	}
	return facts.callsView()
}

// CallsForStatement returns calls associated with a statement, preserving IR
// order. The returned slice is independent of the immutable facts.
func (facts *procedureAnalysisFacts) CallsForStatement(statementID int) []procedureir.CallSite {
	if facts == nil {
		return nil
	}
	return facts.callsByStatement.values(facts.callsView(), statementID)
}

// forEachCallForStatement visits calls without exposing the facts-owned slice
// or allocating a defensive result. Hot flow analysis uses this form; callers
// that need to retain values should use CallsForStatement instead.
func (facts *procedureAnalysisFacts) forEachCallForStatement(statementID int, visit func(procedureir.CallSite)) {
	if facts == nil || visit == nil {
		return
	}
	if span, contiguous := facts.callsByStatement.groups.contiguousSpan(statementID); contiguous {
		for _, call := range facts.callsView()[span.start:span.end] {
			visit(call)
		}
		return
	}
	indexes, ok := facts.callsByStatement.groups.lookup(statementID)
	if !ok {
		return
	}
	for _, index := range indexes {
		calls := facts.callsView()
		if index >= 0 && index < len(calls) {
			visit(calls[index])
		}
	}
}

// Accesses returns a read-only view of the procedure's variable accesses in IR
// order without allocating.
func (facts *procedureAnalysisFacts) Accesses() readOnlySpan[procedureir.VariableAccess] {
	if facts == nil {
		return nil
	}
	return facts.accessesView()
}

// AccessesForStatement returns accesses associated with a statement,
// preserving IR order. The returned slice is independent of the immutable
// facts.
func (facts *procedureAnalysisFacts) AccessesForStatement(statementID int) []procedureir.VariableAccess {
	if facts == nil {
		return nil
	}
	return facts.accessesByStatement.values(facts.accessesView(), statementID)
}

// forEachAccessForStatement is the allocation-free counterpart to
// AccessesForStatement for hot flow transfers.
func (facts *procedureAnalysisFacts) forEachAccessForStatement(statementID int, visit func(procedureir.VariableAccess)) {
	if facts == nil || visit == nil {
		return
	}
	if span, contiguous := facts.accessesByStatement.groups.contiguousSpan(statementID); contiguous {
		for _, access := range facts.accessesView()[span.start:span.end] {
			visit(access)
		}
		return
	}
	indexes, ok := facts.accessesByStatement.groups.lookup(statementID)
	if !ok {
		return
	}
	for _, index := range indexes {
		accesses := facts.accessesView()
		if index >= 0 && index < len(accesses) {
			visit(accesses[index])
		}
	}
}

// MemberExpressionsForStatement returns member expressions belonging to a
// statement, preserving the expression order emitted by the IR builder.
func (facts *procedureAnalysisFacts) MemberExpressionsForStatement(statementID int) []procedureir.Expression {
	if facts == nil {
		return nil
	}
	expressions := facts.expressionsView()
	indexes := facts.memberExpressionsByStatement[statementID]
	if !facts.memberExpressionFallback {
		if len(indexes) == 0 {
			// In authoritative production IR, an empty group proves that this
			// statement has no eligible member expressions. Do not fall back to
			// tree traversal or a procedure-wide scan.
			return nil
		}
		out := make([]procedureir.Expression, 0, len(indexes))
		for _, index := range indexes {
			out = append(out, expressions[index])
		}
		return out
	}
	// Focused tests and recovered IR fixtures may omit StatementID on a child
	// expression even though the statement's root reaches it. Preserve the
	// historical traversal fallback for that uncommon shape without paying its
	// per-statement map cost during fact construction.
	statement, ok := facts.Statement(statementID)
	if !ok || len(statement.ExpressionIDs) == 0 {
		if len(indexes) == 0 {
			return nil
		}
		out := make([]procedureir.Expression, 0, len(indexes))
		for _, index := range indexes {
			if index >= 0 && index < len(expressions) {
				out = append(out, expressions[index])
			}
		}
		return out
	}
	seen := map[int]bool{}
	var out []procedureir.Expression
	var visit func(int)
	visit = func(id int) {
		if seen[id] {
			return
		}
		seen[id] = true
		expression, ok := facts.Expression(id)
		if !ok {
			return
		}
		if !expression.Recovered && expression.Kind == procedureir.ExpressionMember {
			out = append(out, expression)
		}
		for _, childID := range expression.Children {
			visit(childID)
		}
	}
	for _, id := range statement.ExpressionIDs {
		visit(id)
	}
	if len(out) == 0 {
		for _, expression := range expressions {
			if expression.StatementID == statement.ID && !expression.Recovered && expression.Kind == procedureir.ExpressionMember {
				out = append(out, expression)
			}
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		return out[i].Range.StartByte < out[j].Range.StartByte
	})
	return out
}

func (facts *procedureAnalysisFacts) addMemberExpression(statementID, expressionIndex int) {
	if facts.memberExpressionsByStatement == nil {
		facts.memberExpressionsByStatement = make(map[int][]int)
	}
	facts.memberExpressionsByStatement[statementID] = append(facts.memberExpressionsByStatement[statementID], expressionIndex)
}

// indexSpan stores a contiguous range in a source-order IR slice. Grouped
// facts use one span per statement instead of allocating one slice per group.
type indexSpan struct {
	start int
	end   int
}

type callsByStatement struct {
	groups indexGroups
}

func newCallFactsByStatement(calls []procedureir.CallSite) callsByStatement {
	ids := make([]int, len(calls))
	for index, call := range calls {
		ids[index] = call.StatementID
	}
	return callsByStatement{groups: newIndexGroups(ids)}
}

func (facts callsByStatement) values(calls []procedureir.CallSite, statementID int) []procedureir.CallSite {
	if span, contiguous := facts.groups.contiguousSpan(statementID); contiguous {
		return append([]procedureir.CallSite(nil), calls[span.start:span.end]...)
	}
	indexes, ok := facts.groups.lookup(statementID)
	if !ok {
		return nil
	}
	out := make([]procedureir.CallSite, 0, len(indexes))
	for _, index := range indexes {
		if index >= 0 && index < len(calls) {
			out = append(out, calls[index])
		}
	}
	return out
}

type accessesByStatement struct {
	groups indexGroups
}

func newAccessFactsByStatement(accesses []procedureir.VariableAccess) accessesByStatement {
	ids := make([]int, len(accesses))
	for index, access := range accesses {
		ids[index] = access.StatementID
	}
	return accessesByStatement{groups: newIndexGroups(ids)}
}

func (facts accessesByStatement) values(accesses []procedureir.VariableAccess, statementID int) []procedureir.VariableAccess {
	if span, contiguous := facts.groups.contiguousSpan(statementID); contiguous {
		return append([]procedureir.VariableAccess(nil), accesses[span.start:span.end]...)
	}
	indexes, ok := facts.groups.lookup(statementID)
	if !ok {
		return nil
	}
	out := make([]procedureir.VariableAccess, 0, len(indexes))
	for _, index := range indexes {
		if index >= 0 && index < len(accesses) {
			out = append(out, accesses[index])
		}
	}
	return out
}

type indexGroups struct {
	// spans stores the common case where all entries for a statement occupy
	// one source-order range. sparse stores only statement IDs whose entries
	// are interleaved with another statement.
	spans  map[int]indexSpan
	sparse map[int][]int
}

func newIndexGroups(ids []int) indexGroups {
	groups := indexGroups{}
	for index, id := range ids {
		if groups.sparse != nil {
			if indexes, fragmented := groups.sparse[id]; fragmented {
				groups.sparse[id] = append(indexes, index)
				continue
			}
		}
		span, ok := groups.spans[id]
		if !ok {
			if groups.spans == nil {
				groups.spans = make(map[int]indexSpan)
			}
			groups.spans[id] = indexSpan{start: index, end: index + 1}
			continue
		}
		if span.end == index {
			span.end = index + 1
			groups.spans[id] = span
			continue
		}
		if groups.sparse == nil {
			groups.sparse = make(map[int][]int)
		}
		indexes := make([]int, 0, span.end-span.start+1)
		for previous := span.start; previous < span.end; previous++ {
			indexes = append(indexes, previous)
		}
		groups.sparse[id] = append(indexes, index)
		delete(groups.spans, id)
	}
	return groups
}

func (groups indexGroups) lookup(id int) ([]int, bool) {
	if indexes, ok := groups.sparse[id]; ok {
		return indexes, true
	}
	span, ok := groups.spans[id]
	if !ok {
		return nil, false
	}
	indexes := make([]int, 0, span.end-span.start)
	for index := span.start; index < span.end; index++ {
		indexes = append(indexes, index)
	}
	return indexes, true
}

func (groups indexGroups) contiguousSpan(id int) (indexSpan, bool) {
	if groups.sparse != nil {
		if _, fragmented := groups.sparse[id]; fragmented {
			return indexSpan{}, false
		}
	}
	span, ok := groups.spans[id]
	return span, ok
}

// analysisFacts returns the attached facts and provides a compatibility path
// for sourceProcedure values assembled directly by focused unit tests.
func (procedure sourceProcedure) analysisFacts() *procedureAnalysisFacts {
	if procedure.Facts != nil {
		return procedure.Facts
	}
	facts := newProcedureAnalysisFactsWithDeclarations(procedure.Declarations, procedure.Statements, procedure.Expressions, procedure.Calls, procedure.Accesses)
	facts.features.addUnknown(allProcedureFeatures)
	return facts
}
