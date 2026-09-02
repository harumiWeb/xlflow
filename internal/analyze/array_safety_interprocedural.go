package analyze

import (
	"regexp"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type arrayByRefAllocationSummaries map[string]map[int]bool

// arrayByRefConditionalAllocations records a ByRef array output that is
// allocated only when a count-bearing input is positive. The outer key is the
// callee procedure; each entry maps the output array parameter index to the
// count-bearing input parameter index.
type arrayByRefConditionalAllocations map[string]map[int]int

// arrayByRefLengthAllocations records a ByRef array output whose paired
// ByRef scalar output is assigned a successful array length. A positive value
// of that scalar is therefore a conditional allocation proof for the array.
type arrayByRefLengthAllocations map[string]map[int]int

var (
	arrayByRefCountExitRe   = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*\.\s*count\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayByRefCountRedimRe  = regexp.MustCompile(`(?i)^\s*redim\s+([A-Za-z_]\w*)\s*\(\s*0\s+to\s+([A-Za-z_]\w*)\s*\.\s*count\s*-\s*1\s*\)\s*$`)
	arrayByRefLengthFullRe  = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayByRefLengthUpperRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
)

type arrayModuleAllocationSummaries map[string]map[string]bool

// arrayModuleInvalidationSummaries records module arrays that may be
// unallocated or unknown when a project-local procedure returns normally.
// The summary starts from an allocated module-array state, so a fixed-size
// array and an Erase followed by a guaranteed ReDim remain allocated while a
// reachable conditional Erase is retained as an invalidation.
type arrayModuleInvalidationSummaries map[string]map[string]bool

type arrayProcedureDominators map[string]map[vbacfg.BlockID]bool

type arrayModuleConfigurationState struct {
	byProcedure       map[string]map[string]bool
	dataTable         map[string]bool
	genericCollection map[string]bool
}

// arrayModuleEntryStates records module-level arrays that are allocated at
// every known entry into a project-local helper. A private helper is analyzed
// independently from its callers, so without this summary an allocation made
// by a public entry procedure is lost as soon as the call crosses a procedure
// boundary.
type arrayModuleEntryStates map[string]map[string]bool

// arrayModuleReadyGuardStates records the stronger, source-owned invariant
// behind a module Boolean readiness guard. The implication is intentionally
// narrow: the guard has one source-owned True write, that write is reached
// only after the module array is allocated on every path, and direct array
// invalidation is paired with a dominating False write. This lets a public
// consumer prove its module array without trusting arbitrary caller state.
type arrayModuleReadyGuardStates map[string]map[string]map[string]bool

type arrayParticipantGraph struct {
	all              map[string]sourceProcedure
	fileByKey        map[string]parsedFile
	byModule         map[string][]string
	keyByIdentity    map[string]string
	candidateIndex   arrayCandidateIndex
	adjacency        map[string]map[string]bool
	reverse          map[string]map[string]bool
	resolvedReverse  map[string]map[string]bool
	callAdjacency    map[string]map[string]bool
	knownSeeds       map[string]bool
	intrinsicSeeds   map[string]bool
	uncertainFacts   map[bool]map[string]bool
	uncertainCalls   map[string]bool
	moduleArrayUsers map[string][]string
}

type arrayCandidateLineKey struct {
	line int
	kind string
}

type arrayCandidateQualifiedKindKey struct {
	qualified string
	kind      string
}

type arrayCandidateIndex struct {
	byName          map[string]string
	byQualified     map[string]string
	byQualifiedKind map[arrayCandidateQualifiedKindKey]string
	byLineAndKind   map[arrayCandidateLineKey]string
}

// buildArrayCandidateIndex preserves the old sorted-key tie breaking while
// avoiding a project-wide key collection and sort for every uncertain call.
type arrayByRefEntryEvidence struct {
	seen                bool
	allocated           bool
	conditionCompatible bool
	condition           string
}

type arrayByRefCallCandidate struct {
	key    string
	target sourceProcedure
	call   procedureir.CallSite
}

type arrayLocalGoSubSummary struct {
	guaranteedAllocated map[string]bool
	unknown             map[string]bool
}

type arrayLocalGoSubAllocations map[string]arrayLocalGoSubSummary
