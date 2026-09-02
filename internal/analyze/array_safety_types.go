package analyze

import "regexp"

type arrayCFGStrategy uint8

const (
	arrayCFGStrategyAuto arrayCFGStrategy = iota
	arrayCFGStrategyCompact
	arrayCFGStrategyLegacy
)

type arrayFallbackReason uint8

const (
	arrayFallbackUnsupported arrayFallbackReason = iota
	arrayFallbackEmptyState
	arrayFallbackIndex
)

// arrayAllocation is deliberately a three-point lattice.  In particular,
// unknown is not treated as allocated: a VBA runtime operation must be
// proven safe on every path before it is allowed through the analysis.
type arrayAllocation uint8

const (
	arrayUnknown arrayAllocation = iota
	arrayUnallocated
	arrayAllocated
)

type arrayBound struct {
	known      bool
	value      int
	expression string
}

type arrayDimension struct {
	lower arrayBound
	upper arrayBound
}

type arrayOrigin uint8

const (
	arrayOriginUnknown arrayOrigin = iota
	arrayOriginLocal
	arrayOriginRangeValue
)

type arrayVariable struct {
	name      string
	typ       string
	isArray   bool
	isVariant bool
	isObject  bool
	// knownScalar is intentionally narrower than "not an array".  Only
	// built-in scalar declarations are strong enough to prove that an array
	// operation or For Each source is invalid.  User-defined classes/UDTs and
	// unresolved external types remain unknown and therefore fail open.
	knownScalar bool
	fixed       bool
	parameter   bool
	static      bool
	paramArray  bool
	dimensions  []arrayDimension
}

type arrayValue struct {
	kind            arrayAllocation
	knownArray      bool
	mayBeEmpty      bool
	dimensions      []arrayDimension
	preserveShape   []arrayDimension
	origin          arrayOrigin
	allocationProbe string
	safeBoundProbe  string
	// allocationCountSource records a narrow conditional allocation contract:
	// the array is allocated when the named scalar is positive, or when the
	// named collection's Count is positive. The fact is refined only on a
	// matching control-flow branch; it is never treated as unconditional.
	allocationCountSource string
	// conditionalAllocationSource records that a direct ReDim occurred under
	// the true body of a simple scalar comparison. The fact survives the merge
	// with the branch that skipped the ReDim and is consumed only by a later
	// matching true branch.
	conditionalAllocationSource string
	// allocationFlagSource records a local Boolean assigned after a successful
	// ReDim. A later true check of that flag can recover the allocation fact
	// after control-flow joins such as a discovery loop.
	allocationFlagSource string
	// returnNonEmptyArrayParameter is set only on an interprocedural return
	// summary. It names the callee's ByRef array parameter whose non-empty value
	// makes the returned array allocation definite.
	returnNonEmptyArrayParameter string
	// returnPositiveScalarParameter is set only on an interprocedural return
	// summary. It names the callee's scalar parameter whose positive value makes
	// the returned array allocation definite.
	returnPositiveScalarParameter string
	// nonEmptySource records the caller-side array whose non-empty state makes
	// this returned array non-empty. It is consumed by a matching StrPtr guard.
	nonEmptySource string
	// returnDescriptor* records a narrow typed-array return whose SAFEARRAY
	// descriptor is populated from scalar function parameters. The metadata is
	// consumed at the caller so known arguments can recover the returned shape;
	// unknown arguments retain only the allocation proof.
	returnDescriptorSourceParameter string
	returnDescriptorStartParameter  string
	returnDescriptorLengthParameter string
	returnDescriptorLowerParameter  string
	boundsProof                     arrayBoundsProof
}

type arrayFlowState map[string]arrayValue

type arrayBoundsProof struct {
	loopEndLine                      int
	priorKind                        arrayAllocation
	priorKnownArray                  bool
	priorMayBeEmpty                  bool
	priorAllocationCount             string
	priorConditionalAllocationSource string
}

const arrayDictionaryCountSourcePrefix = "dictionary:"

// arrayResumeNextCapacityGuard describes the narrow growable-buffer idiom
// where an unallocated array is probed with UBound under Resume Next, a
// capacity fallback ReDim Preserve allocates it when data is present, and a
// following loop writes only the requested range. The guard is used only for
// the probe and that loop's indexed writes; it does not globally mark the
// array allocated.
type arrayResumeNextCapacityGuard struct {
	target         string
	probeLine      int
	indexStartLine int
	indexEndLine   int
}

type arrayUse struct {
	name string
	args []string
}

type directArrayRedimClause struct {
	name       string
	dimensions string
}

var (
	arrayRedimRe           = regexp.MustCompile(`(?i)^\s*redim\s+(preserve\s+)?(.+)$`)
	arrayRedimClauseRe     = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?)?\s*$`)
	arrayRedimTypeSuffixRe = regexp.MustCompile(`(?i)^as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?$`)
	arrayEraseRe           = regexp.MustCompile(`(?i)^\s*erase\s+(.+)$`)
	arrayEraseNameRe       = regexp.MustCompile(`(?i)^[A-Za-z_]\w*$`)
	arrayBoundCallRe       = regexp.MustCompile(`(?i)\b(lbound|ubound)\s*\(\s*([^,)]*)\s*(?:,\s*([^)]*))?\)`)
	arrayForBoundRe        = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+(?:lbound|ubound)\s*\(\s*([A-Za-z_]\w*)`)
	arrayForUBoundRe       = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+ubound\s*\(\s*([A-Za-z_]\w*)`)
	// A loop with this shape is safe for a zero-based Byte array when its
	// length was obtained from a successful UBound(array) + 1 expression.
	arrayForZeroBasedLengthRe         = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*0\s+to\s+([A-Za-z_]\w*)\s*-\s*1\s*$`)
	arrayDoWhileLengthRe              = regexp.MustCompile(`(?i)^\s*do\s+while\s+([A-Za-z_]\w*)\s*<\s*([A-Za-z_]\w*)\s*$`)
	arrayUBoundPlusOneRe              = regexp.MustCompile(`(?i)^\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayReturnArrayDocRe             = regexp.MustCompile(`(?i)^@returns?\s+(?:(?:variant|object)\s*<)?array(?:<|\b)`)
	arrayTypeNameExpressionRe         = regexp.MustCompile(`(?i)^typename\s*\(\s*([A-Za-z_]\w*)\s*\)$`)
	arrayQuotedCaseRe                 = regexp.MustCompile(`^\s*"([^"]*)"\s*$`)
	arrayEmptyGuardRe                 = regexp.MustCompile(`(?i)^\s*\(?\s*not\s+([A-Za-z_]\w*)\s*\)?\s*=\s*-1\s*$`)
	arrayForScalarBoundRe             = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*[-+]?\d+\s+to\s+([A-Za-z_]\w*)\s*$`)
	arrayForCountRe                   = regexp.MustCompile(`(?i)^\s*for\s+([A-Za-z_]\w*)\s*=\s*(0|1)\s+to\s+([A-Za-z_]\w*)\s*\.\s*count(\s*-\s*1)?\s*$`)
	arrayDimensionCountLoopRe         = regexp.MustCompile(`(?i)^\s*for\s+([A-Za-z_]\w*)\s*=\s*1\s+to\s+[A-Za-z_]\w*\s*$`)
	arrayForEachRe                    = regexp.MustCompile(`(?i)^\s*for\s+each\s+[A-Za-z_]\w*\s+in\s+([^\r\n]+)`)
	arrayIndexedSourceRe              = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\(`)
	arrayGuardCallRe                  = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*(?:(=|<>|>=|<=|>|<)\s*(-?\d+))?\s*$`)
	arrayGuardReversedRe              = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*$`)
	arrayGuardValueRe                 = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayGuardValueReversedRe         = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*)\s*$`)
	arrayIsArrayGuardRe               = regexp.MustCompile(`(?i)^\s*isarray\s*\(\s*(.+)\s*\)\s*$`)
	arrayByteArrayGuardRe             = regexp.MustCompile(`(?i)^\s*(?:vartypeof|vartype)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*=\s*\(?\s*vbarray\s+or\s+vbbyte\s*\)?\s*$`)
	arrayStrPtrGuardRe                = regexp.MustCompile(`(?i)^\s*strptr\s*\(\s*([A-Za-z_]\w*)\s*\)\s*(=|<>)\s*0\s*$`)
	arraySafeArrayZeroExitGuardRe     = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arraySafeArrayPointerCopyRe       = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:[A-Za-z_]\w*\.)?copymemoryfromptr\s+([A-Za-z_]\w*)\s*,\s*([A-Za-z_]\w*)\s*,\s*lenb\s*\(\s*([A-Za-z_]\w*)\s*\)\s*$`)
	arrayByteArrayReadRe              = regexp.MustCompile(`(?i)^\s*(?:[A-Za-z_]\w*\.)*read\s*\(\s*-1\s*\)\s*$`)
	arraySetupGuardRe                 = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s+then\s+exit\s+sub\s*$`)
	arrayStaticReadyGuardRe           = regexp.MustCompile(`(?i)^\s*if\s+not\s+([A-Za-z_]\w*)\s*\.\s*isset\s+then\s*$`)
	arrayModuleReadyGuardRe           = regexp.MustCompile(`(?i)^\s*if\s+not\s+([A-Za-z_]\w*)\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayOnErrorGotoRe                = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+([A-Za-z_]\w*)\s*$`)
	arrayOnErrorResumeNextRe          = regexp.MustCompile(`(?i)^\s*on\s+error\s+resume\s+next\s*$`)
	arrayOnErrorResumeNextStatementRe = regexp.MustCompile(`(?i)(?:^|\bthen\s+)on\s+error\s+resume\s+next(?:\s+else\b.*)?$`)
	arrayOnErrorGotoZeroRe            = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+0\s*$`)
	arrayErrNumberFailureRe           = regexp.MustCompile(`(?i)^\s*if\s+err\.number\s*<>\s*0\s+then\s*$`)
	arrayCapacityProbeRe              = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayBoundsProbeRe                = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayCheckedProbeExitRe           = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*(?:<=|=)\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayCapacityIfRe                 = regexp.MustCompile(`(?i)^\s*if\s+.+\s*>\s*([A-Za-z_]\w*)\s+then\s*$`)
	arrayCapacityComparisonRe         = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(>|<)\s*([A-Za-z_]\w*)\s*$`)
	arrayPositiveLengthExpressionRe   = regexp.MustCompile(`(?i)^\s*(?:[A-Za-z_]\w*\.)?len\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\*\s*(\d+)\s*$`)
	arrayZeroLengthConditionRe        = regexp.MustCompile(`(?i)^\s*(?:[A-Za-z_]\w*\.)?len\s*\(\s*([A-Za-z_]\w*)\s*\)\s*=\s*0\s*$`)
	arrayForZeroToCountRe             = regexp.MustCompile(`(?i)^\s*for\s+[A-Za-z_]\w*\s*=\s*0\s+to\s+[A-Za-z_]\w*\s*-\s*1\s*$`)
	arrayLabelRe                      = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*:\s*$`)
	arrayCountComparisonRe            = regexp.MustCompile(`(?i)^\s*(.*?)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayConditionAndRe               = regexp.MustCompile(`(?i)\s+and\s+`)
	arrayConditionOrRe                = regexp.MustCompile(`(?i)\s+or\s+`)
	arrayScalarConditionRe            = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(=|<>)\s*([A-Za-z_]\w*|-?\d+)\s*$`)
	arrayScalarConditionReversedRe    = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>)\s*([A-Za-z_]\w*)\s*$`)
	arraySelectCaseRe                 = regexp.MustCompile(`(?i)^select\s+case\s+(.+)$`)
	arrayPositiveCaseRe               = regexp.MustCompile(`(?i)^case\s+(-?\d+)\s*$`)
)
