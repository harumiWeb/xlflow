package procedureir

import (
	"math"
	"math/big"
	"regexp"
	"strings"
)

var safeProbeResultInspectionRE = regexp.MustCompile(`(?i)^isarray\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)$`)
var safeProbeIdentifierRE = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)
var safeArrayBoundsLowerBoundRE = regexp.MustCompile(`(?i)^lbound\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)$`)
var safeArrayBoundsLengthRE = regexp.MustCompile(`(?i)^ubound\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)\s*-\s*([A-Za-z_][A-Za-z0-9_]*)\s*\+\s*1$`)
var errorNumberNonZeroConditionRE = regexp.MustCompile(`(?i)^\s*\(?\s*(?:err\s*\.\s*number\s*<>\s*0|0\s*<>\s*err\s*\.\s*number)\s*\)?(?:\s+or\s+[^()]+)?\s*$`)
var errorNumberZeroConditionRE = regexp.MustCompile(`(?i)^\s*\(?\s*(?:err\s*\.\s*number\s*=\s*0|0\s*=\s*err\s*\.\s*number)\s*\)?\s*$`)
var errorConditionAndRE = regexp.MustCompile(`(?:^|[^a-z0-9_])and(?:$|[^a-z0-9_])`)
var errorConditionNotRE = regexp.MustCompile(`(?:^|[^a-z0-9_])not(?:$|[^a-z0-9_])`)
var errorConditionUnsupportedLogicalRE = regexp.MustCompile(`(?:^|[^a-z0-9_])(?:xor|eqv|imp)(?:$|[^a-z0-9_])`)

// SafeArrayBoundsProbeKind identifies the two assignments that can make up a
// checked array-bounds compatibility probe.
type SafeArrayBoundsProbeKind uint8

const (
	SafeArrayBoundsProbeNone SafeArrayBoundsProbeKind = iota
	SafeArrayBoundsProbeLowerBound
	SafeArrayBoundsProbeLength
)

// SafeArrayBoundsProbeAssignment recognizes only the scalar, two-step shape
// used to probe an unallocated/dynamic array safely:
//
//	lowerBound = LBound(values)
//	itemCount = UBound(values) - lowerBound + 1
//
// The caller supplies the assignment target and right-hand side separately so
// recovered or compound assignment syntax cannot be accepted accidentally.
// It returns normalized identifiers for matching the two steps.
func SafeArrayBoundsProbeAssignment(target, value string) (array, lowerTarget string, kind SafeArrayBoundsProbeKind, ok bool) {
	target = strings.TrimSpace(target)
	if !safeProbeIdentifierRE.MatchString(target) {
		return "", "", SafeArrayBoundsProbeNone, false
	}
	value = strings.TrimSpace(value)
	if match := safeArrayBoundsLowerBoundRE.FindStringSubmatch(value); len(match) == 2 {
		return strings.ToLower(match[1]), strings.ToLower(target), SafeArrayBoundsProbeLowerBound, true
	}
	if match := safeArrayBoundsLengthRE.FindStringSubmatch(value); len(match) == 3 {
		return strings.ToLower(match[1]), strings.ToLower(match[2]), SafeArrayBoundsProbeLength, true
	}
	return "", "", SafeArrayBoundsProbeNone, false
}

// SafeProbeResultInspection reports the narrow intrinsic inspection currently
// supported after a single Resume Next probe. IsArray only observes the probe
// value and does not itself introduce another potentially failing operation.
// Keep this allowlist narrow: callers use it to avoid treating an inspection
// as a second protected operation.
func SafeProbeResultInspection(value, probeTarget string) bool {
	probeTarget = strings.TrimSpace(probeTarget)
	if probeTarget == "" {
		return false
	}
	match := safeProbeResultInspectionRE.FindStringSubmatch(strings.TrimSpace(value))
	return len(match) == 2 && strings.EqualFold(match[1], probeTarget)
}

// ErrorNumberThenBranchIsFailure reports the branch polarity for the narrow
// Err.Number zero comparisons used by checked Resume Next recovery. Unknown,
// negated, and And-combined polarity is rejected so callers do not mistake a
// success branch for an error fallback. A direct nonzero comparison may be
// combined with Or because that arm alone guarantees the error branch.
func ErrorNumberThenBranchIsFailure(condition string) (thenFailure, known bool) {
	condition = strings.ToLower(maskVBALiterals(condition))
	condition = strings.Join(strings.Fields(condition), " ")
	if errorConditionAndRE.MatchString(condition) || errorConditionNotRE.MatchString(condition) || errorConditionUnsupportedLogicalRE.MatchString(condition) {
		return false, false
	}
	switch {
	case errorNumberNonZeroConditionRE.MatchString(condition):
		return true, true
	case errorNumberZeroConditionRE.MatchString(condition):
		return false, true
	default:
		return false, false
	}
}

func maskVBALiterals(value string) string {
	var masked strings.Builder
	masked.Grow(len(value))
	inString := false
	for index := 0; index < len(value); index++ {
		if value[index] == '"' {
			masked.WriteByte(' ')
			if inString && index+1 < len(value) && value[index+1] == '"' {
				masked.WriteByte(' ')
				index++
				continue
			}
			inString = !inString
			continue
		}
		if inString {
			masked.WriteByte(' ')
		} else {
			masked.WriteByte(value[index])
		}
	}
	return masked.String()
}

// SafeLiteralAssignment reports whether value can be assigned to targetType
// without relying on a potentially failing VBA coercion. It is intentionally
// conservative because callers use it to recognize recovery assignments made
// while On Error Resume Next is active.
func SafeLiteralAssignment(value, targetType string) bool {
	value = strings.TrimSpace(value)
	targetType = normalizeLiteralTargetType(targetType)
	if value == "" || targetType == "" {
		return false
	}
	if strings.HasSuffix(targetType, "()") {
		return false
	}

	lower := strings.ToLower(value)
	switch lower {
	case "true", "false":
		return targetType == "variant" || targetType == "string" || isNumericLiteralTarget(targetType) || targetType == "boolean"
	case "nothing":
		return targetType == "variant" || isObjectLiteralTarget(targetType)
	case "empty":
		return targetType == "variant" || isScalarLiteralTarget(targetType)
	case "null":
		return targetType == "variant"
	}
	if isCompleteStringLiteral(value) {
		return targetType == "variant" || targetType == "string"
	}

	numeric, ok := parseNumericLiteral(value)
	if !ok {
		return false
	}
	if targetType == "variant" || targetType == "string" || targetType == "boolean" {
		// Even a Variant or a string assignment must start with a valid VBA
		// numeric literal. Values beyond Double's range are rejected by VBA
		// before the target conversion can make them safe.
		return rationalAbsWithin(numeric, new(big.Rat).SetFloat64(math.MaxFloat64))
	}
	switch targetType {
	case "byte":
		return integerLiteralWithin(numeric, "0", "255")
	case "integer":
		return integerLiteralWithin(numeric, "-32768", "32767")
	case "long":
		return integerLiteralWithin(numeric, "-2147483648", "2147483647")
	case "longlong":
		return integerLiteralWithin(numeric, "-9223372036854775808", "9223372036854775807")
	case "longptr":
		// The analyzer does not know whether the workbook runs in 32-bit or
		// 64-bit Office, so use the narrower ABI-compatible range.
		return integerLiteralWithin(numeric, "-2147483648", "2147483647")
	case "single":
		return rationalAbsWithin(numeric, new(big.Rat).SetFloat64(float64(math.MaxFloat32)))
	case "double":
		return rationalAbsWithin(numeric, new(big.Rat).SetFloat64(math.MaxFloat64))
	case "currency":
		return rationalWithin(numeric, "-922337203685477.5808", "922337203685477.5807")
	case "decimal":
		return rationalWithin(numeric, "-79228162514264337593543950335", "79228162514264337593543950335")
	case "date":
		return rationalWithin(numeric, "-657434", "2958465")
	default:
		return false
	}
}

func normalizeLiteralTargetType(targetType string) string {
	targetType = strings.ToLower(strings.TrimSpace(targetType))
	targetType = strings.TrimPrefix(targetType, "byref ")
	return targetType
}

func isNumericLiteralTarget(targetType string) bool {
	switch targetType {
	case "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal":
		return true
	default:
		return false
	}
}

func isScalarLiteralTarget(targetType string) bool {
	return targetType == "variant" || targetType == "string" || targetType == "boolean" || targetType == "date" || isNumericLiteralTarget(targetType)
}

func isObjectLiteralTarget(targetType string) bool {
	switch targetType {
	case "object", "application", "workbook", "worksheet", "range", "chart", "pivot table", "pivottable", "listobject", "dictionary", "collection", "window":
		return true
	default:
		return false
	}
}

func isCompleteStringLiteral(value string) bool {
	if len(value) < 2 || value[0] != '"' {
		return false
	}
	for index := 1; index < len(value); index++ {
		if value[index] != '"' {
			continue
		}
		if index+1 < len(value) && value[index+1] == '"' {
			index++
			continue
		}
		return index == len(value)-1
	}
	return false
}

func integerLiteralWithin(value *big.Rat, minText, maxText string) bool {
	min, minOK := new(big.Int).SetString(minText, 10)
	max, maxOK := new(big.Int).SetString(maxText, 10)
	if !minOK || !maxOK {
		return false
	}
	rounded := roundVBANumeric(value)
	return rounded.Cmp(min) >= 0 && rounded.Cmp(max) <= 0
}

func roundVBANumeric(value *big.Rat) *big.Int {
	abs := new(big.Rat).Abs(value)
	quotient := new(big.Int)
	remainder := new(big.Int)
	quotient.QuoRem(abs.Num(), abs.Denom(), remainder)
	twiceRemainder := new(big.Int).Lsh(new(big.Int).Set(remainder), 1)
	if twiceRemainder.Cmp(abs.Denom()) > 0 ||
		twiceRemainder.Cmp(abs.Denom()) == 0 && quotient.Bit(0) == 1 {
		quotient.Add(quotient, big.NewInt(1))
	}
	if value.Sign() < 0 {
		quotient.Neg(quotient)
	}
	return quotient
}

func rationalWithin(value *big.Rat, minText, maxText string) bool {
	min, minOK := new(big.Rat).SetString(minText)
	max, maxOK := new(big.Rat).SetString(maxText)
	return minOK && maxOK && value.Cmp(min) >= 0 && value.Cmp(max) <= 0
}

func rationalAbsWithin(value, max *big.Rat) bool {
	return new(big.Rat).Abs(value).Cmp(max) <= 0
}

func parseNumericLiteral(value string) (*big.Rat, bool) {
	value = strings.TrimRight(strings.TrimSpace(value), "%&^!#@")
	if value == "" {
		return nil, false
	}
	numeric, ok := new(big.Rat).SetString(value)
	if !ok || numeric == nil {
		return nil, false
	}
	return numeric, true
}
