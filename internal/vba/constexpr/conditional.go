package constexpr

import (
	"cmp"
	"math"
	"strconv"
	"strings"
)

func conditionalLiteral(text string) Result {
	if digits, currency := strings.CutSuffix(text, "@"); currency {
		// Parse decimal fixed-point units without a float64 intermediate.
		// Rounding extra fractional digits and exponent notation is not
		// modelled here: neither may decide which source branch to discard.
		whole, fraction, _ := strings.Cut(digits, ".")
		if len(fraction) > 4 {
			return unknown("conditional Currency literal rounding is not modelled")
		}
		if whole == "" {
			whole = "0"
		}
		units := whole + fraction + strings.Repeat("0", 4-len(fraction))
		for _, digit := range units {
			if digit < '0' || digit > '9' {
				return unknown("conditional Currency literal notation is not modelled")
			}
		}
		value, err := strconv.ParseInt(units, 10, 64)
		if err != nil {
			return unknown("conditional Currency literal overflow")
		}
		return known(Value{Kind: ValueCurrency, Currency: value})
	}
	if len(text) > 1 && text[0] != '&' && strings.ContainsAny(text[len(text)-1:], "!#@") && !strings.ContainsAny(text, ".eE") {
		text = text[:len(text)-1] + ".0" + text[len(text)-1:]
	}
	result := parseNumber(text)
	if result.Kind != Known || len(text) < 3 || text[0] != '&' {
		return result
	}
	// VBA infers the signed width of hexadecimal/octal literals. Preserve
	// uncertain branches for sign-boundary values until that width contract
	// is modelled, rather than inheriting the general unsigned projection.
	limit := int64(math.MaxInt16)
	if strings.HasSuffix(text, "&") {
		limit = math.MaxInt32
	}
	if strings.HasSuffix(text, "^") {
		limit = math.MaxInt64
	}
	if result.Typed.Integer > limit {
		return unknown("signed conditional literal width is not modelled")
	}
	return result
}

func conditionalCompare(left, right Result, op string) Result {
	if left.Kind == Known && right.Kind == Known && (left.Typed.Kind == ValueString || right.Typed.Kind == ValueString) {
		return unknown("conditional string comparison depends on Option Compare")
	}
	if left.Kind == Known && right.Kind == Known {
		if left.Typed.Kind == ValueCurrency && right.Typed.Kind == ValueCurrency {
			// Only the exact ordering (-1, 0, 1) enters the common operator
			// selector; converting the scaled values to float64 loses units.
			return compareFloat(float64(cmp.Compare(left.Typed.Currency, right.Typed.Currency)), 0, op)
		}
		// The shared comparator uses float64. Beyond its exact integer range,
		// adjacent LongLong values can collapse to the same number.
		for _, value := range []Value{left.Typed, right.Typed} {
			if isIntegral(value) && (value.Integer > 1<<53 || value.Integer < -(1<<53)) {
				return unknown("conditional wide numeric comparison precision is not modelled")
			}
		}
		lk, rk := left.Typed.Kind, right.Typed.Kind
		if lk != rk && (lk == ValueSingle || rk == ValueSingle || lk == ValueCurrency || rk == ValueCurrency) {
			return unknown("conditional mixed-precision comparison rounding is not modelled")
		}
	}
	return compare(conditionalNumber(left), conditionalNumber(right), op)
}

// conditionalNumber applies only VBA's unambiguous Boolean/Empty numeric
// coercions. Strings, dates, Null, fractional bitwise rounding and overflow
// remain conservative instead of guessing a target-dependent result.
func conditionalNumber(result Result) Result {
	if result.Kind != Known {
		return result
	}
	switch result.Typed.Kind {
	case ValueBoolean:
		value := int64(0)
		if result.Typed.Boolean {
			value = -1
		}
		return known(Value{Kind: ValueLongLong, Integer: value})
	case ValueEmpty:
		return known(Value{Kind: ValueLongLong, Integer: 0})
	}
	return result
}

func conditionalUnary(value Result, op string) Result {
	if op == "not" && value.Kind == Known && value.Typed.Kind == ValueBoolean {
		return unary(value, op)
	}
	value = conditionalNumber(value)
	if op == "not" && value.Kind == Known && isIntegral(value.Typed) {
		return known(Value{Kind: ValueLongLong, Integer: ^value.Typed.Integer})
	}
	return unary(value, op)
}

func conditionalCombine(left, right Result, op string) Result {
	if left.Kind == Known && right.Kind == Known && (left.Typed.Kind == ValueCurrency || right.Typed.Kind == ValueCurrency) {
		return unknown("conditional Currency arithmetic rounding is not modelled")
	}
	if left.Kind == Known && right.Kind == Known && (left.Typed.Kind == ValueSingle || right.Typed.Kind == ValueSingle) {
		return unknown("conditional Single arithmetic rounding is not modelled")
	}
	if op == "&" {
		return combine(left, right, op)
	}
	if left.Kind == Known && right.Kind == Known && left.Typed.Kind == ValueBoolean && right.Typed.Kind == ValueBoolean {
		switch op {
		case "and", "or", "xor", "eqv", "imp":
			return combine(left, right, op)
		}
	}
	left, right = conditionalNumber(left), conditionalNumber(right)
	if left.Kind == Known && right.Kind == Known && isIntegral(left.Typed) && isIntegral(right.Typed) {
		lv, rv := left.Typed.Integer, right.Typed.Integer
		var value int64
		switch op {
		case "and":
			value = lv & rv
		case "or":
			value = lv | rv
		case "xor":
			value = lv ^ rv
		case "eqv":
			value = ^(lv ^ rv)
		case "imp":
			value = ^lv | rv
		default:
			return combine(left, right, op)
		}
		return known(Value{Kind: ValueLongLong, Integer: value})
	}
	return combine(left, right, op)
}
