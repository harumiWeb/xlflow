// Package constexpr evaluates the deliberately conservative subset of VBA
// expressions that may be used where a compile-time value is required.  It
// never invokes VBA, Excel, or COM.  Runtime-dependent and unsupported
// expressions are represented as Unknown so callers can fail open.
package constexpr

import (
	"math"
	"strconv"
	"strings"
	"unicode"
)

type ResultKind string

const (
	Known   ResultKind = "known"
	Unknown ResultKind = "unknown"
	Invalid ResultKind = "invalid"
)

// ValueKind is the stable type tag carried by a known result.
type ValueKind string

const (
	ValueInteger  ValueKind = "integer"
	ValueLong     ValueKind = "long"
	ValueLongLong ValueKind = "longlong"
	ValueSingle   ValueKind = "single"
	ValueDouble   ValueKind = "double"
	ValueCurrency ValueKind = "currency"
	ValueString   ValueKind = "string"
	ValueBoolean  ValueKind = "boolean"
	ValueDate     ValueKind = "date"
	ValueEmpty    ValueKind = "empty"
	ValueNull     ValueKind = "null"
	ValueNothing  ValueKind = "nothing"
)

// Value is an immutable, tagged VBA value.  Only the field corresponding to
// Kind is meaningful. Currency is stored as an integer scaled by 10,000.
type Value struct {
	Kind     ValueKind
	Integer  int64
	Float    float64
	String   string
	Boolean  bool
	Currency int64
}

// Environment resolves case-insensitive constant names. Implementations must
// be immutable for the duration of one evaluation pass.
type Environment interface {
	Resolve(name string) (Value, bool)
}

// Values adapts a map to Environment. Keys are matched case-insensitively.
type Values map[string]Value

func (v Values) Resolve(name string) (Value, bool) {
	name = strings.ToLower(strings.TrimSpace(name))
	// A map may contain keys that differ only by case when declarations come
	// from separate snapshots. Pick the lexicographically smallest spelling so
	// evaluation does not depend on Go's randomized map iteration order.
	selected := ""
	var value Value
	for key, candidate := range v {
		if !strings.EqualFold(key, name) || (selected != "" && key >= selected) {
			continue
		}
		selected = key
		value = candidate
	}
	if selected == "" {
		return Value{}, false
	}
	return value, true
}

// Result keeps the historical integer Value field for existing consumers and
// adds Typed for all supported scalar kinds.
type Result struct {
	Kind   ResultKind
	Value  int
	Typed  Value
	Type   string
	Reason string
}

// Evaluate evaluates expression against env. A nil environment still knows
// VBA's intrinsic literal values but leaves named constants unresolved.
func Evaluate(expression string, env Environment) Result {
	lexer := newLexer(expression)
	if lexer.invalid != "" {
		return Result{Kind: Invalid, Reason: lexer.invalid}
	}
	if len(lexer.tokens) == 0 {
		return Result{Kind: Unknown, Reason: "empty expression"}
	}
	p := parser{tokens: lexer.tokens, env: env}
	result := p.parseImp()
	if result.Kind == Invalid {
		return result
	}
	if p.pos != len(p.tokens) {
		if p.tokens[p.pos].kind == tokenOperator {
			return Result{Kind: Unknown, Reason: "unsupported operator " + p.tokens[p.pos].text}
		}
		return Result{Kind: Invalid, Reason: "unexpected token"}
	}
	return result
}

// EvaluateValues is a convenience adapter for callers that own a value map.
func EvaluateValues(expression string, values map[string]Value) Result {
	return Evaluate(expression, Values(values))
}

// EvaluateInteger preserves the original API used by array and loop analysis.
// It returns Known only for an integral value representable by int and keeps
// non-integral, typed, and overflowing results fail-open.
func EvaluateInteger(expression string, constants map[string]int) Result {
	values := make(Values, len(constants))
	for name, value := range constants {
		// The compatibility map has no VBA declaration type. LongLong keeps
		// host-sized integer bounds from being mistaken for 16-bit Integer
		// literals while the final int projection still enforces its own width.
		values[name] = Value{Kind: ValueLongLong, Integer: int64(value)}
	}
	result := Evaluate(expression, values)
	if result.Kind != Known {
		return result
	}
	if !isIntegral(result.Typed) || int64(int(result.Typed.Integer)) != result.Typed.Integer {
		return Result{Kind: Unknown, Reason: "non-integral or unrepresentable integer"}
	}
	result.Value = int(result.Typed.Integer)
	result.Type = string(result.Typed.Kind)
	return result
}

func known(value Value) Result {
	if isIntegral(value) && !fitsInteger(value.Kind, value.Integer) {
		return unknown("integer overflow")
	}
	result := Result{Kind: Known, Typed: value, Type: string(value.Kind)}
	if isIntegral(value) {
		if int64(int(value.Integer)) == value.Integer {
			result.Value = int(value.Integer)
		}
	}
	return result
}

func unknown(reason string) Result { return Result{Kind: Unknown, Reason: reason} }
func invalid(reason string) Result { return Result{Kind: Invalid, Reason: reason} }

func isIntegral(value Value) bool {
	return value.Kind == ValueInteger || value.Kind == ValueLong || value.Kind == ValueLongLong
}

func fitsInteger(kind ValueKind, value int64) bool {
	switch kind {
	case ValueInteger:
		return value >= math.MinInt16 && value <= math.MaxInt16
	case ValueLong:
		return value >= math.MinInt32 && value <= math.MaxInt32
	case ValueLongLong:
		return true
	default:
		return false
	}
}

type tokenKind uint8

const (
	tokenNumber tokenKind = iota
	tokenString
	tokenDate
	tokenIdentifier
	tokenOperator
	tokenLParen
	tokenRParen
)

type token struct {
	kind tokenKind
	text string
}

type lexer struct {
	tokens  []token
	invalid string
}

func newLexer(text string) lexer {
	var out lexer
	for i := 0; i < len(text); {
		if text[i] == ' ' || text[i] == '\t' || text[i] == '\r' || text[i] == '\n' {
			i++
			continue
		}
		switch text[i] {
		case '(':
			out.tokens = append(out.tokens, token{kind: tokenLParen, text: "("})
			i++
		case ')':
			out.tokens = append(out.tokens, token{kind: tokenRParen, text: ")"})
			i++
		case '"':
			i++
			var b strings.Builder
			for i < len(text) {
				if text[i] != '"' {
					b.WriteByte(text[i])
					i++
					continue
				}
				if i+1 < len(text) && text[i+1] == '"' {
					b.WriteByte('"')
					i += 2
					continue
				}
				i++
				out.tokens = append(out.tokens, token{kind: tokenString, text: b.String()})
				break
			}
			if i > len(text) || (i == len(text) && (len(text) == 0 || text[i-1] != '"')) {
				out.invalid = "unterminated string literal"
				return out
			}
		case '#':
			start := i
			i++
			for i < len(text) && text[i] != '#' {
				i++
			}
			if i >= len(text) {
				out.invalid = "unterminated date literal"
				return out
			}
			i++
			out.tokens = append(out.tokens, token{kind: tokenDate, text: text[start+1 : i-1]})
		case '&':
			if i+1 < len(text) && (text[i+1] == 'H' || text[i+1] == 'h' || text[i+1] == 'O' || text[i+1] == 'o') {
				start := i
				i += 2
				for i < len(text) && (unicode.IsDigit(rune(text[i])) || unicode.IsLetter(rune(text[i]))) {
					i++
				}
				out.tokens = append(out.tokens, token{kind: tokenNumber, text: text[start:i]})
			} else {
				out.tokens = append(out.tokens, token{kind: tokenOperator, text: "&"})
				i++
			}
		case '<', '>':
			start := i
			i++
			if i < len(text) && (text[i] == '=' || text[i] == '>') {
				i++
			}
			out.tokens = append(out.tokens, token{kind: tokenOperator, text: text[start:i]})
		case '=', '+', '-', '*', '/', '\\', '^':
			out.tokens = append(out.tokens, token{kind: tokenOperator, text: text[i : i+1]})
			i++
		default:
			if unicode.IsDigit(rune(text[i])) || text[i] == '.' {
				start := i
				i++
				for i < len(text) && (unicode.IsDigit(rune(text[i])) || text[i] == '.' || text[i] == 'e' || text[i] == 'E' || text[i] == '+' || text[i] == '-' || strings.ContainsRune("%&!#@^", rune(text[i]))) {
					if (text[i] == '+' || text[i] == '-') && text[i-1] != 'e' && text[i-1] != 'E' {
						break
					}
					i++
				}
				out.tokens = append(out.tokens, token{kind: tokenNumber, text: text[start:i]})
				continue
			}
			if isIdentifierStart(text[i]) {
				start := i
				i++
				for i < len(text) && (isIdentifierPart(text[i]) || text[i] == '.' || text[i] == '$') {
					i++
				}
				word := text[start:i]
				lower := strings.ToLower(word)
				switch lower {
				case "mod", "and", "or", "xor", "eqv", "imp", "like", "is", "not":
					out.tokens = append(out.tokens, token{kind: tokenOperator, text: lower})
				default:
					out.tokens = append(out.tokens, token{kind: tokenIdentifier, text: word})
				}
				continue
			}
			out.invalid = "unexpected character"
			return out
		}
	}
	return out
}

func isIdentifierStart(ch byte) bool { return ch == '_' || unicode.IsLetter(rune(ch)) }
func isIdentifierPart(ch byte) bool  { return isIdentifierStart(ch) || unicode.IsDigit(rune(ch)) }

type parser struct {
	tokens []token
	pos    int
	env    Environment
}

func (p *parser) parseImp() Result { return p.binary(p.parseEqv, "imp") }
func (p *parser) parseEqv() Result { return p.binary(p.parseOr, "eqv") }
func (p *parser) parseOr() Result  { return p.binary(p.parseXor, "or") }
func (p *parser) parseXor() Result { return p.binary(p.parseAnd, "xor") }
func (p *parser) parseAnd() Result { return p.binary(p.parseCompare, "and") }

type parseFn func() Result

func (p *parser) binary(next parseFn, operators ...string) Result {
	left := next()
	for p.pos < len(p.tokens) && p.tokens[p.pos].kind == tokenOperator && contains(operators, strings.ToLower(p.tokens[p.pos].text)) {
		op := strings.ToLower(p.tokens[p.pos].text)
		p.pos++
		right := next()
		left = combine(left, right, op)
	}
	return left
}

func (p *parser) parseCompare() Result {
	left := p.parseConcat()
	for p.pos < len(p.tokens) && p.tokens[p.pos].kind == tokenOperator && contains([]string{"=", "<>", "<", ">", "<=", ">=", "like", "is"}, strings.ToLower(p.tokens[p.pos].text)) {
		op := strings.ToLower(p.tokens[p.pos].text)
		p.pos++
		right := p.parseConcat()
		left = compare(left, right, op)
	}
	return left
}

func (p *parser) parseConcat() Result { return p.binary(p.parseAdd, "&") }
func (p *parser) parseAdd() Result    { return p.binary(p.parseMul, "+", "-") }
func (p *parser) parseMul() Result    { return p.binary(p.parseUnary, "*", "/", "\\", "mod") }

func (p *parser) parseUnary() Result {
	if p.pos < len(p.tokens) && p.tokens[p.pos].kind == tokenOperator && contains([]string{"+", "-", "not"}, strings.ToLower(p.tokens[p.pos].text)) {
		op := strings.ToLower(p.tokens[p.pos].text)
		p.pos++
		return unary(p.parseUnary(), op)
	}
	return p.parsePower()
}

func (p *parser) parsePower() Result {
	left := p.parsePrimary()
	if p.pos < len(p.tokens) && p.tokens[p.pos].kind == tokenOperator && p.tokens[p.pos].text == "^" {
		p.pos++
		return power(left, p.parseUnary())
	}
	return left
}

func (p *parser) parsePrimary() Result {
	if p.pos >= len(p.tokens) {
		return invalid("missing operand")
	}
	tok := p.tokens[p.pos]
	p.pos++
	if tok.kind == tokenLParen {
		result := p.parseImp()
		if p.pos >= len(p.tokens) || p.tokens[p.pos].kind != tokenRParen {
			return invalid("unclosed parentheses")
		}
		p.pos++
		return result
	}
	if tok.kind == tokenString {
		return known(Value{Kind: ValueString, String: tok.text})
	}
	if tok.kind == tokenDate {
		return known(Value{Kind: ValueDate, String: tok.text})
	}
	if tok.kind == tokenNumber {
		return parseNumber(tok.text)
	}
	if tok.kind != tokenIdentifier {
		return invalid("unexpected token")
	}
	if p.pos < len(p.tokens) && p.tokens[p.pos].kind == tokenLParen {
		p.skipCall()
		return unknown("runtime expression")
	}
	switch strings.ToLower(tok.text) {
	case "true":
		return known(Value{Kind: ValueBoolean, Boolean: true})
	case "false":
		return known(Value{Kind: ValueBoolean, Boolean: false})
	case "empty":
		return known(Value{Kind: ValueEmpty})
	case "null":
		return known(Value{Kind: ValueNull})
	case "nothing":
		return known(Value{Kind: ValueNothing})
	}
	if p.env != nil {
		if value, ok := p.env.Resolve(tok.text); ok {
			return known(value)
		}
	}
	return unknown("unresolved constant " + strings.ToLower(tok.text))
}

func (p *parser) skipCall() {
	depth := 0
	for p.pos < len(p.tokens) {
		switch p.tokens[p.pos].kind {
		case tokenLParen:
			depth++
		case tokenRParen:
			depth--
		}
		p.pos++
		if depth == 0 {
			return
		}
	}
}

func parseNumber(text string) Result {
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		return invalid("empty number")
	}
	suffix := byte(0)
	if strings.ContainsRune("%&!#@^", rune(trimmed[len(trimmed)-1])) {
		suffix = trimmed[len(trimmed)-1]
		trimmed = trimmed[:len(trimmed)-1]
	}
	base := 10
	digits := trimmed
	if len(trimmed) > 2 && trimmed[0] == '&' && (trimmed[1] == 'H' || trimmed[1] == 'h' || trimmed[1] == 'O' || trimmed[1] == 'o') {
		if trimmed[1] == 'H' || trimmed[1] == 'h' {
			base = 16
		}
		if trimmed[1] == 'O' || trimmed[1] == 'o' {
			base = 8
		}
		digits = trimmed[2:]
	}
	if base != 10 || (!strings.ContainsAny(trimmed, ".eE")) {
		value, err := strconv.ParseInt(digits, base, 64)
		if err != nil {
			return unknown("integer overflow")
		}
		kind := ValueLong
		switch suffix {
		case '%':
			if value < math.MinInt16 || value > math.MaxInt16 {
				return unknown("integer overflow")
			}
			kind = ValueInteger
		case '^':
			kind = ValueLongLong
		case '&':
			kind = ValueLong
		}
		return known(Value{Kind: kind, Integer: value})
	}
	value, err := strconv.ParseFloat(trimmed, 64)
	if err != nil || math.IsInf(value, 0) || math.IsNaN(value) {
		return unknown("floating-point overflow")
	}
	kind := ValueDouble
	switch suffix {
	case '!':
		kind = ValueSingle
		value = float64(float32(value))
		if math.IsInf(value, 0) {
			return unknown("single-precision overflow")
		}
	case '@':
		scaled := math.Round(value * 10000)
		if scaled < math.MinInt64 || scaled > math.MaxInt64 {
			return unknown("currency overflow")
		}
		return known(Value{Kind: ValueCurrency, Currency: int64(scaled)})
	case '#':
		kind = ValueDouble
	}
	return known(Value{Kind: kind, Float: value})
}

func unary(value Result, op string) Result {
	if value.Kind != Known {
		return value
	}
	if isRuntimeValue(value.Typed) {
		return unknown("operand coercion is not modelled")
	}
	switch op {
	case "+":
		if !isNumeric(value.Typed) {
			return invalid("unary plus requires numeric operand")
		}
		return value
	case "-":
		if isIntegral(value.Typed) {
			if value.Typed.Integer == math.MinInt64 {
				return unknown("integer overflow")
			}
			value.Typed.Integer = -value.Typed.Integer
			return known(value.Typed)
		}
		if isFloating(value.Typed) || value.Typed.Kind == ValueCurrency {
			if value.Typed.Kind == ValueCurrency {
				if value.Typed.Currency == math.MinInt64 {
					return unknown("currency overflow")
				}
				value.Typed.Currency = -value.Typed.Currency
				return known(value.Typed)
			}
			value.Typed.Float = -value.Typed.Float
			return known(value.Typed)
		}
		return invalid("unary minus requires numeric operand")
	case "not":
		if value.Typed.Kind != ValueBoolean {
			return invalid("Not requires Boolean operand")
		}
		value.Typed.Boolean = !value.Typed.Boolean
		return known(value.Typed)
	default:
		return unknown("unsupported unary operator")
	}
}

func power(left, right Result) Result {
	if left.Kind != Known || right.Kind != Known {
		return combineUnknown(left, right)
	}
	if !isNumeric(left.Typed) || !isNumeric(right.Typed) {
		return invalid("exponent requires numeric operands")
	}
	base := numericFloat(left.Typed)
	exponent := numericFloat(right.Typed)
	value := math.Pow(base, exponent)
	if math.IsInf(value, 0) || math.IsNaN(value) {
		return unknown("floating-point overflow")
	}
	if isIntegral(left.Typed) && isIntegral(right.Typed) && right.Typed.Integer >= 0 {
		if exact, ok := checkedPow(left.Typed.Integer, right.Typed.Integer); ok {
			return known(Value{Kind: widerInteger(left.Typed.Kind, right.Typed.Kind), Integer: exact})
		}
		return unknown("integer overflow")
	}
	return known(Value{Kind: ValueDouble, Float: value})
}

func checkedPow(base, exponent int64) (int64, bool) {
	result := int64(1)
	for exponent > 0 {
		if exponent&1 == 1 {
			if !checkedMul(result, base, &result) {
				return 0, false
			}
		}
		exponent >>= 1
		if exponent > 0 && !checkedMul(base, base, &base) {
			return 0, false
		}
	}
	return result, true
}

func combine(left, right Result, op string) Result {
	if left.Kind == Invalid {
		return left
	}
	if right.Kind == Invalid {
		return right
	}
	if left.Kind == Unknown || right.Kind == Unknown {
		return combineUnknown(left, right)
	}
	if isRuntimeValue(left.Typed) || isRuntimeValue(right.Typed) {
		return unknown("operand coercion is not modelled")
	}
	if op == "&" {
		if left.Typed.Kind == ValueString && right.Typed.Kind == ValueString {
			return known(Value{Kind: ValueString, String: left.Typed.String + right.Typed.String})
		}
		return invalid("string concatenation requires String operands")
	}
	if op == "and" || op == "or" || op == "xor" || op == "eqv" || op == "imp" {
		if left.Typed.Kind != ValueBoolean || right.Typed.Kind != ValueBoolean {
			return invalid("Boolean operator requires Boolean operands")
		}
		lv, rv := left.Typed.Boolean, right.Typed.Boolean
		var value bool
		switch op {
		case "and":
			value = lv && rv
		case "or":
			value = lv || rv
		case "xor":
			value = lv != rv
		case "eqv":
			value = lv == rv
		case "imp":
			value = !lv || rv
		}
		return known(Value{Kind: ValueBoolean, Boolean: value})
	}
	if !isNumeric(left.Typed) || !isNumeric(right.Typed) {
		return invalid("arithmetic requires numeric operands")
	}
	if op == "/" {
		divisor := numericFloat(right.Typed)
		if divisor == 0 {
			return invalid("division by zero")
		}
		return known(Value{Kind: ValueDouble, Float: numericFloat(left.Typed) / divisor})
	}
	if op == "\\" || op == "mod" {
		if !isIntegral(left.Typed) || !isIntegral(right.Typed) {
			return unknown("integer operator requires integral operands")
		}
		if right.Typed.Integer == 0 {
			return invalid("division by zero")
		}
		if op == "\\" {
			if left.Typed.Integer == math.MinInt64 && right.Typed.Integer == -1 {
				return unknown("integer overflow")
			}
			return known(Value{Kind: widerInteger(left.Typed.Kind, right.Typed.Kind), Integer: left.Typed.Integer / right.Typed.Integer})
		}
		return known(Value{Kind: widerInteger(left.Typed.Kind, right.Typed.Kind), Integer: left.Typed.Integer % right.Typed.Integer})
	}
	if isFloating(left.Typed) || isFloating(right.Typed) {
		lv, rv := numericFloat(left.Typed), numericFloat(right.Typed)
		var value float64
		switch op {
		case "+":
			value = lv + rv
		case "-":
			value = lv - rv
		case "*":
			value = lv * rv
		default:
			return unknown("unsupported arithmetic operator")
		}
		if math.IsInf(value, 0) || math.IsNaN(value) {
			return unknown("floating-point overflow")
		}
		return known(Value{Kind: widerFloat(left.Typed.Kind, right.Typed.Kind), Float: value})
	}
	if left.Typed.Kind == ValueCurrency || right.Typed.Kind == ValueCurrency {
		if left.Typed.Kind == ValueCurrency && right.Typed.Kind == ValueCurrency && (op == "+" || op == "-") {
			var scaled int64
			var ok bool
			if op == "+" {
				ok = checkedAdd(left.Typed.Currency, right.Typed.Currency, &scaled)
			} else {
				ok = checkedSub(left.Typed.Currency, right.Typed.Currency, &scaled)
			}
			if !ok {
				return unknown("currency overflow")
			}
			return known(Value{Kind: ValueCurrency, Currency: scaled})
		}
		value := numericFloat(left.Typed)
		rightValue := numericFloat(right.Typed)
		switch op {
		case "+":
			value += rightValue
		case "-":
			value -= rightValue
		case "*":
			value *= rightValue
		default:
			return unknown("unsupported currency operator")
		}
		if math.IsInf(value, 0) || math.IsNaN(value) {
			return unknown("currency overflow")
		}
		return known(Value{Kind: ValueDouble, Float: value})
	}
	var value int64
	var ok bool
	switch op {
	case "+":
		ok = checkedAdd(left.Typed.Integer, right.Typed.Integer, &value)
	case "-":
		ok = checkedSub(left.Typed.Integer, right.Typed.Integer, &value)
	case "*":
		ok = checkedMul(left.Typed.Integer, right.Typed.Integer, &value)
	default:
		return unknown("unsupported arithmetic operator")
	}
	if !ok {
		return unknown("integer overflow")
	}
	return known(Value{Kind: widerInteger(left.Typed.Kind, right.Typed.Kind), Integer: value})
}

func combineUnknown(left, right Result) Result {
	if left.Kind == Invalid {
		return left
	}
	if right.Kind == Invalid {
		return right
	}
	return unknown("unresolved operand")
}

func compare(left, right Result, op string) Result {
	if left.Kind != Known || right.Kind != Known {
		return combineUnknown(left, right)
	}
	if isRuntimeValue(left.Typed) || isRuntimeValue(right.Typed) {
		return unknown("comparison coercion is not modelled")
	}
	if op == "like" || op == "is" {
		return unknown("unsupported comparison operator")
	}
	if left.Typed.Kind == ValueString && right.Typed.Kind == ValueString {
		if op != "=" && op != "<>" {
			return unknown("string ordering depends on Option Compare")
		}
		value := left.Typed.String == right.Typed.String
		if op == "<>" {
			value = !value
		}
		return known(Value{Kind: ValueBoolean, Boolean: value})
	}
	if left.Typed.Kind == ValueBoolean && right.Typed.Kind == ValueBoolean {
		return compareFloat(boolFloat(left.Typed.Boolean), boolFloat(right.Typed.Boolean), op)
	}
	if !isNumeric(left.Typed) || !isNumeric(right.Typed) {
		return invalid("comparison requires compatible operands")
	}
	return compareFloat(numericFloat(left.Typed), numericFloat(right.Typed), op)
}

func compareFloat(left, right float64, op string) Result {
	var value bool
	switch op {
	case "=":
		value = left == right
	case "<>":
		value = left != right
	case "<":
		value = left < right
	case ">":
		value = left > right
	case "<=":
		value = left <= right
	case ">=":
		value = left >= right
	default:
		return unknown("unsupported comparison operator")
	}
	return known(Value{Kind: ValueBoolean, Boolean: value})
}

func boolFloat(value bool) float64 {
	if value {
		return -1
	}
	return 0
}
func numericFloat(value Value) float64 {
	if isIntegral(value) {
		return float64(value.Integer)
	}
	if value.Kind == ValueCurrency {
		return float64(value.Currency) / 10000
	}
	return value.Float
}
func isFloating(value Value) bool { return value.Kind == ValueSingle || value.Kind == ValueDouble }
func isNumeric(value Value) bool {
	return isIntegral(value) || isFloating(value) || value.Kind == ValueCurrency
}
func isRuntimeValue(value Value) bool {
	return value.Kind == ValueEmpty || value.Kind == ValueNull || value.Kind == ValueNothing
}
func widerInteger(left, right ValueKind) ValueKind {
	if left == ValueLongLong || right == ValueLongLong {
		return ValueLongLong
	}
	if left == ValueLong || right == ValueLong {
		return ValueLong
	}
	return ValueInteger
}
func widerFloat(left, right ValueKind) ValueKind {
	if left == ValueDouble || right == ValueDouble {
		return ValueDouble
	}
	return ValueSingle
}
func contains(values []string, value string) bool {
	for _, item := range values {
		if item == value {
			return true
		}
	}
	return false
}

func checkedAdd(a, b int64, out *int64) bool {
	if b > 0 && a > math.MaxInt64-b || b < 0 && a < math.MinInt64-b {
		return false
	}
	*out = a + b
	return true
}
func checkedSub(a, b int64, out *int64) bool {
	if b < 0 && a > math.MaxInt64+b || b > 0 && a < math.MinInt64+b {
		return false
	}
	*out = a - b
	return true
}
func checkedMul(a, b int64, out *int64) bool {
	if a == 0 || b == 0 {
		*out = 0
		return true
	}
	if a == math.MinInt64 && b == -1 || b == math.MinInt64 && a == -1 {
		return false
	}
	value := a * b
	if value/b != a {
		return false
	}
	*out = value
	return true
}
