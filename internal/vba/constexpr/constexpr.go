// Package constexpr evaluates the deliberately small integer constant subset
// shared by declaration and array-operation validation. It does not attempt
// to model VBA's full expression language: unresolved names and runtime
// expressions are explicitly represented as Unknown so callers can fail open.
package constexpr

import (
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

type Result struct {
	Kind   ResultKind
	Value  int
	Type   string
	Reason string
}

func EvaluateInteger(expression string, constants map[string]int) Result {
	text := strings.TrimSpace(expression)
	if text == "" {
		return Result{Kind: Unknown, Reason: "empty expression"}
	}
	// Calls and other runtime expressions are intentionally not interpreted as
	// malformed integer syntax; they are simply unresolved at analysis time.
	if looksLikeCall(text) {
		return Result{Kind: Unknown, Reason: "runtime expression"}
	}
	p := parser{text: text, constants: constants}
	result := p.expression()
	p.space()
	if result.Kind == Invalid {
		return result
	}
	if p.pos != len(p.text) {
		return Result{Kind: Invalid, Reason: "unexpected token"}
	}
	return result
}

type parser struct {
	text      string
	pos       int
	constants map[string]int
}

func (p *parser) expression() Result {
	left := p.term()
	for {
		p.space()
		if p.pos >= len(p.text) || (p.text[p.pos] != '+' && p.text[p.pos] != '-') {
			return left
		}
		op := p.text[p.pos]
		p.pos++
		right := p.term()
		left = combine(left, right, op)
	}
}

func (p *parser) term() Result {
	left := p.factor()
	for {
		p.space()
		if p.pos >= len(p.text) || (p.text[p.pos] != '*' && p.text[p.pos] != '/') {
			return left
		}
		op := p.text[p.pos]
		p.pos++
		right := p.factor()
		if op == '/' && right.Kind == Known && right.Value == 0 {
			left = Result{Kind: Invalid, Reason: "division by zero"}
			continue
		}
		left = combine(left, right, op)
	}
}

func (p *parser) factor() Result {
	p.space()
	if p.pos >= len(p.text) {
		return Result{Kind: Invalid, Reason: "missing operand"}
	}
	if p.text[p.pos] == '+' || p.text[p.pos] == '-' {
		op := p.text[p.pos]
		p.pos++
		result := p.factor()
		if result.Kind == Known && op == '-' {
			result.Value = -result.Value
		}
		return result
	}
	if p.text[p.pos] == '(' {
		p.pos++
		result := p.expression()
		p.space()
		if p.pos >= len(p.text) || p.text[p.pos] != ')' {
			return Result{Kind: Invalid, Reason: "unclosed parentheses"}
		}
		p.pos++
		return result
	}
	start := p.pos
	if unicode.IsDigit(rune(p.text[p.pos])) {
		for p.pos < len(p.text) && unicode.IsDigit(rune(p.text[p.pos])) {
			p.pos++
		}
		value, err := strconv.Atoi(p.text[start:p.pos])
		if err != nil {
			return Result{Kind: Invalid, Reason: "integer overflow"}
		}
		return Result{Kind: Known, Value: value, Type: "integer"}
	}
	if unicode.IsLetter(rune(p.text[p.pos])) || p.text[p.pos] == '_' {
		p.pos++
		for p.pos < len(p.text) && (unicode.IsLetter(rune(p.text[p.pos])) || unicode.IsDigit(rune(p.text[p.pos])) || p.text[p.pos] == '_') {
			p.pos++
		}
		name := strings.ToLower(p.text[start:p.pos])
		if value, ok := p.constants[name]; ok {
			return Result{Kind: Known, Value: value, Type: "integer"}
		}
		return Result{Kind: Unknown, Reason: "unresolved constant " + name}
	}
	return Result{Kind: Invalid, Reason: "unexpected token"}
}

func (p *parser) space() {
	for p.pos < len(p.text) && (p.text[p.pos] == ' ' || p.text[p.pos] == '\t') {
		p.pos++
	}
}

func combine(left, right Result, op byte) Result {
	if left.Kind == Invalid {
		return left
	}
	if right.Kind == Invalid {
		return right
	}
	if left.Kind == Unknown || right.Kind == Unknown {
		return Result{Kind: Unknown, Reason: "unresolved operand"}
	}
	value := left.Value
	switch op {
	case '+':
		value += right.Value
	case '-':
		value -= right.Value
	case '*':
		value *= right.Value
	case '/':
		if right.Value == 0 {
			return Result{Kind: Invalid, Reason: "division by zero"}
		}
		value /= right.Value
	}
	return Result{Kind: Known, Value: value, Type: "integer"}
}

func looksLikeCall(text string) bool {
	for i := 0; i < len(text); i++ {
		if !unicode.IsLetter(rune(text[i])) && text[i] != '_' {
			continue
		}
		j := i + 1
		for j < len(text) && (unicode.IsLetter(rune(text[j])) || unicode.IsDigit(rune(text[j])) || text[j] == '_') {
			j++
		}
		for j < len(text) && (text[j] == ' ' || text[j] == '\t') {
			j++
		}
		return j < len(text) && text[j] == '('
	}
	return false
}
