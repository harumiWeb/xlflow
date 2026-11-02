// Package conditional selects provably active VBA conditional-compilation
// source without changing source byte positions. Unknown conditions retain
// every possible branch; this package never contacts Excel.
package conditional

import (
	"bytes"
	"fmt"
	"maps"
	"slices"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
)

type truth uint8

const (
	no truth = iota
	yes
	maybe
)

func (t truth) not() truth {
	if t == maybe {
		return maybe
	}
	if t == yes {
		return no
	}
	return yes
}

func (t truth) and(other truth) truth {
	if t == no || other == no {
		return no
	}
	if t == maybe || other == maybe {
		return maybe
	}
	return yes
}

// An absent ordinary identifier is VBA Empty, whereas a present Unknown
// value and unspecified platform constants must stay unresolved.
type environment map[string]constexpr.Result

func (e environment) Resolve(name string) (constexpr.Value, bool) {
	name = strings.ToLower(name)
	if value, ok := e[name]; ok {
		return value.Typed, value.Kind == constexpr.Known
	}
	switch name {
	case "vba6", "vba7", "win16", "win32", "win64", "mac":
		return constexpr.Value{}, false
	}
	// Empty coerces to zero in conditional expressions. Keep this coercion
	// local instead of changing the general constant evaluator's semantics.
	return constexpr.Value{Kind: constexpr.ValueLong, Integer: 0}, true
}

func evaluate(text string, env environment) constexpr.Result {
	return constexpr.EvaluateConditional(text, env)
}

func condition(text string, env environment) truth {
	result := evaluate(text, env)
	if result.Kind != constexpr.Known {
		return maybe
	}
	switch v := result.Typed; v.Kind {
	case constexpr.ValueBoolean:
		if v.Boolean {
			return yes
		}
		return no
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong:
		if v.Integer != 0 {
			return yes
		}
		return no
	case constexpr.ValueSingle, constexpr.ValueDouble:
		if v.Float != 0 {
			return yes
		}
		return no
	case constexpr.ValueCurrency:
		if v.Currency != 0 {
			return yes
		}
		return no
	case constexpr.ValueEmpty:
		return no
	}
	return maybe
}

// ValidateConstants validates the case-insensitive project constant table.
// Values use VBA constant-expression syntax, encoded as TOML strings. Unknown
// expressions are rejected in configuration rather than silently becoming 0.
func ValidateConstants(constants map[string]string) error {
	_, err := projectEnvironment(constants)
	return err
}

func projectEnvironment(constants map[string]string) (environment, error) {
	env := make(environment, len(constants))
	for _, name := range slices.Sorted(maps.Keys(constants)) {
		key := strings.ToLower(name)
		if !identifier(name) {
			return nil, fmt.Errorf("invalid conditional constant name %q", name)
		}
		if _, exists := env[key]; exists {
			return nil, fmt.Errorf("duplicate conditional constant %q", name)
		}
		// Project arguments are an unordered table, not sequential module
		// directives: each value must stand on its own.
		value := constexpr.EvaluateConditional(constants[name], nil)
		if value.Kind != constexpr.Known {
			return nil, fmt.Errorf("conditional constant %q must be a supported VBA literal expression", name)
		}
		switch value.Typed.Kind {
		case constexpr.ValueNull, constexpr.ValueNothing, constexpr.ValueDate:
			return nil, fmt.Errorf("unsupported conditional constant value for %q", name)
		}
		if value.Typed.Kind == constexpr.ValueEmpty {
			value = evaluate("0", nil)
		}
		env[key] = value
	}
	return env, nil
}

type frame struct {
	parent, remaining, active truth
	before                    environment
	branches                  []environment
	elseSeen                  bool
	uncertain                 bool
	directives                [][2]int
}

// Project masks provably inactive lines and resolved conditional wrappers.
// Line endings and byte offsets are preserved, and input is never modified.
// Unknown wrappers remain for existing IR uncertainty and structure checks.
// Malformed/unclosed directives return the original source conservatively.
func Project(source []byte, constants map[string]string) []byte {
	if !bytes.Contains(source, []byte("#")) || !utf8.Valid(source) {
		return source
	}
	env, err := projectEnvironment(constants)
	if err != nil {
		return source
	} // configuration loaders report this error
	var stack []frame
	active := yes
	var output []byte
	mask := func(start, end int) {
		if output == nil {
			output = bytes.Clone(source)
		}
		for i := start; i < end; i++ {
			if output[i] != '\r' && output[i] != '\n' {
				output[i] = ' '
			}
		}
	}
	for start := 0; start < len(source); {
		end := start
		for end < len(source) && source[end] != '\n' && source[end] != '\r' {
			end++
		}
		next := end
		if next < len(source) && source[next] == '\r' {
			next++
		}
		if next < len(source) && source[next] == '\n' {
			next++
		}
		trimmed := bytes.TrimSpace(source[start:end])
		if len(trimmed) == 0 || trimmed[0] != '#' {
			if active == no {
				mask(start, end)
			}
			start = next
			continue
		}
		line := stripComment(string(trimmed))
		kind, argument := directive(line)
		switch kind {
		case "if":
			expr, ok := conditionExpression(argument)
			if !ok {
				return source
			}
			value := condition(expr, env)
			stack = append(stack, frame{parent: active, remaining: value.not(), active: active.and(value), before: maps.Clone(env), uncertain: value == maybe, directives: [][2]int{{start, end}}})
			active = active.and(value)
		case "elseif", "else":
			if len(stack) == 0 || stack[len(stack)-1].elseSeen {
				return source
			}
			f := &stack[len(stack)-1]
			f.directives = append(f.directives, [2]int{start, end})
			if f.active != no {
				f.branches = append(f.branches, env)
			}
			env = maps.Clone(f.before)
			value := yes
			if kind == "elseif" {
				expr, ok := conditionExpression(argument)
				if !ok {
					return source
				}
				value = condition(expr, env)
			} else {
				if argument != "" {
					return source
				}
				f.elseSeen = true
			}
			f.active = f.parent.and(f.remaining).and(value)
			if f.remaining != no && value == maybe {
				f.uncertain = true
			}
			f.remaining = f.remaining.and(value.not())
			active = f.active
		case "end":
			if len(stack) == 0 || !strings.EqualFold(argument, "if") {
				return source
			}
			f := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			if f.active != no {
				f.branches = append(f.branches, env)
			}
			if f.remaining != no || f.parent != yes {
				f.branches = append(f.branches, f.before)
			}
			env = merge(f.before, f.branches)
			if !f.uncertain || f.parent == no {
				for _, span := range f.directives {
					mask(span[0], span[1])
				}
				mask(start, end)
			}
			active = f.parent
		case "const":
			if active == no {
				mask(start, end)
				break
			}
			name, expr, ok := strings.Cut(argument, "=")
			name = strings.ToLower(strings.TrimSpace(name))
			if !ok || !identifier(name) || strings.TrimSpace(expr) == "" {
				return source
			}
			if strings.HasSuffix(strings.TrimSpace(expr), " _") || strings.HasSuffix(strings.TrimSpace(expr), "\t_") {
				env[name] = constexpr.Result{Kind: constexpr.Unknown}
			} else {
				env[name] = evaluate(expr, env)
			}
			if env[name].Typed.Kind == constexpr.ValueEmpty {
				env[name] = evaluate("0", nil)
			}
		default:
			if active == no {
				mask(start, end)
			}
		}
		start = next
	}
	if len(stack) != 0 || output == nil {
		return source
	}
	return output
}

func merge(before environment, branches []environment) environment {
	if len(branches) == 0 {
		return before
	}
	result := maps.Clone(branches[0])
	for _, branch := range branches[1:] {
		for name, value := range result {
			other, ok := branch[name]
			if !ok {
				other = evaluate(name, branch)
			}
			if value != other {
				result[name] = constexpr.Result{Kind: constexpr.Unknown}
			}
		}
		for name, value := range branch {
			if _, ok := result[name]; !ok {
				prior := evaluate(name, result)
				if prior != value {
					result[name] = constexpr.Result{Kind: constexpr.Unknown}
				} else {
					result[name] = value
				}
			}
		}
	}
	return result
}

func identifier(name string) bool {
	if name == "" {
		return false
	}
	for i, r := range name {
		if !unicode.IsLetter(r) && r != '_' && (i == 0 || !unicode.IsDigit(r)) {
			return false
		}
	}
	return !reservedIdentifiers[strings.ToLower(name)]
}

// MS-VBAL 3.3.5.2 defines IDENTIFIER by excluding these reserved tokens.
// Keep VBA's list separate from the broader VB.NET rename restrictions.
// https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7df907cb-ab6c-40d3-aa81-272742ce00c3
var reservedIdentifiers = func() map[string]bool {
	const names = `call case close const declare defbool defbyte defcur defdate defdbl defint deflng deflnglng deflngptr defobj defsng defstr defvar
dim do else elseif end endif enum erase event exit for friend function get global gosub goto if implements input let lock loop lset next on open option print private public put raiseevent redim resume return rset seek select set static stop sub type unlock wend while with write
rem any as byref byval each in new shared until withevents optional paramarray preserve spc tab then to
addressof and eqv imp is like mod not or typeof xor
abs cbool cbyte ccur cdate cdbl cdec cint clng clnglng clngptr csng cstr cvar cverr date debug doevents fix int len lenb me pset scale sgn string
array circle inputb lbound ubound boolean byte currency double integer long longlong longptr single variant
true false nothing empty null
attribute lineinput vb_base vb_control vb_creatable vb_customizable vb_description vb_exposed vb_ext_key vb_globalnamespace vb_helpid vb_invoke_func vb_invoke_property vb_invoke_propertyput vb_invoke_propertyputref vb_memberflags vb_name vb_predeclaredid vb_procdata vb_templatederived vb_usermemid vb_vardescription vb_varhelpid vb_varmemberflags vb_varprocdata vb_varusermemid
cdecl decimal defdec`
	reserved := make(map[string]bool)
	for name := range strings.FieldsSeq(names) {
		reserved[name] = true
	}
	return reserved
}()

func stripComment(text string) string {
	quoted := false
	for i := 0; i < len(text); i++ {
		if text[i] == '"' {
			if quoted && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			quoted = !quoted
		} else if text[i] == '\'' && !quoted {
			return strings.TrimSpace(text[:i])
		}
	}
	return strings.TrimSpace(text)
}

func directive(line string) (string, string) {
	if !strings.HasPrefix(line, "#") {
		return "", ""
	}
	text := strings.TrimSpace(line[1:])
	i := 0
	for i < len(text) && ((text[i] >= 'A' && text[i] <= 'Z') || (text[i] >= 'a' && text[i] <= 'z')) {
		i++
	}
	kind := strings.ToLower(text[:i])
	if i < len(text) && text[i] != ' ' && text[i] != '\t' {
		return "", ""
	}
	if kind == "endif" {
		return "end", strings.TrimSpace("if " + text[i:])
	}
	return kind, strings.TrimSpace(text[i:])
}

func conditionExpression(argument string) (string, bool) {
	if len(argument) < 5 || !strings.EqualFold(argument[len(argument)-4:], "then") {
		return "", false
	}
	prefix := argument[:len(argument)-4]
	if len(prefix) == 0 || (prefix[len(prefix)-1] != ' ' && prefix[len(prefix)-1] != '\t') {
		return "", false
	}
	return strings.TrimSpace(prefix), strings.TrimSpace(prefix) != ""
}
