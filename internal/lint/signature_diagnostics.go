package lint

import (
	"context"
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// procedureSignature is deliberately a lint-local view.  The procedure IR is
// the long-term shared representation, but keeping this small CST projection
// here lets the declaration rules work with parsed editor buffers while the
// shared model evolves.
type procedureSignature struct {
	name       string
	accessor   string
	returnType string
	params     []signatureParameter
	rng        vbaast.Range
	nameRange  vbaast.Range
	path       []declarationBranch
	startByte  uint
	recovered  bool
}

type signatureParameter struct {
	name         string
	typ          string
	passing      string
	optional     bool
	paramArray   bool
	isArray      bool
	hasBounds    bool
	defaultValue string
	rng          vbaast.Range
	nameRange    vbaast.Range
	text         string
}

var identifierCallDefault = regexp.MustCompile(`(?i)^[A-Za-z_][A-Za-z0-9_\.]*\s*\(`)

func (l Linter) procedureSignatureIssuesContext(ctx context.Context, path string, source []byte, root *tree_sitter.Node, index declarationIndex) ([]Issue, error) {
	if root == nil {
		return nil, nil
	}
	sigs := collectProcedureSignatures(ctx, source, root)
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	udtCounts := make(map[string]int)
	if l.TypeDeclarations == nil {
		for _, record := range index.records {
			if record.kind == "type" && record.scope == "module" && len(record.path) == 0 {
				udtCounts[canonicalDeclarationKey(record.name)]++
			}
		}
	} else {
		for name, count := range l.TypeDeclarations {
			udtCounts[name] = count
		}
	}
	udts := udtCounts
	issues := make([]Issue, 0)
	for i := range sigs {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		issues = append(issues, l.validateProcedureSignature(path, sigs[i], udts, l.ObjectTypeDeclarations)...)
	}
	issues = append(issues, l.validatePropertySignatures(path, sigs, udts, l.ObjectTypeDeclarations)...)
	sort.SliceStable(issues, func(i, j int) bool {
		if issues[i].Line != issues[j].Line {
			return issues[i].Line < issues[j].Line
		}
		if issues[i].Column != issues[j].Column {
			return issues[i].Column < issues[j].Column
		}
		if issues[i].Code != issues[j].Code {
			return issues[i].Code < issues[j].Code
		}
		return issues[i].Message < issues[j].Message
	})
	return issues, nil
}

func (l Linter) validateProcedureSignature(path string, sig procedureSignature, udts, objects map[string]int) []Issue {
	if sig.recovered {
		return nil
	}
	issues := make([]Issue, 0)
	if len(sig.params) > 60 {
		issue := l.issueAt(path, sig.nameRange, "VB048", "error", fmt.Sprintf("Procedure %q has more than the VBA limit of 60 parameters.", sig.name))
		issue.Kind = "parameter_limit"
		issue.Symbol = sig.name
		issues = append(issues, issue)
	}
	optionalSeen := false
	for i := range sig.params {
		param := sig.params[i]
		isSetterValue := (sig.accessor == "let" || sig.accessor == "set") && i == len(sig.params)-1
		if optionalSeen && param.paramArray {
			issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "paramarray_after_optional", "ParamArray cannot follow an Optional parameter."))
		} else if optionalSeen && !param.optional && !isSetterValue {
			issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "required_after_optional", fmt.Sprintf("Required parameter %q cannot follow an Optional parameter.", param.name)))
		}
		if param.optional {
			optionalSeen = true
		}
		if param.paramArray {
			if i != len(sig.params)-1 {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "paramarray_not_last", "ParamArray must be the final parameter."))
			}
			if param.optional {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "paramarray_optional", "ParamArray cannot be Optional."))
			}
			if !param.isArray || !isVariantType(param.typ, udts, objects) {
				// ParamArray is always a Variant array.  This is a structural
				// check, so it remains reportable even when unrelated types are
				// unresolved.
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "paramarray_shape", "ParamArray must be declared as a Variant array."))
			}
		}
		if param.hasBounds {
			issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "fixed_array_parameter", "Procedure parameters must use a dynamic array form, such as values()."))
		}
		kind := classifySignatureType(param.typ, udts, objects)
		if kind == signatureTypeUDT {
			if strings.EqualFold(param.passing, "ByVal") {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "udt_byval", "User-defined type parameters cannot be passed ByVal."))
			}
			if param.optional {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "udt_optional", "User-defined type parameters cannot be Optional."))
			}
		}
		if param.optional && strings.TrimSpace(param.defaultValue) != "" {
			constant, definitelyNonconstant := classifyOptionalDefault(param.defaultValue)
			if definitelyNonconstant {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "optional_default_nonconstant", "Optional parameter defaults must be constant expressions."))
			} else if constant && optionalDefaultTypeMismatch(param.defaultValue, param.typ, kind) {
				issues = append(issues, signatureIssue(l, path, param.nameRange, sig.name, "optional_default_type", fmt.Sprintf("Optional default %q is incompatible with parameter type %s.", strings.TrimSpace(param.defaultValue), displaySignatureType(param.typ))))
			}
		}
	}
	return issues
}

func (l Linter) validatePropertySignatures(path string, sigs []procedureSignature, udts, objects map[string]int) []Issue {
	groups := make(map[string][]procedureSignature)
	for _, sig := range sigs {
		if sig.accessor == "" {
			continue
		}
		// Signatures collected from one parsed file necessarily share a module,
		// so grouping by canonical name needs no project-wide state. Mutually
		// exclusive conditional branches are separated later by
		// declarationBranchesOverlap when accessor pairs are compared.
		key := canonicalDeclarationKey(sig.name)
		groups[key] = append(groups[key], sig)
	}
	issues := make([]Issue, 0)
	for _, group := range groups {
		for i := range group {
			if group[i].recovered {
				continue
			}
			sig := group[i]
			if sig.accessor == "let" || sig.accessor == "set" {
				if len(sig.params) == 0 {
					accessorLabel := strings.ToUpper(sig.accessor[:1]) + sig.accessor[1:]
					issues = append(issues, signatureIssue(l, path, sig.nameRange, sig.name, "setter_value_missing", fmt.Sprintf("Property %s must declare a final value parameter.", accessorLabel)))
				} else {
					value := sig.params[len(sig.params)-1]
					kind := classifySignatureType(value.typ, udts, objects)
					if value.optional {
						issues = append(issues, signatureIssue(l, path, value.nameRange, sig.name, "setter_value_optional", "Property Let and Property Set value parameters cannot be Optional."))
					}
					if value.paramArray {
						issues = append(issues, signatureIssue(l, path, value.nameRange, sig.name, "setter_value_paramarray", "Property Let and Property Set value parameters cannot be ParamArray."))
					}
					if sig.accessor == "set" && kind != signatureTypeUnknown && kind != signatureTypeObject && kind != signatureTypeVariant {
						issues = append(issues, signatureIssue(l, path, value.nameRange, sig.name, "set_value_type", "Property Set value parameters must be object-compatible."))
					}
					if sig.accessor == "let" && kind == signatureTypeObject {
						issues = append(issues, signatureIssue(l, path, value.nameRange, sig.name, "let_value_type", "Property Let value parameters must be value-compatible."))
					}
				}
				if strings.TrimSpace(sig.returnType) != "" {
					issues = append(issues, signatureIssue(l, path, sig.nameRange, sig.name, "setter_return_type", "Property Let and Property Set declarations cannot declare a return type."))
				}
			}
		}
		for i := range group {
			if group[i].accessor != "get" || group[i].recovered {
				continue
			}
			for j := range group {
				if (group[j].accessor != "let" && group[j].accessor != "set") || group[j].recovered || !declarationBranchesOverlap(group[i].path, group[j].path) {
					continue
				}
				issues = append(issues, comparePropertyAccessorPair(l, path, group[i], group[j], udts, objects)...)
			}
		}
	}
	return issues
}

func comparePropertyAccessorPair(l Linter, path string, getter, setter procedureSignature, udts, objects map[string]int) []Issue {
	issues := make([]Issue, 0)
	if len(setter.params) == 0 {
		return issues
	}
	laterIsSetter := setter.startByte > getter.startByte
	laterNameRange, laterSymbol := getter.nameRange, getter.name
	if laterIsSetter {
		laterNameRange, laterSymbol = setter.nameRange, setter.name
	}
	indexParams := setter.params[:len(setter.params)-1]
	if len(getter.params) != len(indexParams) {
		issues = append(issues, signatureIssue(l, path, laterNameRange, laterSymbol, "property_index_count", "Property accessor index parameter counts must match."))
		return issues
	}
	for i := range getter.params {
		left, right := getter.params[i], indexParams[i]
		if left.isArray != right.isArray || left.hasBounds != right.hasBounds || !strings.EqualFold(left.passing, right.passing) || left.optional != right.optional {
			target := right
			if !laterIsSetter {
				target = left
			}
			issues = append(issues, signatureIssue(l, path, target.nameRange, laterSymbol, "property_index_shape", "Property accessor index parameters must have matching passing mode, array shape, and Optional modifier."))
			continue
		}
		leftKind := classifySignatureType(left.typ, udts, objects)
		rightKind := classifySignatureType(right.typ, udts, objects)
		if leftKind != signatureTypeUnknown && rightKind != signatureTypeUnknown && !signatureTypesEqual(left.typ, right.typ) {
			target := right
			if !laterIsSetter {
				target = left
			}
			issues = append(issues, signatureIssue(l, path, target.nameRange, laterSymbol, "property_index_type", "Property accessor index parameter types must match."))
		}
	}
	value := setter.params[len(setter.params)-1]
	getterKind := classifySignatureType(getter.returnType, udts, objects)
	valueKind := classifySignatureType(value.typ, udts, objects)
	if getterKind != signatureTypeUnknown && valueKind != signatureTypeUnknown && !propertyValueTypesCompatible(getter, setter, getterKind, valueKind) {
		target := value.nameRange
		if !laterIsSetter {
			target = getter.nameRange
		}
		issues = append(issues, signatureIssue(l, path, target, laterSymbol, "property_value_type", "Property accessor value type must match the Property Get return type."))
	}
	return issues
}

func signatureTypesEqual(left, right string) bool {
	leftName := cleanIdentifier(strings.TrimSpace(left))
	rightName := cleanIdentifier(strings.TrimSpace(right))
	if leftName == "" {
		leftName = "variant"
	}
	if rightName == "" {
		rightName = "variant"
	}
	return strings.EqualFold(leftName, rightName)
}

func propertyValueTypesCompatible(getter, setter procedureSignature, getterKind, valueKind signatureType) bool {
	// An omitted `As` clause and an explicit Variant are the same effective
	// property value type.  Property Set also permits the normal Variant/Object
	// late-bound pair used by collection-like properties.
	if getterKind == signatureTypeVariant && valueKind == signatureTypeVariant {
		return true
	}
	if setter.accessor == "set" &&
		(getterKind == signatureTypeVariant || getterKind == signatureTypeObject) &&
		(valueKind == signatureTypeVariant || valueKind == signatureTypeObject) {
		return true
	}
	return strings.EqualFold(cleanIdentifier(getter.returnType), cleanIdentifier(setter.params[len(setter.params)-1].typ))
}

func signatureIssue(l Linter, path string, r vbaast.Range, symbol, kind, message string) Issue {
	issue := l.issueAt(path, r, "VB048", "error", message)
	issue.Kind = kind
	issue.Symbol = symbol
	if strings.HasPrefix(kind, "property_") || strings.HasPrefix(kind, "setter_") || strings.HasPrefix(kind, "let_") || strings.HasPrefix(kind, "set_") {
		issue.Code = "VB049"
	}
	return issue
}

func collectProcedureSignatures(ctx context.Context, source []byte, root *tree_sitter.Node) []procedureSignature {
	collector := signatureCollector{ctx: ctx, source: source}
	collector.walk(root, declarationWalkState{scope: "module", module: true})
	sort.SliceStable(collector.signatures, func(i, j int) bool { return collector.signatures[i].startByte < collector.signatures[j].startByte })
	return collector.signatures
}

type signatureCollector struct {
	ctx        context.Context
	source     []byte
	signatures []procedureSignature
}

func (c *signatureCollector) walk(node *tree_sitter.Node, state declarationWalkState) {
	if node == nil || c.ctx.Err() != nil {
		return
	}
	if node.Kind() == "preprocessor_if" || node.Kind() == "type_preprocessor_if" {
		// Reuse the declaration collector's branch semantics.
		c.walkConditional(node, state)
		return
	}
	if isSignatureProcedureDeclarationKind(node.Kind()) {
		header := procedureHeaderNode(node)
		if header != nil {
			sig := buildProcedureSignature(header, c.source)
			sig.path = append([]declarationBranch(nil), state.path...)
			c.signatures = append(c.signatures, sig)
		}
		if strings.HasPrefix(node.Kind(), "conditional_") {
			return
		}
		state = declarationWalkState{scope: "proc:" + fmt.Sprint(node.StartByte()), parent: sigName(header, c.source), procedure: true, path: state.path}
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		c.walk(node.NamedChild(i), state)
	}
}

func isSignatureProcedureDeclarationKind(kind string) bool {
	if isProcedureDeclarationKind(kind) {
		return true
	}
	switch kind {
	case "declare_statement", "declare_sub_statement", "declare_function_statement":
		return true
	default:
		return false
	}
}

func (c *signatureCollector) walkConditional(node *tree_sitter.Node, state declarationWalkState) {
	group := fmt.Sprint(node.StartByte())
	condition := conditionalExpressionKey(node, c.source)
	branch := 0
	if body := node.ChildByFieldName("body"); body != nil {
		c.walk(body, declarationWalkState{scope: state.scope, parent: state.parent, module: state.module, procedure: state.procedure, path: appendBranch(state.path, group, condition, branch)})
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() == "newline" || (!strings.Contains(child.Kind(), "else") && !strings.Contains(child.Kind(), "elseif")) {
			continue
		}
		branch++
		if body := child.ChildByFieldName("body"); body != nil {
			c.walk(body, declarationWalkState{scope: state.scope, parent: state.parent, module: state.module, procedure: state.procedure, path: appendBranch(state.path, group, condition, branch)})
		}
	}
}

func procedureHeaderNode(node *tree_sitter.Node) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	if strings.HasPrefix(node.Kind(), "conditional_") {
		return node.ChildByFieldName("consequence")
	}
	return node
}

func sigName(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	return cleanIdentifier(declarationNodeName(node, source))
}

func buildProcedureSignature(node *tree_sitter.Node, source []byte) procedureSignature {
	sig := procedureSignature{rng: vbaast.NodeRange(node), startByte: node.StartByte(), recovered: node.HasError() || node.IsError() || node.IsMissing()}
	sig.name = sigName(node, source)
	if name := declarationNameNode(node); name != nil {
		sig.nameRange = vbaast.NodeRange(name)
	} else {
		sig.nameRange = sig.rng
	}
	sig.accessor = procedureAccessor(node.Kind(), node.Utf8Text(source))
	sig.returnType = strings.TrimSpace(typeText(node, source))
	if (sig.accessor == "let" || sig.accessor == "set") && node.ChildByFieldName("type") == nil && firstNamedChildKind(node, "as_type_clause") == nil {
		// typeText falls back to an identifier's type-declaration character;
		// setter declarations have no return slot, so that fallback is not a
		// return declaration here.
		sig.returnType = ""
	}
	list := node.ChildByFieldName("parameters")
	if list == nil {
		list = firstNamedChildKind(node, "parameter_list")
	}
	if list == nil {
		return sig
	}
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil || child.Kind() != "parameter" {
			continue
		}
		param := signatureParameter{text: child.Utf8Text(source), typ: strings.TrimSpace(typeText(child, source)), passing: "ByRef", rng: vbaast.NodeRange(child)}
		if name := declarationNameNode(child); name != nil {
			param.name = cleanIdentifier(name.Utf8Text(source))
			param.nameRange = vbaast.NodeRange(name)
		} else {
			param.nameRange = param.rng
		}
		if mode := child.ChildByFieldName("passing_mode"); mode != nil {
			param.passing = normalizeKeyword(mode.Utf8Text(source))
		} else if hasWord(param.text, "ByVal") {
			param.passing = "ByVal"
		} else if hasWord(param.text, "ByRef") {
			param.passing = "ByRef"
		}
		param.optional = child.ChildByFieldName("optional_modifier") != nil || hasWord(param.text, "Optional")
		param.paramArray = child.ChildByFieldName("paramarray_modifier") != nil || hasWord(param.text, "ParamArray")
		if value := child.ChildByFieldName("default_value"); value != nil {
			param.defaultValue = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(value.Utf8Text(source)), "="))
		} else if initializer := firstNamedChildKind(child, "initializer"); initializer != nil {
			param.defaultValue = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(initializer.Utf8Text(source)), "="))
		}
		if name := declarationNameNode(child); name != nil {
			start, end := int(name.EndByte()), int(child.EndByte())
			if start >= 0 && end >= start && end <= len(source) {
				tail := source[start:end]
				trimmed := strings.TrimSpace(string(tail))
				if strings.HasPrefix(trimmed, "(") {
					param.isArray = true
					close := strings.IndexByte(trimmed, ')')
					if close > 1 && strings.TrimSpace(trimmed[1:close]) != "" {
						param.hasBounds = true
					}
				}
			}
		}
		for j := uint(0); j < child.NamedChildCount(); j++ {
			named := child.NamedChild(j)
			if named != nil && named.Kind() == "array_bounds" {
				// The grammar represents both `()` and explicit bounds with an
				// array_bounds node.  Only the latter is a fixed/bounded form;
				// an empty pair is the legal dynamic-array spelling.
				boundsText := strings.TrimSpace(named.Utf8Text(source))
				if open := strings.IndexByte(boundsText, '('); open >= 0 {
					close := strings.LastIndexByte(boundsText, ')')
					if close < open+1 {
						param.hasBounds = false
					} else {
						param.hasBounds = strings.TrimSpace(boundsText[open+1:close]) != ""
					}
				} else {
					param.hasBounds = boundsText != ""
				}
			}
		}
		if param.typ == "" && param.isArray {
			param.typ = "Variant"
		}
		sig.params = append(sig.params, param)
	}
	return sig
}

type signatureType int

const (
	signatureTypeUnknown signatureType = iota
	signatureTypeIntrinsic
	signatureTypeUDT
	signatureTypeObject
	signatureTypeVariant
)

func classifySignatureType(typ string, udts, objects map[string]int) signatureType {
	name := strings.ToLower(cleanIdentifier(strings.TrimSpace(typ)))
	if name == "" || name == "variant" {
		return signatureTypeVariant
	}
	if strings.Contains(name, ".") {
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 && udts[name[dot+1:]] == 1 {
			return signatureTypeUDT
		}
		return signatureTypeObject
	}
	if objects[name] == 1 {
		return signatureTypeObject
	}
	if objects[name] > 1 {
		return signatureTypeUnknown
	}
	if name == "object" || name == "collection" || name == "dictionary" || name == "range" || name == "worksheet" || name == "workbook" || name == "application" {
		return signatureTypeObject
	}
	if udts[name] == 1 {
		return signatureTypeUDT
	}
	switch name {
	case "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "single", "string":
		return signatureTypeIntrinsic
	default:
		return signatureTypeUnknown
	}
}

func isVariantType(typ string, udts, objects map[string]int) bool {
	return classifySignatureType(typ, udts, objects) == signatureTypeVariant
}

func displaySignatureType(typ string) string {
	if strings.TrimSpace(typ) == "" {
		return "Variant"
	}
	return strings.TrimSpace(typ)
}

// classifyOptionalDefault returns whether the expression is a known constant
// and whether it is definitely non-constant.  Expressions such as Const or
// Enum references and arithmetic are intentionally unknown here; #595's
// shared evaluator will own those cases, so this rule must remain fail-open.
func classifyOptionalDefault(value string) (constant, definitelyNonconstant bool) {
	value = strings.TrimSpace(value)
	for strings.HasPrefix(value, "(") && strings.HasSuffix(value, ")") && balancedParens(value) {
		value = strings.TrimSpace(value[1 : len(value)-1])
	}
	if value == "" {
		return false, false
	}
	if identifierCallDefault.MatchString(value) || strings.HasPrefix(strings.ToLower(value), "new ") {
		return false, true
	}
	if strings.EqualFold(value, "true") || strings.EqualFold(value, "false") || strings.EqualFold(value, "empty") || strings.EqualFold(value, "null") || strings.EqualFold(value, "nothing") {
		return true, false
	}
	if strings.HasPrefix(value, `"`) && strings.HasSuffix(value, `"`) {
		return true, false
	}
	if strings.HasPrefix(value, "#") && strings.HasSuffix(value, "#") {
		return true, false
	}
	if isNumericLiteral(value) {
		return true, false
	}
	if strings.HasPrefix(value, "+") || strings.HasPrefix(value, "-") {
		return classifyOptionalDefault(strings.TrimSpace(value[1:]))
	}
	return false, false
}

func optionalDefaultTypeMismatch(value, typ string, kind signatureType) bool {
	if kind == signatureTypeVariant || kind == signatureTypeUnknown || kind == signatureTypeUDT {
		return false
	}
	value = strings.TrimSpace(value)
	for strings.HasPrefix(value, "(") && strings.HasSuffix(value, ")") && balancedParens(value) {
		value = strings.TrimSpace(value[1 : len(value)-1])
	}
	if strings.HasPrefix(value, `"`) && strings.HasSuffix(value, `"`) {
		return kind == signatureTypeIntrinsic && !strings.EqualFold(cleanIdentifier(typ), "string")
	}
	if strings.EqualFold(value, "true") || strings.EqualFold(value, "false") {
		return !strings.EqualFold(cleanIdentifier(typ), "boolean")
	}
	if strings.HasPrefix(value, "#") && strings.HasSuffix(value, "#") {
		return !strings.EqualFold(cleanIdentifier(typ), "date")
	}
	if isNumericLiteral(value) {
		target := strings.ToLower(cleanIdentifier(typ))
		return target == "string" || target == "boolean"
	}
	return false
}

var numericLiteralPattern = regexp.MustCompile(`^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$`)

func isNumericLiteral(value string) bool {
	value = strings.TrimSpace(value)
	value = strings.TrimPrefix(strings.TrimPrefix(value, "+"), "-")
	if !numericLiteralPattern.MatchString(value) {
		return false
	}
	_, err := strconv.ParseFloat(value, 64)
	return err == nil
}

func balancedParens(value string) bool {
	depth := 0
	for _, r := range value {
		switch r {
		case '(':
			depth++
		case ')':
			depth--
			if depth < 0 {
				return false
			}
		}
	}
	return depth == 0
}
