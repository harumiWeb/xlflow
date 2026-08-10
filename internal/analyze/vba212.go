package analyze

// The VBA212 scanner intentionally works on the CST rather than on source
// text.  VBA evaluates Boolean operands eagerly, so the relationship between
// a guard and the access it is meant to protect is an AST relationship.

import (
	"context"
	"strings"
	"unicode/utf16"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type vba212Context struct {
	projectEffects effects.ProjectSummary
	userDefined    map[string]bool
	getterEffects  map[string]bool
}

type vba212Guard struct {
	path  string
	name  string
	array bool
}

type vba212Unsafe struct {
	path   string
	text   string
	node   *tree_sitter.Node
	getter bool
}

type vba212OperandFacts struct {
	guards []vba212Guard
	unsafe []vba212Unsafe
}

type vba212Container struct {
	kind     string
	operands []*tree_sitter.Node
}

func (a Analyzer) vba212ScanWithContext(ctx context.Context, file parsedFile, procedures []sourceProcedure, stats *vba212ScanStats, scanCtx vba212Context) ([]Finding, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if len(procedures) == 0 {
		procedures = []sourceProcedure{{StartLine: 1, EndLine: len(file.Lines), StartByte: 0, EndByte: len(file.Source)}}
	}
	procedures = append([]sourceProcedure(nil), procedures...)
	sortSourceProcedures(procedures)
	summaries := scanCtx.projectEffects.All()
	if len(summaries) == 0 && file.Root != nil && vba212SourceMayHaveGetter(file) {
		// Standalone/realtime callers do not have a project summary.  Building a
		// one-document summary gives them the same-document getter proof while
		// retaining the batch caller's full-project summary when supplied.
		scanCtx.projectEffects = buildProjectEffects([]parsedFile{file})
		summaries = scanCtx.projectEffects.All()
	}
	if scanCtx.userDefined == nil {
		scanCtx.userDefined = map[string]bool{}
		for _, summary := range summaries {
			name := strings.ToLower(summary.Identity.Name)
			if name == "iif" || name == "choose" || name == "switch" {
				scanCtx.userDefined[name] = true
			}
		}
	}
	for _, procedure := range file.IR.Procedures {
		name := strings.ToLower(procedure.Symbol.Name)
		if name == "iif" || name == "choose" || name == "switch" {
			scanCtx.userDefined[name] = true
		}
	}
	for _, line := range file.Lines {
		lower := strings.ToLower(strings.TrimSpace(line))
		for _, name := range []string{"iif", "choose", "switch"} {
			if strings.Contains(lower, "function "+name) || strings.Contains(lower, "sub "+name) || strings.Contains(lower, "property get "+name) {
				scanCtx.userDefined[name] = true
			}
		}
	}
	if scanCtx.getterEffects == nil {
		scanCtx.getterEffects = map[string]bool{}
		getterCounts := map[string]int{}
		getterHasEffect := map[string]bool{}
		for _, summary := range summaries {
			if summary.Identity.Kind != procedureir.ProcedurePropertyGet {
				continue
			}
			key := strings.ToLower(summary.Identity.Module) + "\x00" + strings.ToLower(summary.Identity.Name)
			getterCounts[key]++
			for _, kind := range []effects.EffectKind{effects.WritesCells, effects.ChangesWorkbook, effects.OpensWorkbook, effects.ClosesWorkbook, effects.DisablesEvents, effects.RestoresEvents, effects.ChangesCalculation, effects.Recalculates, effects.ChangesSelection, effects.ChangesControls, effects.ShowsDialog, effects.LaunchesProcess, effects.SuppressesErrors, effects.RaisesError, effects.ChangesApplicationState, effects.RestoresApplicationState} {
				if summary.Has(kind) {
					getterHasEffect[key] = true
					break
				}
			}
		}
		for key, count := range getterCounts {
			if count == 1 && getterHasEffect[key] {
				scanCtx.getterEffects[key] = true
			}
		}
	}
	seen := map[string]bool{}
	factCache := map[int]vba212OperandFacts{}
	var findings []Finding
	if stats != nil {
		stats.RootTraversals++
	}
	visited := 0
	var visit func(*tree_sitter.Node) error
	visit = func(node *tree_sitter.Node) error {
		if node == nil {
			return nil
		}
		visited++
		if stats != nil {
			stats.NodesVisited = visited
		}
		if visited&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return err
			}
		}
		proc, owned := vba212ProcedureAtByte(procedures, int(node.StartByte()), int(node.EndByte()))
		if owned {
			if container, ok := vba212ContainerForNode(node, file.Source, scanCtx.userDefined); ok {
				findings = append(findings, a.vba212FindingsForContainer(file, proc, container, scanCtx, seen, factCache)...)
			}
		}
		for i := uint(0); i < node.NamedChildCount(); i++ {
			if err := visit(node.NamedChild(i)); err != nil {
				return err
			}
		}
		return nil
	}
	if err := visit(file.Root); err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	return findings, nil
}

func vba212SourceMayHaveGetter(file parsedFile) bool {
	for _, line := range file.Lines {
		if strings.Contains(strings.ToLower(strings.TrimSpace(line)), "property get ") {
			return true
		}
	}
	return false
}

func sortSourceProcedures(procedures []sourceProcedure) {
	// Kept local to avoid changing the existing scanner's instrumentation.
	for i := 1; i < len(procedures); i++ {
		for j := i; j > 0 && procedures[j].StartByte < procedures[j-1].StartByte; j-- {
			procedures[j], procedures[j-1] = procedures[j-1], procedures[j]
		}
	}
}

func vba212ContainerForNode(node *tree_sitter.Node, source []byte, userDefined map[string]bool) (vba212Container, bool) {
	if node == nil {
		return vba212Container{}, false
	}
	switch node.Kind() {
	case "condition_binary_expression", "binary_expression":
		if node.ChildCount() < 2 {
			return vba212Container{}, false
		}
		left, right := node.Child(0), node.Child(1)
		if left == nil || right == nil || right.StartByte() < left.EndByte() {
			return vba212Container{}, false
		}
		gap := string(source[left.EndByte():right.StartByte()])
		op := vba212OperatorFromGap(gap)
		if op != "and" && op != "or" {
			return vba212Container{}, false
		}
		return vba212Container{kind: strings.ToUpper(op), operands: []*tree_sitter.Node{left, right}}, true
	case "call_expression":
		fn := node.ChildByFieldName("function")
		if fn == nil || fn.Kind() != "identifier" {
			return vba212Container{}, false
		}
		name := strings.ToLower(cleanIdentifier(fn.Utf8Text(source)))
		if name != "iif" && name != "choose" && name != "switch" {
			return vba212Container{}, false
		}
		if userDefined[name] {
			return vba212Container{}, false
		}
		args := node.ChildByFieldName("arguments")
		if args == nil {
			return vba212Container{}, false
		}
		operands := make([]*tree_sitter.Node, 0, args.NamedChildCount())
		for i := uint(0); i < args.NamedChildCount(); i++ {
			operands = append(operands, args.NamedChild(i))
		}
		kind := map[string]string{"iif": "IIf", "choose": "Choose", "switch": "Switch"}[name]
		return vba212Container{kind: kind, operands: operands}, len(operands) > 1
	}
	return vba212Container{}, false
}

func vba212OperatorFromGap(gap string) string {
	var b strings.Builder
	inString := false
	for i := 0; i < len(gap); i++ {
		ch := gap[i]
		if ch == '"' {
			inString = !inString
			continue
		}
		if !inString && ch == '\'' {
			for i < len(gap) && gap[i] != '\n' {
				i++
			}
			continue
		}
		if !inString {
			b.WriteByte(ch)
		}
	}
	clean := strings.ToLower(b.String())
	if hasWord(clean, "And") {
		return "and"
	}
	if hasWord(clean, "Or") {
		return "or"
	}
	return ""
}

func (a Analyzer) vba212FindingsForContainer(file parsedFile, proc sourceProcedure, container vba212Container, scanCtx vba212Context, seen map[string]bool, factCache map[int]vba212OperandFacts) []Finding {
	facts := make([]vba212OperandFacts, len(container.operands))
	for i, operand := range container.operands {
		key := int(operand.StartByte())
		if cached, ok := factCache[key]; ok {
			facts[i] = cached
		} else {
			facts[i] = vba212Facts(operand, file.Source, file, proc, scanCtx)
			factCache[key] = facts[i]
		}
	}
	var out []Finding
	for i, left := range facts {
		for j, right := range facts {
			if i == j {
				continue
			}
			for _, guard := range left.guards {
				for _, unsafe := range right.unsafe {
					if !vba212PathMatches(guard.path, unsafe.path) {
						continue
					}
					out = append(out, a.vba212Finding(file, proc, container, guard, unsafe, seen))
				}
			}
		}
	}
	// A resolved side-effecting getter is unsafe even without a Nothing guard,
	// but only in an eager container.  Unresolved/ambiguous getters never enter
	// the facts list.
	for _, fact := range facts {
		for _, unsafe := range fact.unsafe {
			if !unsafe.getter {
				continue
			}
			guard := vba212Guard{name: "property getter", path: ""}
			out = append(out, a.vba212Finding(file, proc, container, guard, unsafe, seen))
		}
	}
	filtered := out[:0]
	for _, finding := range out {
		if finding.Code != "" {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func (a Analyzer) vba212Finding(file parsedFile, proc sourceProcedure, container vba212Container, guard vba212Guard, unsafe vba212Unsafe, seen map[string]bool) Finding {
	r := vbaast.NodeRange(unsafe.node)
	guardLabel := guard.name
	if guardLabel == "" {
		guardLabel = "the guarded object"
	}
	key := strconvItoa(proc.StartByte) + ":" + strconvItoa(r.StartByte) + ":" + strings.ToLower(guard.path)
	if seen[key] {
		return Finding{}
	}
	seen[key] = true
	message := unsafe.text + " dereferences " + guardLabel + " in the same non-short-circuit boolean expression (" + strings.ToUpper(container.kind) + ")."
	reason := "VBA And/Or, IIf, Choose, and Switch expressions evaluate every operand eagerly, so this access can execute before the guard takes effect and raise runtime error 91."
	suggestion := vba212Suggestion(container.kind, guardLabel, unsafe.text)
	if unsafe.getter && guard.path == "" {
		message = unsafe.text + " invokes a side-effecting property getter in an eager " + strings.ToLower(container.kind) + " expression."
		reason = "The getter has a resolved VBA effect and is evaluated even when another operand would appear to guard it."
		suggestion = "Read " + unsafe.text + " once into a local variable before the Boolean expression, then use nested If statements."
	}
	finding := a.simpleFinding(file, proc, r.StartLine, "VBA212", "warning", message, reason, suggestion)
	finding.Column = vba212UTF16Column(file.Source, r.StartLine, r.StartColumn)
	finding.EndLine = r.EndLine
	finding.EndColumn = vba212UTF16Column(file.Source, r.EndLine, r.EndColumn)
	return finding
}

func vba212Suggestion(kind, guard, unsafe string) string {
	switch strings.ToUpper(kind) {
	case "OR":
		return "Use separate If/ElseIf statements, for example: If " + guard + " Is Nothing Then ... ElseIf " + unsafe + " Then ... End If."
	case "AND":
		return "Use nested If statements, for example: If Not " + guard + " Is Nothing Then If " + unsafe + " Then ... End If."
	case "IIF":
		return "Replace IIf with an If/Else block: If <condition> Then ... Else ... " + unsafe + " ... End If."
	case "CHOOSE":
		return "Replace Choose with Select Case: Case <selected> ... " + unsafe + " ... End Select."
	case "SWITCH":
		return "Replace Switch with ordered If/ElseIf statements and evaluate " + unsafe + " only in its branch."
	default:
		return "Split the guard and " + unsafe + " into separate If statements."
	}
}

func vba212Facts(node *tree_sitter.Node, source []byte, file parsedFile, proc sourceProcedure, scanCtx vba212Context) vba212OperandFacts {
	var facts vba212OperandFacts
	var walk func(*tree_sitter.Node)
	walk = func(current *tree_sitter.Node) {
		if current == nil {
			return
		}
		if current.Kind() == "comparison_expression" {
			if guard, ok := vba212NothingGuard(current, source); ok {
				facts.guards = append(facts.guards, guard)
			}
		}
		if current.Kind() == "call_expression" {
			vba212ArrayFacts(current, source, &facts)
		}
		if current.Kind() == "qualified_member_expression" {
			// Only report the outermost member in a chain; it gives the user the
			// exact unsafe operand rather than an internal receiver fragment.
			parent := current.Parent()
			if parent == nil || parent.Kind() != "qualified_member_expression" || parent.ChildByFieldName("receiver") == nil || parent.ChildByFieldName("receiver").StartByte() != current.StartByte() {
				if path := vba212AccessPath(current, source); path != "" {
					getter := vba212ResolvedGetter(current, source, file, proc, scanCtx)
					facts.unsafe = append(facts.unsafe, vba212Unsafe{path: path, text: strings.TrimSpace(current.Utf8Text(source)), node: current, getter: getter})
				}
			}
		}
		for i := uint(0); i < current.NamedChildCount(); i++ {
			walk(current.NamedChild(i))
		}
	}
	walk(node)
	return facts
}

func vba212NothingGuard(node *tree_sitter.Node, source []byte) (vba212Guard, bool) {
	left, right, op := node.ChildByFieldName("left"), node.ChildByFieldName("right"), node.ChildByFieldName("operator")
	if left == nil || right == nil || op == nil || !strings.EqualFold(strings.TrimSpace(op.Utf8Text(source)), "Is") {
		return vba212Guard{}, false
	}
	if right.Kind() == "nothing_literal" {
		path := vba212AccessPath(left, source)
		return vba212Guard{path: path, name: strings.TrimSpace(left.Utf8Text(source))}, path != ""
	}
	if left.Kind() == "nothing_literal" {
		path := vba212AccessPath(right, source)
		return vba212Guard{path: path, name: strings.TrimSpace(right.Utf8Text(source))}, path != ""
	}
	return vba212Guard{}, false
}

func vba212AccessPath(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	switch node.Kind() {
	case "identifier":
		return strings.ToLower(cleanIdentifier(node.Utf8Text(source)))
	case "parenthesized_expression", "parenthesized_condition_expression", "unary_expression":
		if node.NamedChildCount() == 0 {
			return ""
		}
		return vba212AccessPath(node.NamedChild(0), source)
	case "qualified_member_expression":
		receiver := node.ChildByFieldName("receiver")
		member := node.ChildByFieldName("member")
		if receiver == nil || member == nil {
			return ""
		}
		base := vba212AccessPath(receiver, source)
		name := strings.ToLower(cleanIdentifier(member.Utf8Text(source)))
		if base == "" {
			return name
		}
		return base + "." + name
	case "call_expression":
		fn := node.ChildByFieldName("function")
		if fn == nil {
			return ""
		}
		return vba212AccessPath(fn, source) + "()"
	case "implicit_member_expression":
		return "me"
	}
	return ""
}

func vba212PathMatches(guard, access string) bool {
	guard, access = strings.ToLower(strings.TrimSuffix(guard, "()")), strings.ToLower(access)
	if guard == "" || access == "" {
		return false
	}
	return access == guard || strings.HasPrefix(access, guard+".") || strings.HasPrefix(access, guard+"(") || strings.HasPrefix(access, guard+"()")
}

func vba212ArrayFacts(node *tree_sitter.Node, source []byte, facts *vba212OperandFacts) {
	fn := node.ChildByFieldName("function")
	args := node.ChildByFieldName("arguments")
	if fn == nil || args == nil || fn.Kind() != "identifier" || args.NamedChildCount() == 0 {
		return
	}
	name := strings.ToLower(cleanIdentifier(fn.Utf8Text(source)))
	if name == "isarray" {
		path := vba212AccessPath(args.NamedChild(0), source)
		if path != "" {
			facts.guards = append(facts.guards, vba212Guard{path: path, name: strings.TrimSpace(args.NamedChild(0).Utf8Text(source)), array: true})
		}
		return
	}
	if name != "lbound" && name != "ubound" {
		// In VBA an unqualified call with arguments is also the syntax used for
		// an array/Collection/Dictionary index (arr(i), dict(key)).  Treating
		// it as an access is safe here: it only becomes a finding when a
		// matching guard exists in another eager operand.
		path := vba212AccessPath(fn, source)
		parent := node.Parent()
		isMemberReceiver := parent != nil && parent.Kind() == "qualified_member_expression" && parent.ChildByFieldName("receiver") != nil && parent.ChildByFieldName("receiver").StartByte() == node.StartByte()
		if path != "" && !isMemberReceiver && name != "iif" && name != "choose" && name != "switch" {
			facts.unsafe = append(facts.unsafe, vba212Unsafe{path: path + "()", text: strings.TrimSpace(node.Utf8Text(source)), node: node})
		}
		return
	}
	arg := args.NamedChild(0)
	path := vba212AccessPath(arg, source)
	if path != "" {
		facts.unsafe = append(facts.unsafe, vba212Unsafe{path: path, text: strings.TrimSpace(node.Utf8Text(source)), node: node})
	}
}

func vba212ResolvedGetter(node *tree_sitter.Node, source []byte, file parsedFile, proc sourceProcedure, scanCtx vba212Context) bool {
	member := node.ChildByFieldName("member")
	receiver := node.ChildByFieldName("receiver")
	if member == nil || receiver == nil {
		return false
	}
	typ := vba212ReceiverType(receiver, source, file, proc)
	if typ == "" {
		return false
	}
	memberName := strings.ToLower(cleanIdentifier(member.Utf8Text(source)))
	return scanCtx.getterEffects[strings.ToLower(typ)+"\x00"+memberName]
}

func vba212ReceiverType(receiver *tree_sitter.Node, source []byte, file parsedFile, proc sourceProcedure) string {
	path := vba212AccessPath(receiver, source)
	if path == "" {
		return ""
	}
	if strings.Contains(path, ".") || strings.Contains(path, "(") {
		// The local declaration/type index proves only a direct receiver.  A
		// nested receiver would require a second type-resolution hop and is
		// intentionally treated as unknown rather than guessed.
		return ""
	}
	root := strings.Split(path, ".")[0]
	if root == "me" {
		return file.Module
	}
	decls := moduleDeclarations(file.Lines, []sourceProcedure{proc})
	for _, decl := range proc.Declarations {
		decls[strings.ToLower(decl.Name)] = sourceDeclaration{Name: decl.Name, Type: decl.Type}
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Parameter: true}
	}
	if decl, ok := decls[strings.ToLower(root)]; ok {
		return strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(decl.Type), "New "))
	}
	return ""
}

func vba212UTF16Column(source []byte, line, oneBasedByteColumn int) int {
	if line < 1 || oneBasedByteColumn < 1 {
		return 0
	}
	lines := strings.Split(strings.ReplaceAll(string(source), "\r\n", "\n"), "\n")
	if line > len(lines) {
		return 0
	}
	b := []byte(lines[line-1])
	col := oneBasedByteColumn - 1
	if col > len(b) {
		col = len(b)
	}
	return len(utf16.Encode([]rune(string(b[:col]))))
}
