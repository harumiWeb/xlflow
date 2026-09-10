package analyze

import (
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectCollectionShapeState tracks non-Nothing values stored below object
// variables.  The ordinary object flow deliberately does not infer a value
// from an arbitrary late-bound member chain; this small shape domain only
// records values established by New, Dictionary.Add, guarded Exists branches,
// and a resolved helper that writes an object argument into a Dictionary.
type objectCollectionShapeState struct {
	values       map[string]bool
	objects      map[string]bool
	dictionaries map[string]bool
	collections  map[string]bool
	guarded      map[string]bool
	regexp       map[string]bool
}

func newObjectCollectionShapeState() objectCollectionShapeState {
	return objectCollectionShapeState{
		values:       map[string]bool{},
		objects:      map[string]bool{},
		dictionaries: map[string]bool{},
		collections:  map[string]bool{},
		guarded:      map[string]bool{},
		regexp:       map[string]bool{},
	}
}

func (state objectCollectionShapeState) clone() objectCollectionShapeState {
	clone := newObjectCollectionShapeState()
	for path := range state.values {
		clone.values[path] = true
	}
	for path := range state.objects {
		clone.objects[path] = true
	}
	for path := range state.dictionaries {
		clone.dictionaries[path] = true
	}
	for path := range state.collections {
		clone.collections[path] = true
	}
	for path := range state.guarded {
		clone.guarded[path] = true
	}
	for path := range state.regexp {
		clone.regexp[path] = true
	}
	return clone
}

func objectCollectionShapeStateEqual(left, right objectCollectionShapeState) bool {
	if len(left.values) != len(right.values) || len(left.objects) != len(right.objects) || len(left.dictionaries) != len(right.dictionaries) || len(left.collections) != len(right.collections) || len(left.guarded) != len(right.guarded) || len(left.regexp) != len(right.regexp) {
		return false
	}
	for path := range left.values {
		if !right.values[path] {
			return false
		}
	}
	for path := range left.objects {
		if !right.objects[path] {
			return false
		}
	}
	for path := range left.dictionaries {
		if !right.dictionaries[path] {
			return false
		}
	}
	for path := range left.collections {
		if !right.collections[path] {
			return false
		}
	}
	for path := range left.guarded {
		if !right.guarded[path] {
			return false
		}
	}
	for path := range left.regexp {
		if !right.regexp[path] {
			return false
		}
	}
	return true
}

func objectCollectionShapeIntersect(left, right objectCollectionShapeState) objectCollectionShapeState {
	merged := newObjectCollectionShapeState()
	for path := range left.values {
		if right.values[path] {
			merged.values[path] = true
		}
	}
	for path := range left.objects {
		if right.objects[path] {
			merged.objects[path] = true
		}
	}
	for path := range left.dictionaries {
		if right.dictionaries[path] {
			merged.dictionaries[path] = true
		}
	}
	for path := range left.collections {
		if right.collections[path] {
			merged.collections[path] = true
		}
	}
	for path := range left.guarded {
		if right.guarded[path] {
			merged.guarded[path] = true
		}
	}
	for path := range left.regexp {
		if right.regexp[path] {
			merged.regexp[path] = true
		}
	}
	return merged
}

func objectCollectionShapePath(root string, keys ...string) string {
	path := strings.ToLower(strings.TrimSpace(cleanIdentifier(root)))
	for _, key := range keys {
		key = strings.ToLower(strings.Join(strings.Fields(strings.TrimSpace(key)), ""))
		if key == "" {
			return ""
		}
		if path == "" {
			path = key
		} else {
			path += "|" + key
		}
	}
	return path
}

func objectCollectionShapeKey(text string) string {
	return strings.ToLower(strings.Join(strings.Fields(strings.TrimSpace(text)), ""))
}

func objectCollectionShapePathText(text string) (string, bool) {
	text = strings.TrimSpace(text)
	if text == "" {
		return "", false
	}
	rootEnd := 0
	for rootEnd < len(text) && (text[rootEnd] == '_' || text[rootEnd] >= '0' && text[rootEnd] <= '9' || text[rootEnd] >= 'A' && text[rootEnd] <= 'Z' || text[rootEnd] >= 'a' && text[rootEnd] <= 'z') {
		rootEnd++
	}
	if rootEnd == 0 {
		return "", false
	}
	path := objectCollectionShapePath(text[:rootEnd])
	position := rootEnd
	for {
		for position < len(text) && (text[position] == ' ' || text[position] == '\t') {
			position++
		}
		if position >= len(text) {
			return path, true
		}
		if text[position] == '(' {
			inside, end, ok := dcBalancedContentSpan(text[position:])
			if !ok {
				return "", false
			}
			path = objectCollectionShapePath(path, inside)
			position += end
			continue
		}
		if text[position] != '.' {
			return "", false
		}
		position++
		memberStart := position
		for position < len(text) && (text[position] == '_' || text[position] >= '0' && text[position] <= '9' || text[position] >= 'A' && text[position] <= 'Z' || text[position] >= 'a' && text[position] <= 'z') {
			position++
		}
		if memberStart == position || !strings.EqualFold(text[memberStart:position], "item") {
			return "", false
		}
		for position < len(text) && (text[position] == ' ' || text[position] == '\t') {
			position++
		}
		if position >= len(text) || text[position] != '(' {
			return "", false
		}
		inside, end, ok := dcBalancedContentSpan(text[position:])
		if !ok {
			return "", false
		}
		path = objectCollectionShapePath(path, inside)
		position += end
	}
}

func objectCollectionShapePathHasPrefix(path, prefix string) bool {
	return path == prefix || strings.HasPrefix(path, prefix+"|")
}

func objectCollectionShapeCopy(state *objectCollectionShapeState, source, destination string) {
	if state == nil || source == "" || destination == "" || !state.objects[source] {
		return
	}
	values := make(map[string]bool)
	objects := make(map[string]bool)
	dictionaries := make(map[string]bool)
	collections := make(map[string]bool)
	guarded := make(map[string]bool)
	regexp := make(map[string]bool)
	for path := range state.values {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			values[destination+suffix] = true
		}
	}
	for path := range state.objects {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			objects[destination+suffix] = true
		}
	}
	for path := range state.dictionaries {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			dictionaries[destination+suffix] = true
		}
	}
	for path := range state.collections {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			collections[destination+suffix] = true
			guarded[destination+suffix] = true
		}
	}
	for path := range state.guarded {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			guarded[destination+suffix] = true
		}
	}
	for path := range state.regexp {
		if objectCollectionShapePathHasPrefix(path, source) {
			suffix := strings.TrimPrefix(path, source)
			regexp[destination+suffix] = true
		}
	}
	objectCollectionShapeDelete(state, destination)
	for path := range values {
		state.values[path] = true
	}
	for path := range objects {
		state.objects[path] = true
	}
	for path := range dictionaries {
		state.dictionaries[path] = true
	}
	for path := range collections {
		state.collections[path] = true
	}
	for path := range guarded {
		state.guarded[path] = true
	}
	for path := range regexp {
		state.regexp[path] = true
	}
	state.values[destination] = true
	state.objects[destination] = true
}

func objectCollectionShapeDelete(state *objectCollectionShapeState, prefix string) {
	for path := range state.values {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.values, path)
		}
	}
	for path := range state.objects {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.objects, path)
		}
	}
	for path := range state.dictionaries {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.dictionaries, path)
		}
	}
	for path := range state.collections {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.collections, path)
		}
	}
	for path := range state.guarded {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.guarded, path)
		}
	}
	for path := range state.regexp {
		if objectCollectionShapePathHasPrefix(path, prefix) {
			delete(state.regexp, path)
		}
	}
}

func objectCollectionShapeInvalidate(state *objectCollectionShapeState) {
	if state == nil {
		return
	}
	*state = newObjectCollectionShapeState()
}

func objectCollectionShapeSetValue(state *objectCollectionShapeState, path, value string) {
	if state == nil || path == "" {
		return
	}
	objectCollectionShapeDelete(state, path)
	if value == "nothing" || value == "" {
		return
	}
	if strings.HasPrefix(value, "new collection") {
		state.values[path] = true
		state.objects[path] = true
		state.collections[path] = true
		state.guarded[path] = true
		return
	}
	if strings.HasPrefix(value, "new dictionary") || strings.HasPrefix(value, "new scripting.dictionary") {
		state.values[path] = true
		state.objects[path] = true
		state.dictionaries[path] = true
		return
	}
	if strings.HasPrefix(value, "createobject(\"scripting.dictionary\"") {
		state.values[path] = true
		state.objects[path] = true
		state.dictionaries[path] = true
		return
	}
	if strings.HasPrefix(value, "createobject(\"vbscript.regexp\"") {
		state.values[path] = true
		state.objects[path] = true
		state.regexp[path] = true
		return
	}
	if receiver, ok := objectCollectionShapeRegExpExecuteReceiver(value); ok && state.regexp[receiver] {
		state.values[path] = true
		state.objects[path] = true
		return
	}
	if source, ok := objectCollectionShapePathText(value); ok {
		objectCollectionShapeCopy(state, source, path)
	}
}

func objectCollectionShapeRegExpExecuteReceiver(text string) (string, bool) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	position := strings.Index(lower, ".execute")
	if position <= 0 {
		return "", false
	}
	rest := strings.TrimSpace(text[position+len(".execute"):])
	if !strings.HasPrefix(rest, "(") {
		return "", false
	}
	if _, end, ok := dcBalancedContentSpan(rest); !ok || strings.TrimSpace(rest[end:]) != "" {
		return "", false
	}
	return objectCollectionShapePathText(text[:position])
}

func objectCollectionShapeStatementSource(index *objectContainerIndex, statement procedureir.Statement) string {
	if statement.Kind == procedureir.StatementSet && strings.HasPrefix(strings.ToLower(strings.TrimSpace(statement.Text)), "set ") {
		return strings.TrimSpace(statement.Text)
	}
	if index != nil && statement.Range.StartLine > 0 && statement.Range.StartLine <= len(index.file.Lines) {
		return strings.TrimSpace(index.file.Lines[statement.Range.StartLine-1])
	}
	text := statement.Text
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		text = text[:newline]
	}
	return strings.TrimSpace(text)
}

func objectCollectionShapeSetAssignment(text string) (string, string, bool) {
	text = strings.TrimSpace(text)
	if len(text) < len("Set ") || !strings.EqualFold(text[:len("Set ")], "Set ") {
		return "", "", false
	}
	body := strings.TrimSpace(text[len("Set "):])
	position := objectCollectionShapeTopLevelEqual(body)
	if position < 0 {
		return "", "", false
	}
	return strings.TrimSpace(body[:position]), strings.TrimSpace(body[position+1:]), true
}

func objectCollectionShapeBareAssignment(text string) (string, string, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "set ") {
		return "", "", false
	}
	if strings.HasPrefix(strings.ToLower(text), "let ") {
		text = strings.TrimSpace(text[len("Let "):])
	}
	position := objectCollectionShapeTopLevelEqual(text)
	if position < 0 {
		return "", "", false
	}
	return strings.TrimSpace(text[:position]), strings.TrimSpace(text[position+1:]), true
}

func objectCollectionShapeTopLevelEqual(text string) int {
	depth := 0
	inString := false
	for position := 0; position < len(text); position++ {
		switch text[position] {
		case '"':
			if inString && position+1 < len(text) && text[position+1] == '"' {
				position++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '=':
			if !inString && depth == 0 {
				return position
			}
		}
	}
	return -1
}

func objectCollectionShapeMemberCall(text, member string) (string, []string, bool) {
	lower := strings.ToLower(text)
	needle := "." + strings.ToLower(member)
	position := strings.Index(lower, needle)
	if position <= 0 {
		return "", nil, false
	}
	receiver := strings.TrimSpace(text[:position])
	rest := strings.TrimSpace(text[position+len(needle):])
	if strings.HasPrefix(rest, "(") {
		inside, end, ok := dcBalancedContentSpan(rest)
		if !ok || strings.TrimSpace(rest[end:]) != "" {
			return "", nil, false
		}
		return receiver, dcSplitArgs(inside), true
	}
	return receiver, dcSplitArgs(rest), true
}

func objectCollectionShapeBareCall(text string) (string, []string, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "call ") {
		text = strings.TrimSpace(text[len("Call "):])
	}
	if strings.HasPrefix(strings.ToLower(text), "me.") {
		text = strings.TrimSpace(text[len("Me."):])
	}
	nameEnd := 0
	for nameEnd < len(text) && (text[nameEnd] == '_' || text[nameEnd] >= '0' && text[nameEnd] <= '9' || text[nameEnd] >= 'A' && text[nameEnd] <= 'Z' || text[nameEnd] >= 'a' && text[nameEnd] <= 'z') {
		nameEnd++
	}
	if nameEnd == 0 {
		return "", nil, false
	}
	name := text[:nameEnd]
	rest := strings.TrimSpace(text[nameEnd:])
	if strings.HasPrefix(rest, "(") {
		inside, end, ok := dcBalancedContentSpan(rest)
		if !ok || strings.TrimSpace(rest[end:]) != "" {
			return "", nil, false
		}
		return name, dcSplitArgs(inside), true
	}
	if rest == "" {
		return name, nil, true
	}
	return name, dcSplitArgs(rest), true
}

func objectCollectionShapeKnownObjectExpression(state *objectCollectionShapeState, text string) bool {
	if state == nil {
		return false
	}
	value := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(value, "new ") || strings.HasPrefix(value, "createobject(") {
		return true
	}
	path, ok := objectCollectionShapePathText(value)
	return ok && state.objects[path]
}

func objectCollectionShapeKnownObjectGuard(statement procedureir.Statement, knownObjectParameters map[string]bool) bool {
	if statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf || statement.Condition == nil {
		return false
	}
	condition := strings.ToLower(strings.Join(strings.Fields(statement.Condition.Text), ""))
	condition = strings.TrimSuffix(condition, "then")
	condition = strings.TrimPrefix(condition, "if")
	if !strings.HasPrefix(condition, "isobject(") || !strings.HasSuffix(condition, ")") {
		return false
	}
	parameter := cleanIdentifier(condition[len("isobject(") : len(condition)-1])
	return parameter != "" && knownObjectParameters[parameter]
}

func objectCollectionShapeNormalExitCoveredByStatement(proc sourceProcedure, statementID int, knownObjectParameters map[string]bool) bool {
	if proc.Graph == nil || objectErrorResumeNextAt(proc, statementID) || objectContainerErrorHandlerActiveAt(proc, statementID) {
		return false
	}
	view := proc.Graph.WithoutNormalErrRaiseContinuationView()
	if !view.IsReachable(view.NormalExit()) {
		return false
	}
	target, ok := view.BlockForStatement(statementID)
	if !ok {
		return false
	}
	blocks := map[vbacfg.BlockID]vbacfg.Block{}
	view.ForEachBlock(func(block vbacfg.Block) bool {
		blocks[block.ID] = block
		return true
	})
	type pathState struct {
		block   vbacfg.BlockID
		covered bool
	}
	queue := []pathState{{block: view.Entry()}}
	seen := map[pathState]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		covered := current.covered || current.block == target.ID
		if current.block == view.NormalExit() {
			if !covered {
				return false
			}
			continue
		}
		block := blocks[current.block]
		view.ForEachOutgoing(current.block, func(edge vbacfg.Edge) bool {
			if block.Statement != nil && edge.Kind == vbacfg.EdgeBranchFalse && objectCollectionShapeKnownObjectGuard(*block.Statement, knownObjectParameters) {
				return true
			}
			nextCovered := covered
			if edge.Class == vbacfg.EdgeExceptional && current.block == target.ID {
				nextCovered = false
			}
			queue = append(queue, pathState{block: edge.To, covered: nextCovered})
			return true
		})
	}
	return true
}

func objectCollectionShapeStatementReachable(proc sourceProcedure, statementID int, knownObjectParameters map[string]bool) bool {
	if proc.Graph == nil {
		return false
	}
	view := proc.Graph.WithoutNormalErrRaiseContinuationView()
	target, ok := view.BlockForStatement(statementID)
	if !ok {
		return false
	}
	blocks := map[vbacfg.BlockID]vbacfg.Block{}
	view.ForEachBlock(func(block vbacfg.Block) bool {
		blocks[block.ID] = block
		return true
	})
	queue := []vbacfg.BlockID{view.Entry()}
	seen := map[vbacfg.BlockID]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current == target.ID {
			return true
		}
		block := blocks[current]
		view.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if block.Statement != nil && edge.Kind == vbacfg.EdgeBranchFalse && objectCollectionShapeKnownObjectGuard(*block.Statement, knownObjectParameters) {
				return true
			}
			queue = append(queue, edge.To)
			return true
		})
	}
	return false
}

func objectCollectionShapeExternalObjectPath(proc sourceProcedure, declarations declarationScope, text string) bool {
	path, ok := objectCollectionShapePathText(text)
	if !ok {
		return false
	}
	root := strings.Split(path, "|")[0]
	declaration, scope, declared := objectDeclarationBinding(root, declarations)
	return declared && declaration.Object && (scope == procedureir.ScopeModule || scope == procedureir.ScopeParameter)
}

func objectCollectionShapeHelperHasUnsupportedMutation(index *objectContainerIndex, callee sourceProcedure, knownObjectParameters map[string]bool, recognized map[int]bool) bool {
	if index == nil {
		return true
	}
	declarations := objectFlowDeclarations(index.file, callee, index.moduleDecls)
	for statement := range callee.Statements.All() {
		if !objectCollectionShapeStatementReachable(callee, statement.ID, knownObjectParameters) {
			continue
		}
		text := objectCollectionShapeStatementSource(index, statement)
		for _, assignment := range []func(string) (string, string, bool){objectCollectionShapeSetAssignment, objectCollectionShapeBareAssignment} {
			target, _, ok := assignment(text)
			if ok && objectCollectionShapeExternalObjectPath(callee, declarations, target) && !recognized[statement.ID] {
				return true
			}
		}
		for _, member := range []string{"Add", "Remove", "RemoveAll"} {
			receiver, _, ok := objectCollectionShapeMemberCall(text, member)
			if ok && objectCollectionShapeExternalObjectPath(callee, declarations, receiver) {
				return true
			}
		}
	}
	for call := range callee.Calls.All() {
		if !objectCollectionShapeStatementReachable(callee, call.StatementID, knownObjectParameters) {
			continue
		}
		if call.Callee.Receiver != nil {
			receiver := strings.TrimSpace(*call.Callee.Receiver)
			if objectCollectionShapeExternalObjectPath(callee, declarations, receiver) {
				switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
				case "exists", "item", "count", "keys":
				default:
					return true
				}
			}
		}
		for _, argument := range objectContainerCallArguments(callee, call) {
			if objectCollectionShapeExternalObjectPath(callee, declarations, argument) && call.Resolution.Status != procedureir.ResolutionBuiltinLike {
				return true
			}
		}
	}
	return false
}

func objectCollectionShapeApplyStatement(index *objectContainerIndex, proc sourceProcedure, statement procedureir.Statement, state *objectCollectionShapeState) {
	if index == nil || state == nil || objectErrorResumeNextAt(proc, statement.ID) || objectContainerErrorHandlerActiveAt(proc, statement.ID) {
		return
	}
	text := objectCollectionShapeStatementSource(index, statement)
	if target, value, ok := objectCollectionShapeSetAssignment(text); ok {
		path, pathOK := objectCollectionShapePathText(target)
		if pathOK {
			objectCollectionShapeSetValue(state, path, strings.ToLower(strings.TrimSpace(value)))
		}
		return
	}
	if target, value, ok := objectCollectionShapeBareAssignment(text); ok {
		path, pathOK := objectCollectionShapePathText(target)
		if pathOK {
			objectCollectionShapeSetValue(state, path, strings.ToLower(strings.TrimSpace(value)))
		}
		return
	}
	if receiver, args, ok := objectCollectionShapeMemberCall(text, "Add"); ok && len(args) >= 2 {
		receiverPath, receiverOK := objectCollectionShapePathText(receiver)
		if receiverOK && state.dictionaries[receiverPath] {
			key := objectCollectionShapeKey(args[0])
			if key != "" {
				value := strings.ToLower(strings.TrimSpace(args[1]))
				objectCollectionShapeSetValue(state, objectCollectionShapePath(receiverPath, key), value)
			}
		}
		return
	}
	name, args, ok := objectCollectionShapeBareCall(text)
	if ok && statement.Kind == procedureir.StatementCall && !objectCollectionShapeApplyHelper(index, proc, name, args, state) {
		objectCollectionShapeInvalidate(state)
	}
}

func objectCollectionShapeApplyHelper(index *objectContainerIndex, proc sourceProcedure, name string, args []string, state *objectCollectionShapeState) bool {
	if index == nil || state == nil {
		return false
	}
	var callee sourceProcedure
	found := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(candidate.Module, proc.Module) || !strings.EqualFold(candidate.Name, name) {
			continue
		}
		if found {
			return false
		}
		callee, found = candidate, true
	}
	if !found {
		return false
	}
	parameters := make([]string, 0, callee.Params.Len())
	for parameter := range callee.Params.All() {
		parameters = append(parameters, strings.ToLower(cleanIdentifier(parameter.Name)))
	}
	knownObjectParameters := map[string]bool{}
	for parameterIndex, parameter := range parameters {
		if parameterIndex < len(args) && objectCollectionShapeKnownObjectExpression(state, args[parameterIndex]) {
			knownObjectParameters[parameter] = true
		}
	}
	recognized := map[int]bool{}
	for statement := range callee.Statements.All() {
		text := objectCollectionShapeStatementSource(index, statement)
		target, value, ok := objectCollectionShapeSetAssignment(text)
		if !ok {
			continue
		}
		targetPath, targetOK := objectCollectionShapePathText(target)
		valueName := strings.ToLower(cleanIdentifier(value))
		if !targetOK || valueName == "" {
			continue
		}
		valueParameter := -1
		for parameterIndex, parameter := range parameters {
			if parameter == valueName && parameterIndex < len(args) {
				valueParameter = parameterIndex
				break
			}
		}
		if valueParameter < 0 {
			continue
		}
		keyParameter := -1
		keyParameterName := ""
		for segmentIndex, segment := range strings.Split(targetPath, "|") {
			for parameterIndex, parameter := range parameters {
				if parameterIndex != valueParameter && parameter == segment && parameterIndex < len(args) {
					keyParameter = segmentIndex
					keyParameterName = parameter
				}
			}
		}
		if keyParameter < 0 || keyParameter >= len(strings.Split(targetPath, "|")) {
			continue
		}
		if !objectCollectionShapeNormalExitCoveredByStatement(callee, statement.ID, knownObjectParameters) {
			continue
		}
		actualKeyIndex := -1
		for parameterIndex, parameter := range parameters {
			if parameter == keyParameterName {
				actualKeyIndex = parameterIndex
				break
			}
		}
		if actualKeyIndex < 0 || actualKeyIndex >= len(args) {
			continue
		}
		segments := strings.Split(targetPath, "|")
		segments[keyParameter] = objectCollectionShapeKey(args[actualKeyIndex])
		destination := strings.Join(segments, "|")
		if actualValue, valueOK := objectCollectionShapePathText(args[valueParameter]); valueOK {
			objectCollectionShapeCopy(state, actualValue, destination)
		} else {
			objectCollectionShapeSetValue(state, destination, strings.ToLower(strings.TrimSpace(args[valueParameter])))
		}
		recognized[statement.ID] = true
	}
	return !objectCollectionShapeHelperHasUnsupportedMutation(index, callee, knownObjectParameters, recognized)
}

func objectCollectionShapeCollectionValue(proc sourceProcedure, value string, declarations declarationScope) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	if strings.HasPrefix(value, "new collection") {
		return true
	}
	path, ok := objectCollectionShapePathText(value)
	if !ok {
		return false
	}
	root := strings.Split(path, "|")[0]
	declaration, _, declared := objectDeclarationBinding(root, declarations)
	return declared && declaration.Object && dcKindFromType(declaration.Type) == dcCollection
}

func objectCollectionShapeCollectionContracts(index *objectContainerIndex, proc sourceProcedure) map[string]bool {
	contracts := map[string]bool{}
	if index == nil {
		return contracts
	}
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	writes := map[string]bool{}
	unsafe := map[string]bool{}
	markWrite := func(path string, collection bool) {
		segments := strings.Split(path, "|")
		if len(segments) == 0 || !strings.HasPrefix(segments[len(segments)-1], "\"") || !strings.HasSuffix(segments[len(segments)-1], "\"") {
			return
		}
		writes[path] = true
		if !collection {
			unsafe[path] = true
		}
	}
	for statement := range proc.Statements.All() {
		text := objectCollectionShapeStatementSource(index, statement)
		if target, value, ok := objectCollectionShapeSetAssignment(text); ok {
			if path, pathOK := objectCollectionShapePathText(target); pathOK {
				markWrite(path, objectCollectionShapeCollectionValue(proc, value, declarations))
			}
		}
		if target, value, ok := objectCollectionShapeBareAssignment(text); ok {
			if path, pathOK := objectCollectionShapePathText(target); pathOK {
				markWrite(path, objectCollectionShapeCollectionValue(proc, value, declarations))
			}
		}
		if receiver, args, ok := objectCollectionShapeMemberCall(text, "Add"); ok && len(args) >= 2 {
			if receiverPath, receiverOK := objectCollectionShapePathText(receiver); receiverOK {
				markWrite(objectCollectionShapePath(receiverPath, args[0]), objectCollectionShapeCollectionValue(proc, args[1], declarations))
			}
		}
		if receiver, args, ok := objectCollectionShapeMemberCall(text, "Remove"); ok && len(args) >= 1 {
			if receiverPath, receiverOK := objectCollectionShapePathText(receiver); receiverOK {
				markWrite(objectCollectionShapePath(receiverPath, args[0]), false)
			}
		}
	}
	for path := range writes {
		if !unsafe[path] {
			contracts[path] = true
		}
	}
	return contracts
}

func objectCollectionShapeExistsCondition(text string) (string, string, bool, bool) {
	lower := strings.ToLower(text)
	position := strings.Index(lower, ".exists")
	if position <= 0 {
		return "", "", false, false
	}
	receiver := strings.TrimSpace(text[:position])
	if strings.HasPrefix(strings.ToLower(receiver), "if ") {
		receiver = strings.TrimSpace(receiver[3:])
	}
	negated := false
	if strings.HasPrefix(strings.ToLower(receiver), "not ") {
		negated = true
		receiver = strings.TrimSpace(receiver[4:])
	}
	rest := strings.TrimSpace(text[position+len(".Exists"):])
	inside, end, ok := dcBalancedContentSpan(rest)
	if !ok || strings.TrimSpace(rest[end:]) != "" && !strings.EqualFold(strings.TrimSpace(rest[end:]), "then") {
		return "", "", false, false
	}
	return receiver, inside, negated, true
}

func objectCollectionShapeDefaultItemCondition(text string) (string, string, bool, bool) {
	text = strings.TrimSpace(text)
	if len(text) >= len("If ") && strings.EqualFold(text[:len("If ")], "If ") {
		text = strings.TrimSpace(text[len("If "):])
	}
	if then := strings.Index(strings.ToLower(text), " then"); then >= 0 {
		text = strings.TrimSpace(text[:then])
	}
	negated := false
	if len(text) >= len("Not ") && strings.EqualFold(text[:len("Not ")], "Not ") {
		negated = true
		text = strings.TrimSpace(text[len("Not "):])
	}
	path, ok := objectCollectionShapePathText(text)
	if !ok {
		return "", "", false, false
	}
	parts := strings.Split(path, "|")
	if len(parts) != 2 || parts[0] == "" || parts[1] == "" {
		return "", "", false, false
	}
	return parts[0], parts[1], negated, true
}

func objectCollectionShapeGuardBranchImpossible(proc sourceProcedure, statement procedureir.Statement, edge vbacfg.Edge, index *objectContainerIndex) bool {
	if edge.Kind != vbacfg.EdgeBranchFalse {
		return false
	}
	text := objectCollectionShapeStatementSource(index, statement)
	receiver, key, negated, ok := objectCollectionShapeDefaultItemCondition(text)
	if !ok || !negated {
		return false
	}
	return objectCollectionShapeStaticDictionaryKeyAbsent(proc, receiver, key, statement.ID, index)
}

// objectCollectionShapeStaticDictionaryKeyAbsent proves only a narrow lazy-cache
// case: a local Static Dictionary is constructed in the procedure and no keyed
// write or alias can mutate it. This lets the false branch of
// `If Not cache(key) Then` be removed without assuming that dynamic keys which
// look different in one invocation stay different across invocations.
func objectCollectionShapeStaticDictionaryKeyAbsent(proc sourceProcedure, receiver, key string, guardID int, index *objectContainerIndex) bool {
	if receiver == "" || key == "" || strings.Contains(receiver, "|") || index == nil {
		return false
	}
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	declaration, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || scope != procedureir.ScopeLocal || !declaration.Static || !declaration.Object {
		return false
	}
	receiverPath := objectCollectionShapePath(receiver)
	for statement := range proc.Statements.All() {
		text := objectCollectionShapeStatementSource(index, statement)
		target, value, assignmentOK := objectCollectionShapeSetAssignment(text)
		if !assignmentOK {
			target, value, assignmentOK = objectCollectionShapeBareAssignment(text)
		}
		if !assignmentOK {
			continue
		}
		targetPath, targetOK := objectCollectionShapePathText(target)
		if !targetOK {
			continue
		}
		if sourcePath, sourceOK := objectCollectionShapePathText(objectTrimOuterParens(value)); sourceOK && sourcePath == receiverPath {
			return false
		}
		if strings.HasPrefix(targetPath, receiverPath+"|") {
			return false
		}
		if targetPath == receiverPath {
			value = strings.ToLower(strings.TrimSpace(value))
			if value == "" || value == "nothing" {
				continue
			}
			if objectCollectionShapeKnownValueKind(proc, receiverPath, index) != "dictionary" {
				return false
			}
		}
	}
	for call := range proc.Calls.All() {
		if call.StatementID == guardID {
			continue
		}
		if call.Callee.Receiver != nil {
			receiverText := strings.TrimSpace(*call.Callee.Receiver)
			if objectCollectionShapePath(receiverText) == receiverPath {
				member := strings.ToLower(cleanIdentifier(call.Callee.Member))
				switch member {
				case "", "exists", "item", "count", "keys":
				default:
					return false
				}
			}
		}
		for _, actual := range objectCallActuals(call, proc.analysisFacts()) {
			if strings.EqualFold(cleanIdentifier(actual.text), cleanIdentifier(receiver)) {
				return false
			}
		}
	}
	return true
}

func objectCollectionShapeApplyGuard(proc sourceProcedure, statement procedureir.Statement, edge vbacfg.Edge, state *objectCollectionShapeState, index *objectContainerIndex, contracts map[string]bool) {
	if state == nil || (edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse) {
		return
	}
	text := statement.Text
	if index != nil && statement.Range.StartLine > 0 && statement.Range.StartLine <= len(index.file.Lines) {
		text = index.file.Lines[statement.Range.StartLine-1]
	}
	receiver, key, negated, ok := objectCollectionShapeExistsCondition(strings.TrimSpace(text))
	if !ok {
		condition := strings.TrimSpace(text)
		if len(condition) >= len("If ") && strings.EqualFold(condition[:len("If ")], "If ") {
			condition = strings.TrimSpace(condition[len("If "):])
		}
		name, negated, singleOK := objectSingleNothingGuard(condition)
		if !singleOK {
			return
		}
		path := objectCollectionShapePath(name)
		nonNothing := (edge.Kind == vbacfg.EdgeBranchTrue && negated) ||
			(edge.Kind == vbacfg.EdgeBranchFalse && !negated)
		if !nonNothing {
			objectCollectionShapeDelete(state, path)
			return
		}
		state.values[path] = true
		switch objectCollectionShapeKnownValueKind(proc, path, index) {
		case "collection":
			state.objects[path] = true
			state.collections[path] = true
			state.guarded[path] = true
		case "dictionary":
			state.objects[path] = true
			state.dictionaries[path] = true
		case "regexp":
			state.objects[path] = true
			state.regexp[path] = true
		}
		return
	}
	receiverPath, receiverOK := objectCollectionShapePathText(receiver)
	if !receiverOK {
		return
	}
	path := objectCollectionShapePath(receiverPath, key)
	if state.values[receiverPath] {
		state.dictionaries[receiverPath] = true
	}
	present := edge.Kind == vbacfg.EdgeBranchTrue
	if negated {
		present = !present
	}
	if present {
		state.values[path] = true
		if contracts[path] {
			state.objects[path] = true
			state.collections[path] = true
		}
	} else {
		objectCollectionShapeDelete(state, path)
	}
}

func objectCollectionShapeKnownValueKind(proc sourceProcedure, path string, index *objectContainerIndex) string {
	if path == "" || strings.Contains(path, "|") {
		return ""
	}
	kind := ""
	for statement := range proc.Statements.All() {
		text := objectCollectionShapeStatementSource(index, statement)
		target, value, ok := objectCollectionShapeSetAssignment(text)
		if !ok {
			continue
		}
		targetPath, targetOK := objectCollectionShapePathText(target)
		if !targetOK || targetPath != path {
			continue
		}
		value = strings.ToLower(strings.TrimSpace(value))
		if value == "" || value == "nothing" {
			continue
		}
		current := ""
		switch {
		case strings.HasPrefix(value, "new collection"):
			current = "collection"
		case strings.HasPrefix(value, "new dictionary"), strings.HasPrefix(value, "new scripting.dictionary"), strings.HasPrefix(value, "createobject(\"scripting.dictionary\""):
			current = "dictionary"
		case strings.HasPrefix(value, "createobject(\"vbscript.regexp\""):
			current = "regexp"
		default:
			return ""
		}
		if kind != "" && kind != current {
			return ""
		}
		kind = current
	}
	return kind
}

func objectCollectionShapeInitialState(declarations declarationScope) objectCollectionShapeState {
	state := newObjectCollectionShapeState()
	for _, scope := range []map[string]sourceDeclaration{declarations.module, declarations.extra, declarations.local, declarations.parameters} {
		for name, declaration := range scope {
			if !declaration.Object || !declaration.NewExpression {
				continue
			}
			path := objectCollectionShapePath(name)
			state.values[path] = true
			state.objects[path] = true
			if dcKindFromType(declaration.Type) == dcCollection {
				state.collections[path] = true
				state.guarded[path] = true
			} else if dcKindFromType(declaration.Type) == dcDictionary {
				state.dictionaries[path] = true
			}
		}
	}
	return state
}

func objectCollectionShapeBeforeStatement(proc sourceProcedure, statementID int, context objectFlowContext, declarations declarationScope) (objectCollectionShapeState, bool) {
	if context.graph.BlockCount() == 0 {
		return objectCollectionShapeState{}, false
	}
	target, ok := context.graph.BlockForStatement(statementID)
	if !ok {
		return objectCollectionShapeState{}, false
	}
	input := map[vbacfg.BlockID]objectCollectionShapeState{}
	seen := map[vbacfg.BlockID]bool{}
	entry := context.graph.Entry()
	input[entry] = objectCollectionShapeInitialState(declarations)
	seen[entry] = true
	queue := []vbacfg.BlockID{entry}
	queued := map[vbacfg.BlockID]bool{entry: true}
	contracts := objectCollectionShapeCollectionContracts(context.containerIndex, proc)
	blocks := map[vbacfg.BlockID]vbacfg.Block{}
	context.graph.ForEachBlock(func(block vbacfg.Block) bool {
		blocks[block.ID] = block
		return true
	})
	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		queued[blockID] = false
		before := input[blockID]
		after := before.clone()
		block, blockOK := blocks[blockID]
		if blockOK && block.Statement != nil {
			objectCollectionShapeApplyStatement(context.containerIndex, proc, *block.Statement, &after)
		}
		context.graph.ForEachOutgoing(blockID, func(edge vbacfg.Edge) bool {
			if blockOK && block.Statement != nil && objectCollectionShapeGuardBranchImpossible(proc, *block.Statement, edge, context.containerIndex) {
				return true
			}
			candidate := after.clone()
			if edge.Class == vbacfg.EdgeExceptional {
				candidate = before.clone()
			}
			if blockOK && block.Statement != nil {
				objectCollectionShapeApplyGuard(proc, *block.Statement, edge, &candidate, context.containerIndex, contracts)
			}
			if !seen[edge.To] {
				input[edge.To] = candidate
				seen[edge.To] = true
				if !queued[edge.To] {
					queue = append(queue, edge.To)
					queued[edge.To] = true
				}
				return true
			}
			merged := objectCollectionShapeIntersect(input[edge.To], candidate)
			if !objectCollectionShapeStateEqual(merged, input[edge.To]) {
				input[edge.To] = merged
				if !queued[edge.To] {
					queue = append(queue, edge.To)
					queued[edge.To] = true
				}
			}
			return true
		})
	}
	state, ok := input[target.ID]
	return state, ok
}

func objectCollectionShapeExpressionAssigned(proc sourceProcedure, expression procedureir.Expression, statementID int, context objectFlowContext, declarations declarationScope) bool {
	if expression.Kind != procedureir.ExpressionCall && expression.Kind != procedureir.ExpressionMember {
		return false
	}
	statement, ok := context.facts.Statement(statementID)
	if !ok || statement.Target == nil || statement.Target.Kind == procedureir.ExpressionCall || statement.Target.Kind == procedureir.ExpressionMember {
		return false
	}
	targetName := cleanIdentifier(statement.Target.Text)
	declaration, _, declared := objectDeclarationBinding(targetName, declarations)
	if !declared || !declaration.Object {
		return false
	}
	path, pathOK := objectCollectionShapePathText(expression.Text)
	if !pathOK {
		return false
	}
	if context.shapeStateReady == nil {
		context.shapeStateReady = map[int]bool{}
		context.shapeStates = map[int]objectCollectionShapeState{}
	}
	state, ready := context.shapeStates[statementID]
	if !context.shapeStateReady[statementID] {
		state, ready = objectCollectionShapeBeforeStatement(proc, statementID, context, declarations)
		context.shapeStates[statementID] = state
		context.shapeStateReady[statementID] = true
	}
	if !ready {
		return false
	}
	return state.objects[path] || state.collections[path] || state.guarded[path] && dcKindFromType(declaration.Type) == dcCollection
}
