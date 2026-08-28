package intel

import (
	"context"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// interactiveDocumentIndex contains only source declarations needed by
// latency-sensitive editor queries. Procedure bodies and their locals are
// intentionally absent; locals are loaded for the current procedure on
// demand by the query path.
//
// The index is immutable after construction. Posting slices contain indexes
// into symbols and are therefore safe to share between concurrent readers.
type interactiveDocumentIndex struct {
	symbols       []Symbol
	incomplete    bool
	recoveryKnown bool
	exact         map[string][]int
	qualified     map[string][]int
	module        map[string][]int
	kind          map[string][]int
	prefix        map[string][]int
	prefixKeys    []string
}

func initializeInteractiveDocumentIndex(symbols []Symbol) *interactiveDocumentIndex {
	idx := &interactiveDocumentIndex{
		symbols:   symbols,
		exact:     make(map[string][]int),
		qualified: make(map[string][]int),
		module:    make(map[string][]int),
		kind:      make(map[string][]int),
		prefix:    make(map[string][]int),
	}
	for i, symbol := range idx.symbols {
		name := indexName(symbol.Name)
		if name == "" {
			continue
		}
		idx.exact[name] = append(idx.exact[name], i)
		qualifiedName := indexName(qualifiedName(symbol.Module, symbol.Name))
		if qualifiedName != "" {
			idx.qualified[qualifiedName] = append(idx.qualified[qualifiedName], i)
		}
		module := indexName(symbol.Module)
		if module != "" {
			idx.module[module] = append(idx.module[module], i)
		}
		kind := indexName(symbol.Kind)
		if kind != "" {
			idx.kind[kind] = append(idx.kind[kind], i)
		}
		idx.prefix[name] = append(idx.prefix[name], i)
		if qualifiedName != "" {
			idx.prefix[qualifiedName] = append(idx.prefix[qualifiedName], i)
		}
	}
	idx.prefixKeys = make([]string, 0, len(idx.prefix))
	for key := range idx.prefix {
		idx.prefixKeys = append(idx.prefixKeys, key)
	}
	sort.Strings(idx.prefixKeys)
	return idx
}

// buildInteractiveDocumentIndexOwned transfers ownership of symbols produced
// by the declaration extractor into the immutable index. The caller must not
// mutate the slice or any nested metadata after this call; using the owned
// variant avoids a second deep clone on the cold path while query results
// remain defensive clones.
func buildInteractiveDocumentIndexOwned(symbols []Symbol) *interactiveDocumentIndex {
	return initializeInteractiveDocumentIndex(symbols)
}

func (idx *interactiveDocumentIndex) query(query WorkspaceSymbolQuery) []Symbol {
	if idx == nil {
		return nil
	}
	key := indexName(query.Text)
	var postings []int
	switch query.Mode {
	case WorkspaceSymbolQueryExact:
		postings = idx.exact[key]
	case WorkspaceSymbolQueryQualified:
		postings = idx.qualified[key]
		if len(postings) == 0 && !strings.Contains(key, ".") {
			// Historical qualified callers also accepted an unqualified name;
			// preserve that compatibility while retaining strict module matching
			// whenever the query contains a qualifier.
			postings = idx.exact[key]
		}
	case WorkspaceSymbolQueryModule:
		postings = idx.module[key]
	case WorkspaceSymbolQueryKind:
		postings = idx.kind[key]
	case WorkspaceSymbolQueryPrefix:
		if key == "" {
			postings = make([]int, len(idx.symbols))
			for i := range postings {
				postings[i] = i
			}
			break
		}
		seen := make(map[int]struct{})
		start := sort.SearchStrings(idx.prefixKeys, key)
		for i := start; i < len(idx.prefixKeys) && strings.HasPrefix(idx.prefixKeys[i], key); i++ {
			for _, posting := range idx.prefix[idx.prefixKeys[i]] {
				if _, ok := seen[posting]; ok {
					continue
				}
				seen[posting] = struct{}{}
				postings = append(postings, posting)
			}
		}
	case WorkspaceSymbolQueryContains:
		for i, symbol := range idx.symbols {
			if key == "" || strings.Contains(indexName(symbol.Name), key) || strings.Contains(indexName(qualifiedName(symbol.Module, symbol.Name)), key) {
				postings = append(postings, i)
			}
		}
	default:
		return nil
	}
	result := make([]Symbol, 0, len(postings))
	for _, posting := range postings {
		if posting >= 0 && posting < len(idx.symbols) {
			result = append(result, idx.symbols[posting])
		}
	}
	sort.SliceStable(result, func(i, j int) bool {
		if result[i].File != result[j].File {
			return result[i].File < result[j].File
		}
		if result[i].Range.Start.Line != result[j].Range.Start.Line {
			return result[i].Range.Start.Line < result[j].Range.Start.Line
		}
		if result[i].Range.Start.Character != result[j].Range.Start.Character {
			return result[i].Range.Start.Character < result[j].Range.Start.Character
		}
		return result[i].Name < result[j].Name
	})
	return cloneAnalysisSymbols(result)
}

func (s *AnalysisSnapshot) buildInteractiveIndex() (*interactiveDocumentIndex, bool, error) {
	return s.buildInteractiveIndexContext(context.Background())
}

func (s *AnalysisSnapshot) buildInteractiveIndexContext(ctx context.Context) (*interactiveDocumentIndex, bool, error) {
	if s == nil {
		return nil, false, errAnalysisSnapshotRetired
	}
	if ctx == nil {
		ctx = context.Background()
	}
	for {
		if err := ctx.Err(); err != nil {
			return nil, false, err
		}
		s.interactiveIndexMu.Lock()
		if s.interactiveIndexDone {
			idx, err := s.interactiveIndex, s.interactiveIndexErr
			s.interactiveIndexMu.Unlock()
			s.interactiveIndexHits.Add(1)
			return idx, true, err
		}
		if s.interactiveIndexWait != nil {
			wait := s.interactiveIndexWait
			s.interactiveIndexMu.Unlock()
			finishWait := analysisstats.MeasureWait(ctx, "interactive_index_singleflight")
			select {
			case <-ctx.Done():
				finishWait(ctx.Err())
				return nil, false, ctx.Err()
			case <-wait:
				finishWait(nil)
			}
			continue
		}
		wait := make(chan struct{})
		s.interactiveIndexWait = wait
		s.interactiveIndexMu.Unlock()

		idx, err := s.buildInteractiveIndexValue(ctx)
		if err == nil {
			err = ctx.Err()
		}
		s.interactiveIndexMu.Lock()
		if !isRetryableContextError(err) {
			s.interactiveIndex, s.interactiveIndexErr, s.interactiveIndexDone = idx, err, true
		}
		if s.interactiveIndexWait == wait {
			s.interactiveIndexWait = nil
			close(wait)
		}
		s.interactiveIndexMu.Unlock()
		return idx, false, err
	}
}

func (s *AnalysisSnapshot) buildInteractiveIndexValue(ctx context.Context) (*interactiveDocumentIndex, error) {
	s.interactiveIndexBuilds.Add(1)
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	// Reuse a full snapshot tree when another analysis has already prepared
	// it. Cold interactive requests deliberately avoid creating that tree: a
	// compact source containing module declarations and procedure headers is
	// enough for the declaration-only extractor and is proportional to the
	// declaration surface rather than the procedure bodies.
	var (
		file    symbols.FileResult
		err     error
		lineMap []int
	)
	s.parsedMu.Lock()
	parsed := s.parsedDocument
	parsedErr := s.parsedErr
	s.parsedMu.Unlock()
	if parsed != nil {
		file, err = symbols.InspectParsedDeclarationsContext(ctx, symbols.SourceOptions{
			RootDir:        filepath.Dir(s.path),
			Path:           s.path,
			ModuleKind:     s.moduleKind,
			IncludePrivate: true,
			IncludeLabels:  false,
		}, parsed)
	} else if parsedErr != nil {
		// A failed full parse is not a reason to make an interactive request
		// pay the same cost again. The lexical declaration source below still
		// gives best-effort names and lets parser-recovery callers fail open.
		file, lineMap, err = s.inspectCompactDeclarations(ctx)
	} else {
		file, lineMap, err = s.inspectCompactDeclarations(ctx)
	}
	if err != nil {
		return nil, err
	}
	if len(lineMap) != 0 {
		rebaseDeclarationFileResult(&file, lineMap)
	}
	converted := symbolsFromFile(file, "")
	module := moduleNameForDocument(s.Document())
	for i := range converted {
		converted[i].File = s.path
		converted[i].ModuleKind = s.moduleKind
		if converted[i].Module == "" {
			converted[i].Module = module
		}
	}
	idx := buildInteractiveDocumentIndexOwned(converted)
	idx.incomplete = file.Parse.HasError || file.Parse.HasMissing || compactDeclarationRecovery(s.lines, s.Procedures())
	idx.recoveryKnown = parsed != nil
	return idx, nil
}

// inspectCompactDeclarations builds the smallest valid tree-sitter input that
// retains module declarations, procedure headers, and their source positions.
// It is used only before a snapshot's full tree exists; the normal semantic
// path continues to use the original parsed source.
func (s *AnalysisSnapshot) inspectCompactDeclarations(ctx context.Context) (symbols.FileResult, []int, error) {
	if s == nil {
		return symbols.FileResult{}, nil, errAnalysisSnapshotRetired
	}
	procedures := s.Procedures()
	lines := s.sourceLines()
	compact, lineMap := compactDeclarationSource(lines, procedures)
	if err := ctx.Err(); err != nil {
		return symbols.FileResult{}, nil, err
	}
	parsed, err := vbaast.ParseDocument(s.path, []byte(compact))
	if err != nil {
		return symbols.FileResult{}, nil, err
	}
	defer parsed.Close()
	file, err := symbols.InspectParsedDeclarationsContext(ctx, symbols.SourceOptions{
		RootDir:        filepath.Dir(s.path),
		Path:           s.path,
		ModuleKind:     s.moduleKind,
		IncludePrivate: true,
		IncludeLabels:  false,
	}, parsed)
	return file, lineMap, err
}

func compactDeclarationSource(lines []string, procedures []ProcedureInfo) (string, []int) {
	if len(lines) == 0 {
		return "", nil
	}
	starts := make(map[int]ProcedureInfo, len(procedures))
	for _, procedure := range procedures {
		if procedure.Range.Start.Line >= 0 && procedure.Range.Start.Line < len(lines) {
			starts[procedure.Range.Start.Line] = procedure
		}
	}
	var out strings.Builder
	lineMap := make([]int, 0, len(lines))
	appendLine := func(text string, originalLine int) {
		out.WriteString(text)
		out.WriteByte('\n')
		lineMap = append(lineMap, originalLine)
	}
	activeEnd := -1
	for lineNo := 0; lineNo < len(lines); {
		if activeEnd >= 0 {
			if lineNo > activeEnd {
				activeEnd = -1
				continue
			}
			if lineNo == activeEnd {
				appendLine(lines[lineNo], lineNo)
				activeEnd = -1
			}
			lineNo++
			continue
		}
		procedure, ok := starts[lineNo]
		if !ok {
			appendLine(lines[lineNo], lineNo)
			lineNo++
			continue
		}
		endHeader := lineNo
		for endHeader < len(lines) {
			appendLine(lines[endHeader], endHeader)
			if !strings.HasSuffix(strings.TrimSpace(lines[endHeader]), "_") {
				endHeader++
				break
			}
			endHeader++
		}
		endLine := procedure.Range.End.Line
		if endLine >= endHeader && endLine < len(lines) {
			activeEnd = endLine
		} else if endLine < lineNo {
			appendLine(compactProcedureEnd(lines[lineNo]), min(len(lines)-1, lineNo))
		} else if endLine >= len(lines) {
			appendLine(compactProcedureEnd(lines[lineNo]), len(lines)-1)
			lineNo = len(lines)
			continue
		}
		lineNo++
	}
	return out.String(), lineMap
}

func compactProcedureEnd(declaration string) string {
	lower := strings.ToLower(declaration)
	if strings.Contains(lower, "property") {
		return "End Property"
	}
	if strings.Contains(lower, "function") {
		return "End Function"
	}
	return "End Sub"
}

// compactDeclarationRecovery preserves the conservative behavior of the full
// parser for errors that disappear when procedure bodies are omitted. The
// declaration parser still reports malformed headers and module declarations;
// this bounded lexical pass covers the structural errors most likely to make a
// recovered document unsafe for exact or prefix lookup (unclosed procedures,
// block statements, and conditional-compilation branches).
func compactDeclarationRecovery(lines []string, procedures []ProcedureInfo) bool {
	conditionalDepth := 0
	for _, line := range lines {
		text := strings.TrimSpace(line[:codeLimit(line)])
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		switch {
		case strings.HasPrefix(lower, "#if "):
			conditionalDepth++
		case lower == "#end if" || lower == "#endif":
			if conditionalDepth == 0 {
				return true
			}
			conditionalDepth--
		case strings.HasPrefix(lower, "#else") || strings.HasPrefix(lower, "#elseif"):
			if conditionalDepth == 0 {
				return true
			}
		}
	}
	if conditionalDepth != 0 {
		return true
	}
	for _, procedure := range procedures {
		endLine := procedure.Range.End.Line
		if endLine <= procedure.Range.Start.Line || endLine >= len(lines) {
			return true
		}
		blockStack := make([]string, 0, 4)
		pendingIf := false
		parenthesisDepth := 0
		bodyStart := procedure.Range.Start.Line + 1
		if strings.HasSuffix(strings.TrimSpace(lines[procedure.Range.Start.Line][:codeLimit(lines[procedure.Range.Start.Line])]), "_") {
			for bodyStart < endLine {
				headerLine := strings.TrimSpace(lines[bodyStart][:codeLimit(lines[bodyStart])])
				bodyStart++
				if !strings.HasSuffix(headerLine, "_") {
					break
				}
			}
		}
		for lineNo := bodyStart; lineNo < endLine; lineNo++ {
			text := strings.TrimSpace(lines[lineNo][:codeLimit(lines[lineNo])])
			if text == "" {
				continue
			}
			if compactBodySyntaxRecovery(text, &parenthesisDepth) {
				return true
			}
			for _, statement := range splitRecoveryStatements(text) {
				lower := strings.ToLower(strings.TrimSpace(statement))
				trimmedLower := strings.TrimSpace(lower)
				if trimmedLower == "" {
					continue
				}
				if strings.HasSuffix(trimmedLower, "=") {
					// A statement that ends at an assignment operator is a parser
					// recovery node in VBA unless it explicitly continues to the next
					// line. The compact declaration source omits this body, so retain
					// the conservative full-symbol fallback for the revision.
					return true
				}
				if pendingIf {
					if strings.HasSuffix(trimmedLower, "_") {
						continue
					}
					pendingIf = false
					if trimmedLower == "then" {
						blockStack = append(blockStack, "if")
						continue
					}
					if then := strings.LastIndex(trimmedLower, " then"); then >= 0 && strings.TrimSpace(trimmedLower[then+len(" then"):]) == "" {
						blockStack = append(blockStack, "if")
						continue
					}
				}
				if strings.HasPrefix(trimmedLower, "if ") && !strings.Contains(trimmedLower, " then") && strings.HasSuffix(trimmedLower, "_") {
					pendingIf = true
					continue
				}
				switch {
				case recoveryKeyword(trimmedLower, "end if"):
					if !popRecoveryBlock(&blockStack, "if") {
						return true
					}
				case strings.HasPrefix(lower, "if ") && multilineIfLine(lower):
					blockStack = append(blockStack, "if")
				case strings.HasPrefix(lower, "for "):
					blockStack = append(blockStack, "for")
				case recoveryKeyword(trimmedLower, "next"):
					if !popRecoveryBlock(&blockStack, "for") {
						return true
					}
				case strings.HasPrefix(lower, "do") && (lower == "do" || strings.HasPrefix(lower, "do ")):
					blockStack = append(blockStack, "do")
				case recoveryKeyword(trimmedLower, "loop"):
					if !popRecoveryBlock(&blockStack, "do") {
						return true
					}
				case strings.HasPrefix(lower, "select case"):
					blockStack = append(blockStack, "select")
				case recoveryKeyword(trimmedLower, "end select"):
					if !popRecoveryBlock(&blockStack, "select") {
						return true
					}
				case strings.HasPrefix(lower, "with "):
					blockStack = append(blockStack, "with")
				case recoveryKeyword(trimmedLower, "end with"):
					if !popRecoveryBlock(&blockStack, "with") {
						return true
					}
				case strings.HasPrefix(lower, "while "):
					blockStack = append(blockStack, "while")
				case recoveryKeyword(trimmedLower, "wend"):
					if !popRecoveryBlock(&blockStack, "while") {
						return true
					}
				}
			}
		}
		if len(blockStack) != 0 || pendingIf || parenthesisDepth != 0 {
			return true
		}
	}
	return false
}

func popRecoveryBlock(stack *[]string, want string) bool {
	if stack == nil || len(*stack) == 0 || (*stack)[len(*stack)-1] != want {
		return false
	}
	*stack = (*stack)[:len(*stack)-1]
	return true
}

func recoveryKeyword(text, keyword string) bool {
	return text == keyword || strings.HasPrefix(text, keyword+" ") || strings.HasPrefix(text, keyword+":")
}

func splitRecoveryStatements(text string) []string {
	statements := make([]string, 0, 2)
	start := 0
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case ':':
			if inString {
				continue
			}
			statements = append(statements, text[start:i])
			start = i + 1
		}
	}
	return append(statements, text[start:])
}

func compactBodySyntaxRecovery(text string, parenthesisDepth *int) bool {
	if invalidInlineDeclarationSyntax(text) {
		return true
	}
	inString := false
	questionRun := 0
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			questionRun = 0
		case '(':
			if !inString && parenthesisDepth != nil {
				*parenthesisDepth++
			}
			questionRun = 0
		case ')':
			if !inString && parenthesisDepth != nil {
				*parenthesisDepth--
				if *parenthesisDepth < 0 {
					return true
				}
			}
			questionRun = 0
		case '?':
			if !inString {
				questionRun++
				if questionRun >= 2 {
					return true
				}
			}
		default:
			questionRun = 0
		}
	}
	return inString
}

func invalidInlineDeclarationSyntax(text string) bool {
	for _, statement := range splitRecoveryStatements(text) {
		lower := strings.ToLower(strings.TrimSpace(statement))
		for _, keyword := range []string{"dim", "static", "const"} {
			if !strings.HasPrefix(lower, keyword+" ") && lower != keyword {
				continue
			}
			for _, part := range strings.Split(lower, ",")[1:] {
				part = strings.TrimSpace(part)
				if part == keyword || strings.HasPrefix(part, keyword+" ") {
					return true
				}
			}
		}
	}
	return false
}

func multilineIfLine(lower string) bool {
	then := strings.LastIndex(lower, " then")
	if then < 0 {
		return false
	}
	// A block If has no statement after Then. A single-line If may continue
	// its statement with an underscore, so the trailing continuation marker
	// alone must not make it look like a block opener.
	return strings.TrimSpace(lower[then+len(" then"):]) == ""
}

func rebaseDeclarationFileResult(file *symbols.FileResult, lineMap []int) {
	if file == nil || len(lineMap) == 0 {
		return
	}
	mapLine := func(oneBased int) int {
		index := oneBased - 1
		if index < 0 {
			return oneBased
		}
		if index >= len(lineMap) {
			index = len(lineMap) - 1
		}
		return lineMap[index] + 1
	}
	for i := range file.Symbols {
		file.Symbols[i].StartLine = mapLine(file.Symbols[i].StartLine)
		file.Symbols[i].EndLine = mapLine(file.Symbols[i].EndLine)
		if file.Symbols[i].DocStartLine > 0 {
			file.Symbols[i].DocStartLine = mapLine(file.Symbols[i].DocStartLine)
		}
	}
}

func interactiveIndexForDocument(doc Document) (*interactiveDocumentIndex, bool, func(), error) {
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		idx, initialized, err := snapshot.buildInteractiveIndex()
		return idx, initialized, func() {}, err
	}
	snapshot := NewAnalysisSnapshot(doc)
	idx, initialized, err := snapshot.buildInteractiveIndex()
	return idx, initialized, snapshot.Retire, err
}
