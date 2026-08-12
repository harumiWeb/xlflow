package analyze

import (
	"fmt"
	"regexp"
	"sort"
	"strings"
)

// filePathValue is deliberately small and procedure-local. It mirrors the
// conservative value facts used by the shared data-flow package while keeping
// path-specific classifications out of the generic sink catalog.
type filePathValue struct {
	raw       string
	constant  string
	origin    string // clean, tainted, unknown
	anchor    string
	temporary bool
	known     bool
}

type fileOperationUse struct {
	operation string
	paths     []fileOperationPath
	overwrite *bool
	column    int
}

type fileOperationPath struct {
	role string
	expr string
}

var (
	fileKillRe        = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:vba\s*\.\s*)?kill\b`)
	fileRmDirRe       = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:vba\s*\.\s*)?rmdir\b`)
	fileNameRe        = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:vba\s*\.\s*)?name\b`)
	fileCopyRe        = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:vba\s*\.\s*)?filecopy\b`)
	fileOpenRe        = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:vba\s*\.\s*)?open\b`)
	fileMethodRe      = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_\.\(\)]*|createobject\(\s*"[^"]+"\s*\))\s*\.\s*(deletefile|deletefolder|copyfile|copyfolder|movefile|movefolder|createtextfile|opentextfile)\b`)
	fileSaveMethodRe  = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_\.\(\)]*)\s*\.\s*(saveas|savecopyas)\b`)
	fileRootRe        = regexp.MustCompile(`(?i)^(?:[a-z]:[\\/]?|[\\/]{1,2})$`)
	fileDriveRe       = regexp.MustCompile(`(?i)^[a-z]:`)
	fileUNCRe         = regexp.MustCompile(`^\\\\[^\\]+\\[^\\]+(?:\\|/)?$`)
	fileTraversalRe   = regexp.MustCompile(`(^|[\\/])\.\.($|[\\/])`)
	fileNameAsRe      = regexp.MustCompile(`(?i)^(.+?)\s+as\s+(.+)$`)
	fileBinaryOpenRe  = regexp.MustCompile(`(?i)^\s*open\s+(.+?)\s+for\s+binary\b.*?\bas\s*#\s*([0-9]+)`)
	fileOpenHandleRe  = regexp.MustCompile(`(?i)^\s*(?:call\s+)?open\b.*?\bas\s*#\s*([0-9]+)`)
	fileBinaryWriteRe = regexp.MustCompile(`(?i)^\s*(?:put|write)\s*#\s*([0-9]+)\b`)
	fileCloseRe       = regexp.MustCompile(`(?i)^\s*(?:call\s+)?close\b(.*)$`)
	fileTempSegmentRe = regexp.MustCompile(`(?i)(^|[\\/])(temp|tmp)([\\/]|$)`)
)

func (a Analyzer) filePathSafetyFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectUnsafeFilePath {
		return nil
	}
	values := map[string]filePathValue{}
	// Module/procedure Const declarations are clean symbolic values even when
	// they do not appear as assignments inside the current procedure.
	for _, line := range file.Lines {
		if name, expr, ok := fileConstDeclaration(line); ok {
			values[strings.ToLower(name)] = classifyFilePathExpr(expr, values, nil)
		}
	}
	values["vbnullstring"] = filePathValue{raw: "vbNullString", constant: "", origin: "clean", known: true}
	params := map[string]bool{}
	for _, p := range proc.Params {
		params[strings.ToLower(p.Name)] = true
		values[strings.ToLower(p.Name)] = filePathValue{raw: p.Name, origin: "tainted", known: false}
	}
	localProcedures := map[string]bool{}
	for _, p := range file.IR.Procedures {
		localProcedures[strings.ToLower(p.Symbol.Name)] = true
	}
	fsos := map[string]bool{}
	workbooks := map[string]bool{}
	for _, d := range proc.Declarations {
		name := strings.ToLower(strings.TrimSpace(d.Name))
		if strings.Contains(strings.ToLower(d.Type), "filesystemobject") || name == "fso" || name == "fs" || name == "filesystemobject" {
			fsos[strings.ToLower(d.Name)] = true
		}
		if strings.Contains(strings.ToLower(d.Type), "workbook") {
			workbooks[strings.ToLower(d.Name)] = true
		}
	}
	temps := map[string]int{}
	cleaned := map[string]bool{}
	cleanupLines := map[string][]int{}
	binaryHandles := map[string]fileOperationPath{}
	guardedDestinations := fileGuardedDestinations(file.Lines, proc.StartLine, proc.EndLine)
	var findings []Finding
	seen := map[string]bool{}
	for lineNo := proc.StartLine; lineNo <= proc.EndLine && lineNo <= len(file.Lines); lineNo++ {
		raw := stripVBAFileComment(file.Lines[lineNo-1])
		stmt := strings.TrimSpace(raw)
		if stmt == "" {
			continue
		}
		lower := strings.ToLower(stmt)
		// Track simple aliases and path construction before inspecting the sink
		// on the same line. This is enough for the common `p = root & name`
		// pattern while retaining unknown provenance for unsupported expressions.
		if name, expr, ok := fileAssignment(stmt); ok {
			value := classifyFilePathExpr(expr, values, params)
			values[strings.ToLower(name)] = value
		}
		if strings.Contains(lower, "createobject(\"scripting.filesystemobject\")") {
			if name, _, ok := fileAssignment(stmt); ok {
				fsos[strings.ToLower(name)] = true
			}
		}
		// File numbers are reusable. Any Open assignment invalidates the old
		// binary mapping; a binary Open then installs the current path.
		if match := fileOpenHandleRe.FindStringSubmatch(stmt); len(match) == 2 {
			delete(binaryHandles, match[1])
		}
		if match := fileBinaryOpenRe.FindStringSubmatch(stmt); len(match) == 3 {
			binaryHandles[match[2]] = fileOperationPath{role: "target", expr: strings.TrimSpace(match[1])}
		}
		if match := fileCloseRe.FindStringSubmatch(stmt); len(match) == 2 {
			clearClosedBinaryHandles(match[1], binaryHandles)
		}
		uses := fileOperationUses(stmt, localProcedures, fsos, values, workbooks)
		if match := fileBinaryWriteRe.FindStringSubmatch(stmt); len(match) == 2 {
			if path, ok := binaryHandles[match[1]]; ok {
				uses = append(uses, fileOperationUse{operation: "Open For Binary", paths: []fileOperationPath{path}})
			}
		}
		for _, use := range uses {
			if len(use.paths) >= 2 && (use.operation == "Name" || strings.Contains(strings.ToLower(use.operation), "copy") || strings.Contains(strings.ToLower(use.operation), "move")) {
				left := classifyFilePathExpr(use.paths[0].expr, values, params)
				right := classifyFilePathExpr(use.paths[1].expr, values, params)
				if filePathEquivalent(left, right) {
					key := fileFindingKey(lineNo, use, fileOperationPath{role: "destination", expr: use.paths[1].expr}, "same_source_destination")
					if seen[key] {
						continue
					}
					seen[key] = true
					finding := a.fileOperationFinding(file, proc, lineNo, use, fileOperationPath{role: "destination", expr: use.paths[1].expr}, right, "definite", "same_source_destination")
					findings = append(findings, finding)
					continue
				}
			}
			for _, path := range use.paths {
				value := classifyFilePathExpr(path.expr, values, params)
				if value.temporary && isFileCreationOperation(use.operation) {
					if key := fileValueKey(path.expr); key != "" {
						temps[key] = lineNo
					}
				}
				if value.temporary && isFileCleanupOperation(use.operation) {
					if key := fileValueKey(path.expr); key != "" {
						cleaned[key] = true
						cleanupLines[key] = append(cleanupLines[key], lineNo)
					}
				}
				overwrite := use.overwrite
				if path.role == "destination" && guardedDestinations[fileValueKey(path.expr)] {
					guarded := false
					overwrite = &guarded
				}
				risk, class := filePathRisk(use.operation, path.role, value, overwrite)
				if risk == "" {
					continue
				}
				key := fileFindingKey(lineNo, use, path, risk)
				if seen[key] {
					continue
				}
				seen[key] = true
				findingUse := use
				findingUse.overwrite = overwrite
				finding := a.fileOperationFinding(file, proc, lineNo, findingUse, path, value, class, risk)
				finding.Column = max(0, strings.Index(strings.ToLower(raw), strings.ToLower(operationToken(use.operation))))
				findings = append(findings, finding)
			}
		}
	}
	// Report a temporary path only when no matching cleanup operation exists
	// anywhere in this procedure. This is intentionally conservative and does
	// not claim interprocedural cleanup summaries.
	tempKeys := make([]string, 0, len(temps))
	for key := range temps {
		tempKeys = append(tempKeys, key)
	}
	sort.Strings(tempKeys)
	for _, key := range tempKeys {
		lineNo := temps[key]
		if lineNo <= 0 {
			continue
		}
		if cleaned[key] && !fileTemporaryCleanupMaySkip(file.Lines, lineNo, cleanupLines[key]) {
			continue
		}
		if fileFindingAtLine(findings, lineNo, "temporary_cleanup_missing") {
			continue
		}
		use := fileOperationUse{operation: "temporary-file creation", column: 0}
		path := fileOperationPath{role: "target", expr: key}
		findingKey := fileFindingKey(lineNo, use, path, "temporary_cleanup_missing")
		if seen[findingKey] {
			continue
		}
		seen[findingKey] = true
		value := filePathValue{raw: key, constant: key, origin: "clean", temporary: true, known: true}
		findings = append(findings, a.fileOperationFinding(file, proc, lineNo, use, path, value, "lifecycle", "temporary_cleanup_missing"))
	}
	return findings
}

func fileFindingKey(line int, use fileOperationUse, path fileOperationPath, risk string) string {
	return fmt.Sprintf("%d|%s|%s|%s|%s", line, strings.ToLower(use.operation), strings.ToLower(path.role), strings.TrimSpace(path.expr), risk)
}

func (a Analyzer) fileOperationFinding(file parsedFile, proc sourceProcedure, line int, use fileOperationUse, path fileOperationPath, value filePathValue, class, risk string) Finding {
	operation := use.operation
	message := fmt.Sprintf("Potential unsafe %s path (%s).", operation, risk)
	if class == "definite" {
		message = fmt.Sprintf("Unsafe %s path: %s.", operation, risk)
	}
	reason := fmt.Sprintf("The %s operation receives a %s path classification.", operation, risk)
	switch value.origin {
	case "tainted":
		reason += " The path contains procedure input or external data."
	case "unknown":
		reason += " The path origin cannot be proven safe by the procedure-local analysis."
	}
	suggestion := filePathSuggestion(risk)
	finding := a.simpleFinding(file, proc, line, "VBA245", "warning", message, reason, suggestion)
	finding.FileOperation = &FileOperationContext{
		Operation: operation, PathRole: path.role, RiskClass: class, RiskKind: risk,
		OriginState: value.origin, Anchor: value.anchor, Overwrite: use.overwrite,
	}
	return finding
}

func filePathSuggestion(risk string) string {
	switch risk {
	case "empty_path", "root_path":
		return "Reject empty and root paths before the destructive operation; require a non-empty file name under a trusted root."
	case "wildcard_delete":
		return "Do not pass wildcards to deletion APIs; enumerate exact files and validate each path before deleting."
	case "relative_path", "current_directory_dependency":
		return "Resolve the path against ThisWorkbook.Path or a configured project root instead of relying on CurDir or a relative path."
	case "directory_traversal":
		return "Reject . and .. path segments and verify the canonical path remains under the trusted project root."
	case "unchecked_overwrite":
		return "Check destination existence and choose an explicit overwrite policy before writing or saving."
	case "same_source_destination":
		return "Compare normalized source and destination paths and reject a same-file copy or rename."
	case "untrusted_filename":
		return "Allow only a validated leaf name, reject separators and .., then join it to a trusted project path."
	case "temporary_cleanup_missing":
		return "Close and delete temporary files from a shared cleanup path on success, error, and early exits."
	default:
		return "Validate the path, reject traversal and wildcards, and enforce containment under a trusted project root before the operation."
	}
}

func filePathRisk(operation, _ string, value filePathValue, overwrite *bool) (string, string) {
	if value.known && value.constant == "" {
		if value.anchor != "" {
			return "root_path", "definite"
		}
		return "empty_path", "definite"
	}
	if value.known && (fileRootRe.MatchString(value.constant) || fileUNCRe.MatchString(value.constant)) {
		return "root_path", "definite"
	}
	if value.known && fileTraversalRe.MatchString(value.constant) {
		return "directory_traversal", "definite"
	}
	if value.known && (strings.Contains(value.constant, "*") || strings.Contains(value.constant, "?")) && isFileDeleteOperation(operation) {
		return "wildcard_delete", "definite"
	}
	if !value.known && (strings.Contains(value.raw, "*") || strings.Contains(value.raw, "?")) && isFileDeleteOperation(operation) {
		return "wildcard_delete", "input_dependent"
	}
	driveRelative := fileDriveRe.MatchString(value.constant) && len(value.constant) > 2 && value.constant[2] != '\\' && value.constant[2] != '/'
	if value.origin == "clean" && value.known && value.anchor == "" && (driveRelative || (!fileDriveRe.MatchString(value.constant) && !strings.HasPrefix(value.constant, "\\") && !strings.HasPrefix(value.constant, "/"))) {
		return "relative_path", "definite"
	}
	if value.anchor == "" && strings.Contains(strings.ToLower(value.raw), "curdir") {
		return "current_directory_dependency", "definite"
	}
	if overwrite != nil && *overwrite {
		if value.origin == "clean" {
			return "unchecked_overwrite", "definite"
		}
		return "unchecked_overwrite", "input_dependent"
	}
	// An explicit False overwrite argument is a proof obligation already
	// discharged by the caller.  Only implicit/default overwrite behavior is
	// unsafe; do not re-report FSO Copy*/CreateTextFile calls that pass False.
	if overwrite == nil && isImplicitOverwriteOperation(operation) {
		if value.origin == "clean" {
			return "unchecked_overwrite", "definite"
		}
		return "unchecked_overwrite", "input_dependent"
	}
	// GetTempName/unique temporary-folder origins are safe destinations in
	// themselves; the lifecycle check below still requires cleanup on every
	// local exit.
	if value.temporary && value.origin == "unknown" && isFileCreationOperation(operation) {
		return "", ""
	}
	switch value.origin {
	case "tainted":
		if value.anchor != "" {
			return "untrusted_filename", "input_dependent"
		}
		return "unknown_path", "input_dependent"
	case "unknown":
		return "unknown_path", "input_dependent"
	}
	return "", ""
}

func filePathEquivalent(left, right filePathValue) bool {
	if left.known && right.known {
		return strings.EqualFold(strings.TrimSpace(left.constant), strings.TrimSpace(right.constant))
	}
	return strings.EqualFold(strings.TrimSpace(left.raw), strings.TrimSpace(right.raw)) && strings.TrimSpace(left.raw) != ""
}

func fileOperationUses(stmt string, localProcedures map[string]bool, fsos map[string]bool, values map[string]filePathValue, workbooks map[string]bool) []fileOperationUse {
	lower := strings.ToLower(stmt)
	var out []fileOperationUse
	addBuiltin := func(re *regexp.Regexp, operation string, count int) bool {
		loc := re.FindStringIndex(stmt)
		if loc == nil {
			return false
		}
		name := strings.ToLower(strings.TrimSpace(stmt[loc[0]:loc[1]]))
		name = strings.TrimSpace(strings.TrimPrefix(name, "call"))
		name = strings.TrimSpace(strings.TrimPrefix(name, "vba."))
		name = strings.ToLower(lastName(name))
		if localProcedures[name] {
			return false
		}
		rest := strings.TrimSpace(stmt[loc[1]:])
		if strings.HasPrefix(rest, "(") {
			rest = strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(rest, "("), ")"))
		}
		args := splitArgs(strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(rest), "Call ")))
		if len(args) < count {
			return false
		}
		paths := make([]fileOperationPath, 0, count)
		for i := 0; i < count; i++ {
			paths = append(paths, fileOperationPath{role: "target", expr: args[i]})
		}
		out = append(out, fileOperationUse{operation: operation, paths: paths, column: loc[0]})
		return true
	}
	if addBuiltin(fileKillRe, "Kill", 1) || addBuiltin(fileRmDirRe, "RmDir", 1) || addBuiltin(fileCopyRe, "FileCopy", 2) {
		return out
	}
	if m := fileNameRe.FindStringIndex(stmt); m != nil {
		rest := strings.TrimSpace(stmt[m[1]:])
		if parts := fileNameAsRe.FindStringSubmatch(rest); len(parts) == 3 {
			out = append(out, fileOperationUse{operation: "Name", paths: []fileOperationPath{{role: "source", expr: strings.TrimSpace(parts[1])}, {role: "destination", expr: strings.TrimSpace(parts[2])}}, column: m[0]})
			return out
		}
	}
	if m := fileOpenRe.FindStringIndex(stmt); m != nil && !strings.Contains(lower, " for input") && !strings.Contains(lower, " for random") {
		rest := strings.TrimSpace(stmt[m[1]:])
		args := strings.Fields(strings.ToLower(rest))
		mode := ""
		for i, arg := range args {
			if arg == "for" && i+1 < len(args) {
				mode = args[i+1]
				break
			}
		}
		if mode == "output" || mode == "append" {
			pathEnd := strings.Index(strings.ToLower(rest), " for ")
			if pathEnd > 0 {
				out = append(out, fileOperationUse{operation: "Open For " + capitalizeASCII(mode), paths: []fileOperationPath{{role: "target", expr: strings.TrimSpace(rest[:pathEnd])}}, column: m[0]})
				return out
			}
		}
	}
	if m := fileMethodRe.FindStringSubmatchIndex(stmt); m != nil {
		receiver := strings.TrimSpace(stmt[m[2]:m[3]])
		method := strings.ToLower(stmt[m[4]:m[5]])
		base := strings.ToLower(lastName(strings.TrimSuffix(receiver, "()")))
		if !fsos[base] && !strings.Contains(strings.ToLower(receiver), "filesystemobject") && !strings.Contains(strings.ToLower(receiver), "createfilesystemobject") && !strings.Contains(strings.ToLower(receiver), "createobject(\"scripting.filesystemobject\")") {
			return out
		}
		args := fileCallArgs(stmt, m[1])
		count := 1
		if strings.HasPrefix(method, "copy") || strings.HasPrefix(method, "move") {
			count = 2
		}
		// Named FSO arguments are common in VBA and may appear in any order.
		// Normalize them to the positional path roles before classifying the
		// operation; this also lets an explicit OverwriteFiles:=False discharge
		// the overwrite risk.
		if named := fileNamedArguments(args); len(named) > 0 {
			normalized, namedOverwrite, mode := normalizeFSOArguments(method, args, named)
			if method == "opentextfile" && mode == "reading" {
				return out
			}
			if len(normalized) >= count {
				args = normalized
				var overwrite *bool
				if namedOverwrite != nil {
					overwrite = namedOverwrite
				}
				paths := make([]fileOperationPath, 0, count)
				for i := 0; i < count; i++ {
					role := "target"
					if count == 2 && i == 0 {
						role = "source"
					} else if count == 2 {
						role = "destination"
					}
					paths = append(paths, fileOperationPath{role: role, expr: args[i]})
				}
				out = append(out, fileOperationUse{operation: method, paths: paths, overwrite: overwrite, column: m[0]})
				return out
			}
		}
		if method == "opentextfile" && len(args) > 1 {
			if fileTextMode(args[1]) == "reading" {
				return out
			}
		}
		if len(args) < count {
			return out
		}
		paths := make([]fileOperationPath, 0, count)
		for i := 0; i < count; i++ {
			role := "target"
			if count == 2 && i == 0 {
				role = "source"
			} else if count == 2 {
				role = "destination"
			}
			paths = append(paths, fileOperationPath{role: role, expr: args[i]})
		}
		var overwrite *bool
		if len(args) > count {
			if b, ok := parseFileBool(args[count]); ok {
				overwrite = &b
			}
		}
		out = append(out, fileOperationUse{operation: method, paths: paths, overwrite: overwrite, column: m[0]})
		return out
	}
	if m := fileSaveMethodRe.FindStringSubmatchIndex(stmt); m != nil {
		receiver := strings.ToLower(strings.TrimSpace(stmt[m[2]:m[3]]))
		method := strings.ToLower(stmt[m[4]:m[5]])
		if !isWorkbookReceiver(receiver, values, workbooks) {
			return out
		}
		args := fileCallArgs(stmt, m[1])
		if len(args) == 0 {
			return out
		}
		path := args[0]
		for _, arg := range args {
			parts := strings.SplitN(arg, ":=", 2)
			if len(parts) != 2 {
				continue
			}
			key := strings.ToLower(strings.TrimSpace(parts[0]))
			if key == "filename" || key == "name" {
				path = strings.TrimSpace(parts[1])
				break
			}
		}
		out = append(out, fileOperationUse{operation: method, paths: []fileOperationPath{{role: "destination", expr: strings.TrimSpace(path)}}, column: m[0]})
	}
	return out
}

func fileNamedArguments(args []string) map[string]string {
	result := map[string]string{}
	for _, arg := range args {
		key, value, ok := fileNamedArgument(arg)
		if !ok {
			continue
		}
		result[key] = value
	}
	return result
}

func fileNamedArgument(arg string) (string, string, bool) {
	parts := strings.SplitN(arg, ":=", 2)
	if len(parts) != 2 {
		return "", "", false
	}
	key := strings.ToLower(strings.TrimSpace(parts[0]))
	if key == "" {
		return "", "", false
	}
	return key, strings.TrimSpace(parts[1]), true
}

func normalizeFSOArguments(method string, args []string, named map[string]string) ([]string, *bool, string) {
	positional := make([]string, 0, len(args))
	for _, arg := range args {
		if _, _, ok := fileNamedArgument(arg); !ok {
			positional = append(positional, strings.TrimSpace(arg))
		}
	}
	path := func(position int, keys ...string) string {
		for _, key := range keys {
			if value, ok := named[key]; ok {
				return value
			}
		}
		if position >= 0 && position < len(positional) {
			return positional[position]
		}
		return ""
	}
	var overwrite *bool
	setOverwrite := func(expr string) {
		if value, ok := parseFileBool(expr); ok {
			overwrite = &value
		}
	}
	if value := path(-1, "overwritefiles", "overwrite", "replace"); value != "" {
		setOverwrite(value)
	}
	pathValues := func(roles ...[]string) ([]string, []string) {
		result := make([]string, len(roles))
		remaining := append([]string(nil), positional...)
		for i, keys := range roles {
			for _, key := range keys {
				if value, ok := named[key]; ok {
					result[i] = value
					break
				}
			}
		}
		for i := range result {
			if result[i] == "" && len(remaining) > 0 {
				result[i] = remaining[0]
				remaining = remaining[1:]
			}
		}
		return result, remaining
	}
	switch method {
	case "deletefile", "deletefolder":
		paths, remaining := pathValues([]string{"pathname", "path", "filename", "name"})
		if value := paths[0]; value != "" {
			if overwrite == nil && len(remaining) > 0 {
				setOverwrite(remaining[0])
			}
			return []string{value}, overwrite, ""
		}
	case "copyfile", "copyfolder", "movefile", "movefolder":
		paths, remaining := pathValues(
			[]string{"source", "sourcefile", "sourcefolder"},
			[]string{"destination", "destinationfile", "destinationfolder", "dest"},
		)
		source, destination := paths[0], paths[1]
		if source != "" && destination != "" {
			if overwrite == nil && len(remaining) > 0 {
				setOverwrite(remaining[0])
			}
			return []string{source, destination}, overwrite, ""
		}
	case "createtextfile":
		paths, remaining := pathValues([]string{"filename", "pathname", "path", "name"})
		if value := paths[0]; value != "" {
			if overwrite == nil && len(remaining) > 0 {
				setOverwrite(remaining[0])
			}
			return []string{value}, overwrite, ""
		}
	case "opentextfile":
		paths, remaining := pathValues([]string{"filename", "pathname", "path", "name"})
		filename := paths[0]
		modeExpr := path(-1, "iomode", "mode")
		if modeExpr == "" && len(remaining) > 0 {
			modeExpr = remaining[0]
		}
		mode := fileTextMode(modeExpr)
		if filename != "" {
			return []string{filename}, overwrite, mode
		}
	}
	return args, overwrite, fileTextMode("")
}

func fileTextMode(expr string) string {
	lower := strings.ToLower(strings.TrimSpace(expr))
	switch lower {
	case "1", "forreading", "iomode:=forreading":
		return "reading"
	case "2", "forwriting", "iomode:=forwriting":
		return "writing"
	case "8", "forappending", "iomode:=forappending":
		return "appending"
	default:
		return "unknown"
	}
}

func classifyFilePathExpr(expr string, values map[string]filePathValue, params map[string]bool) filePathValue {
	raw := strings.TrimSpace(expr)
	for strings.HasPrefix(strings.ToLower(raw), "byval ") || strings.HasPrefix(strings.ToLower(raw), "byref ") {
		raw = strings.TrimSpace(raw[strings.Index(raw, " ")+1:])
	}
	for len(raw) >= 2 && strings.HasPrefix(raw, "(") && strings.HasSuffix(raw, ")") {
		raw = strings.TrimSpace(raw[1 : len(raw)-1])
	}
	value := filePathValue{raw: raw, origin: "unknown"}
	if raw == "" {
		value.known, value.constant, value.origin = true, "", "clean"
		return value
	}
	if literal, ok := fileStringLiteral(raw); ok {
		value.known, value.constant, value.origin = true, literal, "clean"
		value.temporary = isTemporaryPathText(literal)
		return value
	}
	lower := strings.ToLower(raw)
	if args := fileBuildPathArgs(raw); len(args) >= 2 {
		base := classifyFilePathExpr(args[0], values, params)
		leaf := classifyFilePathExpr(args[1], values, params)
		joined := filePathValue{raw: raw, origin: "clean", known: base.known && leaf.known, anchor: base.anchor, temporary: base.temporary || leaf.temporary}
		if base.origin == "tainted" || leaf.origin == "tainted" {
			joined.origin = "tainted"
		} else if base.origin == "unknown" || leaf.origin == "unknown" {
			joined.origin = "unknown"
		}
		if joined.known {
			joined.constant = strings.TrimRight(base.constant, "\\/") + "\\" + strings.TrimLeft(leaf.constant, "\\/")
		}
		return joined
	}
	if strings.Contains(lower, "thisworkbook.path") {
		value.anchor, value.origin, value.known = "ThisWorkbook.Path", "clean", true
	}
	if strings.Contains(lower, "application.path") || strings.Contains(lower, "project.path") || strings.Contains(lower, "projectroot") || strings.Contains(lower, "projectpath") || strings.Contains(lower, "srcroot") || strings.Contains(lower, "sourceroot") || strings.Contains(lower, "workbookdirectory") || strings.Contains(lower, "settingsdirectory") {
		value.anchor, value.origin, value.known = "configured project path", "clean", true
	}
	if strings.Contains(lower, "curdir") || strings.Contains(lower, "chdir") {
		value.origin, value.known = "unknown", false
		return value
	}
	if strings.Contains(lower, "inputbox") || strings.Contains(lower, "range(") || strings.Contains(lower, "cells(") || strings.Contains(lower, ".value") || strings.Contains(lower, "environ") || strings.Contains(lower, "command$") || strings.Contains(lower, "responsetext") || strings.Contains(lower, "field(") || strings.Contains(lower, "line input") || strings.Contains(lower, "input #") || strings.Contains(lower, "readall") || strings.Contains(lower, "readtext") {
		value.origin, value.known = "tainted", false
	}
	if isTemporaryPathText(raw) {
		value.temporary = true
	}
	if strings.Contains(raw, "&") {
		parts := splitTopLevelAmpersand(raw)
		if len(parts) > 1 {
			joined := filePathValue{raw: raw, origin: "clean", known: true}
			for _, part := range parts {
				item := classifyFilePathExpr(part, values, params)
				if item.origin == "tainted" {
					joined.origin = "tainted"
				} else if item.origin == "unknown" && joined.origin == "clean" {
					joined.origin = "unknown"
				}
				if item.anchor != "" {
					joined.anchor = item.anchor
				}
				joined.temporary = joined.temporary || item.temporary
				if !item.known {
					joined.known = false
				} else {
					joined.constant += item.constant
				}
			}
			return joined
		}
	}
	if key := strings.ToLower(cleanIdentifier(raw)); key != "" {
		if prior, ok := values[key]; ok {
			prior.raw = raw
			return prior
		}
		if params[key] {
			value.origin = "tainted"
		}
	}
	return value
}

func fileBuildPathArgs(raw string) []string {
	lower := strings.ToLower(raw)
	idx := strings.Index(lower, "buildpath")
	if idx < 0 {
		return nil
	}
	open := strings.Index(raw[idx:], "(")
	if open < 0 {
		return nil
	}
	text := strings.TrimSpace(raw[idx+open+1:])
	text = strings.TrimSpace(strings.TrimSuffix(text, ")"))
	return splitArgs(text)
}

func fileAssignment(stmt string) (string, string, bool) {
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(stmt)), "set ") {
		stmt = strings.TrimSpace(stmt[4:])
	}
	idx := strings.Index(stmt, "=")
	if idx <= 0 || strings.Contains(strings.ToLower(stmt[:idx]), "if ") {
		return "", "", false
	}
	name := cleanIdentifier(strings.TrimSpace(stmt[:idx]))
	if strings.ContainsAny(name, " .()") || name == "" {
		return "", "", false
	}
	return name, strings.TrimSpace(stmt[idx+1:]), true
}

func fileConstDeclaration(line string) (string, string, bool) {
	line = stripVBAFileComment(line)
	fields := strings.Fields(line)
	if len(fields) < 4 {
		return "", "", false
	}
	index := 0
	for index < len(fields) {
		keyword := strings.ToLower(fields[index])
		if keyword == "public" || keyword == "private" || keyword == "friend" || keyword == "static" {
			index++
			continue
		}
		if keyword != "const" || index+2 >= len(fields) {
			return "", "", false
		}
		name := cleanIdentifier(fields[index+1])
		rest := strings.TrimSpace(line)
		constIndex := strings.Index(strings.ToLower(rest), "const")
		equal := strings.Index(rest[constIndex+len("const"):], "=")
		if name == "" || equal < 0 {
			return "", "", false
		}
		equal += constIndex + len("const")
		return name, strings.TrimSpace(rest[equal+1:]), true
	}
	return "", "", false
}

func fileStringLiteral(text string) (string, bool) {
	text = strings.TrimSpace(text)
	if len(text) < 2 || text[0] != '"' || text[len(text)-1] != '"' {
		return "", false
	}
	return strings.ReplaceAll(text[1:len(text)-1], "\"\"", "\""), true
}

func splitTopLevelAmpersand(text string) []string {
	var out []string
	start, depth := 0, 0
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
			} else {
				inString = !inString
			}
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '&':
			if !inString && depth == 0 {
				out = append(out, strings.TrimSpace(text[start:i]))
				start = i + 1
			}
		}
	}
	if len(out) > 0 {
		out = append(out, strings.TrimSpace(text[start:]))
	}
	return out
}

func parseFileBool(text string) (bool, bool) {
	text = strings.ToLower(strings.TrimSpace(text))
	switch text {
	case "true", "-1":
		return true, true
	case "false", "0":
		return false, true
	default:
		return false, false
	}
}

func fileCallArgs(stmt string, start int) []string {
	if start < 0 || start >= len(stmt) {
		return nil
	}
	text := strings.TrimSpace(stmt[start:])
	if strings.HasPrefix(text, "(") {
		text = strings.TrimSpace(text[1:])
	}
	text = strings.TrimSpace(strings.TrimSuffix(text, ")"))
	return splitArgs(text)
}

func isWorkbookReceiver(receiver string, values map[string]filePathValue, workbooks map[string]bool) bool {
	receiver = strings.ToLower(strings.TrimSpace(receiver))
	if receiver == "thisworkbook" || receiver == "activeworkbook" || receiver == "workbooks" || strings.Contains(receiver, "workbooks(") {
		return true
	}
	name := strings.ToLower(lastName(strings.TrimSuffix(receiver, "()")))
	if workbooks[name] {
		return true
	}
	value, ok := values[name]
	return ok && strings.Contains(strings.ToLower(value.raw), "workbook")
}

func isFileDeleteOperation(operation string) bool {
	lower := strings.ToLower(operation)
	return lower == "kill" || lower == "rmdir" || strings.Contains(lower, "delete")
}

func isFileCleanupOperation(operation string) bool {
	return isFileDeleteOperation(operation)
}

func isFileCreationOperation(operation string) bool {
	lower := strings.ToLower(operation)
	return strings.Contains(lower, "create") || strings.HasPrefix(lower, "open for")
}

func isImplicitOverwriteOperation(operation string) bool {
	lower := strings.ToLower(operation)
	return lower == "saveas" || lower == "savecopyas" || lower == "filecopy" || lower == "open for output" || lower == "open for append" || lower == "copyfile" || lower == "copyfolder" || lower == "movefile" || lower == "movefolder" || lower == "createtextfile" || lower == "opentextfile"
}

func fileValueKey(expr string) string {
	expr = strings.TrimSpace(expr)
	for len(expr) >= 2 && strings.HasPrefix(expr, "(") && strings.HasSuffix(expr, ")") {
		expr = strings.TrimSpace(expr[1 : len(expr)-1])
	}
	if literal, ok := fileStringLiteral(expr); ok {
		return literal
	}
	if strings.ContainsAny(expr, "&()") {
		return ""
	}
	return strings.ToLower(cleanIdentifier(expr))
}

// fileGuardedDestinations recognizes the narrow, local proof used by this
// rule: an existence probe dominates an immediate exit/raise/goto. It is not
// a general sanitizer summary; a destination is only marked when the guard is
// visible in this procedure's nearby source lines.
func fileGuardedDestinations(lines []string, start, end int) map[string]bool {
	guarded := map[string]bool{}
	if start < 1 {
		start = 1
	}
	if end > len(lines) {
		end = len(lines)
	}
	for lineNo := start; lineNo <= end; lineNo++ {
		line := stripVBAFileComment(lines[lineNo-1])
		lower := strings.ToLower(line)
		if !strings.Contains(lower, "if ") || (!strings.Contains(lower, "dir(") && !strings.Contains(lower, "fileexists(") && !strings.Contains(lower, "folderexists(") && !strings.Contains(lower, ".fileexists(") && !strings.Contains(lower, ".folderexists(")) {
			continue
		}
		expr := fileExistenceArgument(line)
		if expr == "" {
			continue
		}
		sameLineExit := strings.Contains(lower, " then exit") || strings.Contains(lower, " then err.raise") || strings.Contains(lower, " then goto")
		nearbyExit := sameLineExit
		if !nearbyExit {
			for next := lineNo + 1; next <= end && next <= lineNo+2; next++ {
				nextLower := strings.ToLower(stripVBAFileComment(lines[next-1]))
				if strings.Contains(nextLower, "exit ") || strings.Contains(nextLower, "err.raise") || strings.Contains(nextLower, "goto ") {
					nearbyExit = true
					break
				}
			}
		}
		if nearbyExit {
			if key := fileValueKey(expr); key != "" {
				guarded[key] = true
			}
		}
	}
	return guarded
}

func fileExistenceArgument(line string) string {
	lower := strings.ToLower(line)
	for _, token := range []string{"dir(", "fileexists(", "folderexists(", ".fileexists(", ".folderexists("} {
		idx := strings.Index(lower, token)
		if idx < 0 {
			continue
		}
		start := idx + strings.Index(token, "(") + 1
		text := line[start:]
		if close := strings.Index(text, ")"); close >= 0 {
			args := splitArgs(text[:close])
			if len(args) > 0 {
				return strings.TrimSpace(args[0])
			}
		}
	}
	return ""
}

func fileTemporaryCleanupMaySkip(lines []string, creationLine int, cleanup []int) bool {
	if len(cleanup) == 0 {
		return true
	}
	earliestCleanup := cleanup[0]
	for _, line := range cleanup[1:] {
		if line < earliestCleanup {
			earliestCleanup = line
		}
	}
	for lineNo := creationLine + 1; lineNo < earliestCleanup && lineNo <= len(lines); lineNo++ {
		lower := strings.ToLower(stripVBAFileComment(lines[lineNo-1]))
		if strings.Contains(lower, "exit sub") || strings.Contains(lower, "exit function") || strings.Contains(lower, "exit property") || strings.Contains(lower, "err.raise") {
			return true
		}
	}
	return false
}

func isTemporaryPathText(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.Contains(lower, "gettempname") {
		return true
	}
	if strings.HasSuffix(lower, ".tmp") {
		return true
	}
	return fileTempSegmentRe.MatchString(lower)
}

func clearClosedBinaryHandles(rest string, handles map[string]fileOperationPath) {
	rest = strings.TrimSpace(rest)
	if rest == "" {
		clear(handles)
		return
	}
	for _, token := range strings.Split(rest, ",") {
		fileNumber := strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(token), "#"))
		if fileNumber != "" {
			delete(handles, fileNumber)
		}
	}
}

func fileFindingAtLine(findings []Finding, line int, risk string) bool {
	for _, finding := range findings {
		if finding.Line == line && finding.FileOperation != nil && finding.FileOperation.RiskKind == risk {
			return true
		}
	}
	return false
}

func operationToken(operation string) string {
	if strings.HasPrefix(strings.ToLower(operation), "open for") {
		return "Open"
	}
	return operation
}

func capitalizeASCII(value string) string {
	if value == "" {
		return value
	}
	return strings.ToUpper(value[:1]) + value[1:]
}

func stripVBAFileComment(line string) string {
	line = strings.TrimSpace(line)
	if idx := strings.Index(strings.ToLower(line), " rem "); idx >= 0 {
		line = strings.TrimSpace(line[:idx])
	}
	if strings.HasPrefix(strings.ToLower(line), "rem ") {
		return ""
	}
	return rawWorksheetCodeLine(line)
}
