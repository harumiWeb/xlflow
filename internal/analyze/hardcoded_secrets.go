package analyze

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	secretConnectionFieldRe  = regexp.MustCompile(`(?i)(^|[;,&[:space:]])(password|pwd|account[[:space:]]*key|client[[:space:]]*secret|access[[:space:]]*token)[[:space:]]*=[[:space:]]*([^;,&]+)`)
	secretAuthSchemeRe       = regexp.MustCompile(`(?i)\b(bearer|basic)[[:space:]]+([A-Za-z0-9._~+/=-]{16,})`)
	secretURLCredentialRe    = regexp.MustCompile(`(?i)^[a-z][a-z0-9+.-]*://[^/[:space:]@:]+:([^@[:space:]/]+)@`)
	secretPrivateKeyRe       = regexp.MustCompile(`(?i)-----begin([[:space:]][a-z0-9-]+)? private key-----`)
	secretKnownTokenRe       = regexp.MustCompile(`(?i)\b(AKIA|ASIA)[0-9A-Z]{16}\b|ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}|glpat-[A-Za-z0-9_-]{20,}|xox[baprs]-[A-Za-z0-9-]{10,}|AIza[0-9A-Za-z_-]{20,}|sk-[A-Za-z0-9]{20,}`)
	secretWebhookRe          = regexp.MustCompile(`(?i)(hooks\.slack\.com/services/|discord\.com/api/webhooks/|discordapp\.com/api/webhooks/|api\.telegram\.org/bot)[^[:space:]"']+`)
	secretPercentExpansionRe = regexp.MustCompile(`(?i)%[a-z_][a-z0-9_]*%`)
)

var credentialVariableNames = []string{
	"password", "passwd", "pwd", "passphrase", "apikey", "apitoken", "accesskey",
	"accesstoken", "refreshtoken", "authtoken", "bearertoken", "clientsecret", "secretkey",
	"secret", "credential", "authorization", "privatekey", "webhooktoken", "awssecretaccesskey",
}

var credentialMetadataNameSuffixes = []string{
	"label", "prompt", "hint", "message", "caption", "field", "input",
	"url", "uri", "endpoint", "resource",
}

// hardcodedSecretFindings scans each parsed statement once. Keeping this at
// document scope lets it cover module-level Const declarations and prevents
// the same finding from being emitted once per procedure in realtime mode.
func (a Analyzer) hardcodedSecretFindings(file parsedFile, procedures []sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectHardcodedSecrets {
		return nil
	}

	seen := map[string]bool{}
	findings := make([]Finding, 0)
	add := func(line, endLine int, proc sourceProcedure) {
		if line <= 0 {
			return
		}
		if endLine < line {
			endLine = line
		}
		key := fmt.Sprintf("%d:%d", line, endLine)
		if seen[key] {
			return
		}
		seen[key] = true
		findings = append(findings, a.hardcodedSecretFinding(file, proc, line, endLine))
	}

	for _, procedure := range file.IR.Procedures {
		proc := procedureForLine(procedures, procedure.Symbol.BodyRange.StartLine, len(file.Lines))
		for _, statement := range procedure.Statements {
			if statement.Recovered || !statementHasHardcodedSecret(statement) {
				continue
			}
			add(statement.Range.StartLine, statement.Range.EndLine, proc)
		}
	}

	facts := file.moduleAnalysisFacts()
	for lineNo, rawLine := range file.Lines {
		line := lineNo + 1
		if facts != nil && facts.lineInProcedure(line) {
			continue
		}
		code := strings.TrimSpace(gui.StripComment(rawLine))
		if code == "" || !sourceLineHasHardcodedSecret(code) {
			continue
		}
		add(line, line, procedureForLine(procedures, line, len(file.Lines)))
	}

	sortFindings(findings)
	return findings
}

func statementHasHardcodedSecret(statement procedureir.Statement) bool {
	target := ""
	if statement.Target != nil {
		target = statement.Target.Text
	}
	if target == "" {
		target = assignmentTarget(statement.Text)
	}
	targetIsCredential := isCredentialVariable(target)

	if targetIsCredential && statement.Value != nil && statement.Value.Kind == procedureir.ExpressionLiteral {
		if value, ok := singleStringLiteral(statement.Value.Text); ok && !isSecretPlaceholderForTarget(value, target) {
			return true
		}
	}
	if targetIsCredential {
		if value, ok := directAssignmentStringLiteral(statement.Text); ok && !isSecretPlaceholderForTarget(value, target) {
			return true
		}
	}
	text := statement.Text
	if statement.Kind == procedureir.StatementWith && statement.Value != nil {
		text = statement.Value.Text
	}
	for _, literal := range extractStringLiterals(text) {
		if credentialLiteralEvidence(literal) {
			return true
		}
	}
	return false
}

func sourceLineHasHardcodedSecret(code string) bool {
	targetIsCredential := isCredentialVariable(assignmentTarget(code))
	if targetIsCredential {
		target := assignmentTarget(code)
		if value, ok := directAssignmentStringLiteral(code); ok && !isSecretPlaceholderForTarget(value, target) {
			return true
		}
	}
	for _, literal := range extractStringLiterals(code) {
		if credentialLiteralEvidence(literal) {
			return true
		}
	}
	return false
}

func credentialLiteralEvidence(literal string) bool {
	if isSecretPlaceholder(literal) {
		return false
	}
	for _, match := range secretConnectionFieldRe.FindAllStringSubmatch(literal, -1) {
		if len(match) > 3 && isSubstantiveSecretValue(match[3]) {
			return true
		}
	}
	for _, match := range secretAuthSchemeRe.FindAllStringSubmatch(literal, -1) {
		if len(match) > 2 && !isSecretPlaceholder(strings.TrimSpace(match[2])) {
			return true
		}
	}
	if secretURLCredentialRe.MatchString(literal) {
		matches := secretURLCredentialRe.FindStringSubmatch(literal)
		if len(matches) > 1 && !isSecretPlaceholder(matches[1]) {
			return true
		}
	}
	if secretPrivateKeyRe.MatchString(literal) || secretKnownTokenRe.MatchString(literal) || secretWebhookRe.MatchString(literal) {
		return true
	}
	return false
}

func isCredentialVariable(raw string) bool {
	name := credentialTargetName(raw)
	if name == "" {
		return false
	}
	compact := compactCredentialText(name)
	for _, suffix := range credentialMetadataNameSuffixes {
		if strings.HasSuffix(compact, suffix) {
			return false
		}
	}
	for _, candidate := range credentialVariableNames {
		candidate = strings.ReplaceAll(candidate, "_", "")
		if strings.Contains(compact, candidate) {
			return true
		}
	}
	return false
}

func credentialTargetName(raw string) string {
	name := strings.TrimSpace(raw)
	if name == "" || strings.ContainsAny(name, "()") {
		return ""
	}
	if _, member, _, ok := splitExcelMemberAccess(name); ok {
		name = member
	}
	return strings.ToLower(cleanIdentifier(name))
}

func compactCredentialText(text string) string {
	return strings.Map(func(r rune) rune {
		if r >= 'a' && r <= 'z' || r >= '0' && r <= '9' {
			return r
		}
		if r >= 'A' && r <= 'Z' {
			return r + ('a' - 'A')
		}
		return -1
	}, text)
}

func isSecretPlaceholderForTarget(raw, target string) bool {
	if isSecretPlaceholder(raw) {
		return true
	}
	if strings.EqualFold(strings.TrimSpace(raw), "abc") {
		return true
	}
	targetName := credentialTargetName(target)
	return targetName != "" && compactCredentialText(raw) == compactCredentialText(targetName)
}

func isSubstantiveSecretValue(raw string) bool {
	value := strings.Trim(strings.TrimSpace(raw), `"'`)
	return value != "" && !isSecretPlaceholder(value)
}

func isSecretPlaceholder(raw string) bool {
	value := strings.ToLower(strings.TrimSpace(raw))
	if value == "" {
		return true
	}
	if strings.Contains(value, "${") || secretPercentExpansionRe.MatchString(value) || strings.Contains(value, "<") && strings.Contains(value, ">") {
		return true
	}
	if strings.HasPrefix(value, "your-") || strings.HasPrefix(value, "your_") || strings.HasPrefix(value, "replace-") || strings.HasPrefix(value, "replace_") {
		return true
	}
	if strings.Contains(value, "example.com") || strings.Contains(value, "changeme") || strings.Contains(value, "change-me") {
		return true
	}
	tokens := strings.FieldsFunc(value, func(r rune) bool {
		return r < 'a' || r > 'z'
	})
	for _, token := range tokens {
		switch token {
		case "test", "testing", "dummy", "placeholder", "sample", "example", "todo", "fixme", "xxxx", "xxxxx":
			return true
		}
	}
	trimmed := strings.Trim(value, "*x")
	return trimmed == ""
}

func assignmentTarget(statement string) string {
	eq := indexOutsideString(statement, '=')
	if eq < 0 {
		return ""
	}
	left := strings.TrimSpace(statement[:eq])
	for {
		lower := strings.ToLower(left)
		removed := false
		for _, prefix := range []string{"set ", "let ", "const ", "public ", "private ", "friend ", "global ", "static ", "dim "} {
			if strings.HasPrefix(lower, prefix) {
				left = strings.TrimSpace(left[len(prefix):])
				removed = true
				break
			}
		}
		if !removed {
			break
		}
	}
	if as := strings.Index(strings.ToLower(left), " as "); as >= 0 {
		left = strings.TrimSpace(left[:as])
	}
	if strings.ContainsAny(left, "().") {
		return ""
	}
	fields := strings.Fields(left)
	if len(fields) == 0 {
		return ""
	}
	return fields[0]
}

func indexOutsideString(text string, wanted byte) int {
	inString := false
	for i := 0; i < len(text); i++ {
		if inString {
			if text[i] != '"' {
				continue
			}
			if i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = false
			continue
		}
		if text[i] == '"' {
			inString = true
			continue
		}
		if text[i] == wanted {
			return i
		}
	}
	return -1
}

func extractStringLiterals(text string) []string {
	var literals []string
	for i := 0; i < len(text); i++ {
		if text[i] != '"' {
			continue
		}
		i++
		var value strings.Builder
		for i < len(text) {
			if text[i] != '"' {
				value.WriteByte(text[i])
				i++
				continue
			}
			if i+1 < len(text) && text[i+1] == '"' {
				value.WriteByte('"')
				i += 2
				continue
			}
			literals = append(literals, value.String())
			break
		}
	}
	return literals
}

func singleStringLiteral(text string) (string, bool) {
	if len(extractStringLiterals(text)) != 1 {
		return "", false
	}
	trimmed := strings.TrimSpace(text)
	if len(trimmed) < 2 || trimmed[0] != '"' || trimmed[len(trimmed)-1] != '"' {
		return "", false
	}
	return extractStringLiterals(trimmed)[0], true
}

func directAssignmentStringLiteral(statement string) (string, bool) {
	eq := indexOutsideString(statement, '=')
	if eq < 0 {
		return "", false
	}
	return singleStringLiteral(strings.TrimSpace(statement[eq+1:]))
}

func (a Analyzer) hardcodedSecretFinding(file parsedFile, proc sourceProcedure, line, endLine int) Finding {
	finding := a.simpleFinding(
		file,
		proc,
		line,
		"VBA223",
		"warning",
		"Possible hardcoded secret or credential detected.",
		"Credential-like values embedded in VBA source can be exposed through source control, workbook distribution, or diagnostics.",
		"Load credentials from a secure external store or environment-specific configuration; suppress VBA223 only for intentional test fixtures.",
	)
	finding.NearbyCode = redactedNearbyCode(file.Lines, line, endLine, 2)
	return finding
}

func redactedNearbyCode(lines []string, startLine, endLine, radius int) []string {
	if len(lines) == 0 {
		return nil
	}
	if startLine < 1 {
		startLine = 1
	}
	if endLine < startLine {
		endLine = startLine
	}
	start := startLine - radius
	if start < 1 {
		start = 1
	}
	end := endLine + radius
	if end > len(lines) {
		end = len(lines)
	}
	out := make([]string, 0, end-start+1)
	for line := start; line <= end; line++ {
		prefix := "  "
		if line >= startLine && line <= endLine {
			prefix = "> "
		}
		out = append(out, prefix+strconvItoa(line)+" | "+redactVBAStringsAndComments(lines[line-1]))
	}
	return out
}

func redactVBAStringsAndComments(line string) string {
	var b strings.Builder
	inString := false
	for i := 0; i < len(line); i++ {
		if !inString {
			if i+3 <= len(line) && strings.EqualFold(line[i:i+3], "rem") && (i == 0 || isVBACommentBoundary(line[i-1])) && (i+3 == len(line) || isVBACommentBoundary(line[i+3])) {
				b.WriteString("Rem [REDACTED]")
				return b.String()
			}
			switch line[i] {
			case '\'':
				b.WriteString("' [REDACTED]")
				return b.String()
			case '"':
				b.WriteString("\"[REDACTED]\"")
				inString = true
			default:
				b.WriteByte(line[i])
			}
			continue
		}
		if line[i] == '"' {
			if i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = false
		}
	}
	return b.String()
}

func isVBACommentBoundary(char byte) bool {
	return char == ':' || char == ' ' || char == '\t'
}
