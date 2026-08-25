package analyze

import (
	"context"
	"fmt"
	"net"
	"net/url"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	"github.com/harumiWeb/xlflow/internal/config"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type httpClientKind string

const (
	httpUnknown   httpClientKind = ""
	httpXML       httpClientKind = "MSXML2.XMLHTTP"
	httpServerXML httpClientKind = "MSXML2.ServerXMLHTTP"
	httpWinHTTP   httpClientKind = "WinHttp.WinHttpRequest"
	httpADOStream httpClientKind = "ADODB.Stream"
)

type httpObjectState struct {
	kind              httpClientKind
	identity          string
	url               string
	urlKnown          bool
	hasCredentials    bool
	timeoutConfigured bool
	timeoutInfinite   bool
	downloaded        bool
	savedExecutable   map[string]bool
	credentialSinks   map[int]bool
}

type httpAnalysisState struct {
	objects   map[string]httpObjectState
	launchers map[string]string
	strings   map[string]string
	known     map[string]bool
	sensitive map[string]bool
}

type httpFindingSpec struct {
	line, column int
	code         string
	api          string
	risk         string
	header       string
	origin       string
	timeout      string
	redact       bool
	ownedSinks   map[int]bool
}

var (
	httpObjectAssignmentRe = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$`)
	httpMemberCallRe       = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*(open|send|setrequestheader|setcredentials|settimeouts|write|savetofile)\b`)
	httpOptionAssignmentRe = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*(?:option\s*\((.+?)\)|setoption\s*\(?\s*([^,\)]+))\s*\)?\s*=\s*(.+)$`)
	httpSetOptionCallRe    = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*setoption\b`)
	httpIdentifierRe       = regexp.MustCompile(`(?i)\b[A-Za-z_][A-Za-z0-9_]*\b`)
	httpLogRe              = regexp.MustCompile(`(?i)^\s*(?:debug\s*\.\s*print|print\s*#\s*[^,]+,?|xlflowdebug\s*\.\s*log)\b`)
	httpSensitiveHeaderRe  = regexp.MustCompile(`(?i)^(authorization|proxy-authorization|cookie|set-cookie|x-api-key|api-key|x-auth-token)$`)
	httpAuthLiteralRe      = regexp.MustCompile(`(?i)\b(?:bearer|basic)\s+[A-Za-z0-9._~+/=-]{4,}|\b(?:api[_-]?key|access[_-]?token|auth(?:orization)?[_-]?token)\s*[:=]`)
	httpExecutableExtRe    = regexp.MustCompile(`(?i)\.(exe|com|bat|cmd|ps1|vbs|js|jse|wsf|hta|msi)$`)
	httpObjectLauncherRe   = regexp.MustCompile(`(?i)\.(?:run|exec|shellexecute)\s*(?:\(|\s)`)
	httpWin32LauncherRe    = regexp.MustCompile(`(?i)^shellexecute(?:a|w)?\s*(?:\(|\s)`)
	httpWin32LauncherName  = regexp.MustCompile(`(?i)^shellexecute(?:a|w)?\b`)
	httpNewObjectRe        = regexp.MustCompile(`(?i)^\s*new\s+([A-Za-z_][A-Za-z0-9_.]*)\s*$`)
	httpCreateObjectRe     = regexp.MustCompile(`(?i)^\s*(?:vba\s*\.\s*)?createobject\s*\(\s*"([^"]+)"\s*\)\s*$`)
	httpServerXMLTypeRe    = regexp.MustCompile(`^(?:msxml2\.)?serverxmlhttp(?:(?:\.\d+\.\d+)|(?:30|40|60))?$`)
	httpWinHTTPTypeRe      = regexp.MustCompile(`^winhttp\.winhttprequest(?:\.5\.1)?$`)
	httpXMLTypeRe          = regexp.MustCompile(`^(?:msxml2\.)?xmlhttp(?:(?:\.\d+\.\d+)|(?:30|40|60))?$`)
	httpDeclarationTypeRe  = regexp.MustCompile(`(?i)\bas\s+(?:new\s+)?([A-Za-z_][A-Za-z0-9_.]*)`)
)

func (a Analyzer) httpTransportFindingsContext(ctx context.Context, file parsedFile, proc sourceProcedure) ([]Finding, error) {
	if (!a.Config.Analyze.DetectUnsafeHTTPConfiguration && !a.Config.Analyze.DetectMissingHTTPTimeout) || proc.Graph == nil {
		return nil, nil
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	ir, ok := procedureIRForSource(file.IR, proc)
	if !ok {
		return nil, nil
	}
	initial := newHTTPAnalysisState(file, ir)
	entryStates, err := solveHTTPStates(ctx, a, file, proc, *proc.Graph, initial)
	if err != nil {
		return nil, err
	}
	var specs []httpFindingSpec
	for index, block := range proc.Graph.Blocks {
		if index&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if block.Statement == nil || block.Statement.Recovered {
			continue
		}
		state, ok := entryStates[block.ID]
		if !ok {
			continue
		}
		_, found := a.transferHTTPStatement(file, proc, *block.Statement, state, true)
		specs = append(specs, found...)
	}
	// Module constants are storage findings under VBA223. VBA246 reports the
	// HTTP-specific storage policy only when the constant is actually used by a
	// sensitive request header, which avoids duplicate file-wide observations.
	sort.SliceStable(specs, func(i, j int) bool {
		if specs[i].line != specs[j].line {
			return specs[i].line < specs[j].line
		}
		return specs[i].risk < specs[j].risk
	})
	seen := map[string]bool{}
	findings := make([]Finding, 0, len(specs))
	for _, spec := range specs {
		key := fmt.Sprintf("%s|%d|%s|%s", spec.code, spec.line, spec.risk, spec.api)
		if seen[key] || (spec.code == "VBA246" && !a.Config.Analyze.DetectUnsafeHTTPConfiguration) || (spec.code == "VBA247" && !a.Config.Analyze.DetectMissingHTTPTimeout) {
			continue
		}
		seen[key] = true
		finding := a.httpFinding(file, proc, spec)
		findings = append(findings, finding)
	}
	return findings, nil
}

func newHTTPAnalysisState(file parsedFile, ir procedureir.ProcedureIR) httpAnalysisState {
	state := httpAnalysisState{objects: map[string]httpObjectState{}, launchers: map[string]string{}, strings: map[string]string{}, known: map[string]bool{}, sensitive: map[string]bool{}}
	seedDeclaration := func(declaration sourceDeclaration) {
		kind := httpKindFromText(declaration.Type)
		if kind == httpUnknown && declaration.Line > 0 && declaration.Line <= len(file.Lines) {
			if match := httpDeclarationTypeRe.FindStringSubmatch(file.Lines[declaration.Line-1]); len(match) == 2 {
				kind = httpKindFromText(match[1])
			}
		}
		if kind == httpUnknown {
			return
		}
		name := strings.ToLower(strings.TrimSpace(declaration.Name))
		if name != "" {
			state.objects[name] = newHTTPObjectState(kind, name)
		}
	}
	// The normal batch/realtime setup attaches ModuleFacts before any rule
	// worker runs. Do not rebuild facts from a standalone parsedFile on every
	// statement; retain the old IR fallback for package-local callers that do
	// not perform file setup.
	if facts := file.ModuleFacts; facts != nil {
		facts.forEachConstant(func(constant moduleConstantFact) {
			httpRecordConstant(&state, constant.Name, constant.Expression)
		})
		for _, declaration := range facts.moduleDeclarations {
			seedDeclaration(declaration)
		}
	} else {
		for _, line := range file.Lines {
			name, expr, ok := fileConstDeclaration(line)
			if ok {
				httpRecordConstant(&state, name, expr)
			}
		}
	}
	for _, declaration := range ir.Declarations {
		seedDeclaration(sourceDeclaration{Name: declaration.Name, Type: declaration.Type, Line: declaration.Range.StartLine})
	}
	return state
}

func httpRecordConstant(state *httpAnalysisState, name, expression string) {
	if state == nil {
		return
	}
	if value, known := httpConstantString(expression, *state); known {
		key := strings.ToLower(name)
		state.strings[key] = value
		state.known[key] = true
		state.sensitive[key] = httpAuthLiteralRe.MatchString(value) || credentialLiteralEvidence(value)
	} else if value, err := strconv.ParseInt(strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(expression), "&H", "0x"), "&h", "0x"), 0, 64); err == nil {
		key := strings.ToLower(name)
		state.strings[key] = strconv.FormatInt(value, 10)
		state.known[key] = true
	}
}

func solveHTTPStates(ctx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, graph vbacfg.Graph, initial httpAnalysisState) (map[vbacfg.BlockID]httpAnalysisState, error) {
	// Keep the existing HTTP state as the domain value at one indexed slot. The
	// solver owns scheduling, edge classification, cancellation, and snapshot
	// isolation; the adapter preserves the established transfer/join semantics
	// while the HTTP domain is incrementally migrated to scalar slots.
	index, err := semanticstate.NewIndex(graph)
	if err != nil {
		return nil, err
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(graph.Blocks))
	for _, block := range graph.Blocks {
		blocks[block.ID] = block
	}
	environment := semanticstate.NewEnvironment([]string{"http-state"}, []string{"http-state"})
	const lane semanticstate.LaneOrdinal = 0
	var symbol semanticstate.SymbolID
	lattice := httpStateLattice{}
	solver, err := semanticstate.NewSolver(index, environment, lattice, []semanticstate.Lane[httpAnalysisState]{
		{
			Initialize: func(_ context.Context, _ semanticstate.LaneOrdinal, state *semanticstate.State[httpAnalysisState]) error {
				state.Set(symbol, cloneHTTPState(initial))
				return nil
			},
			Transfer: func(_ context.Context, _ semanticstate.LaneOrdinal, block semanticstate.BlockOrdinal, input semanticstate.StateView[httpAnalysisState], output *semanticstate.State[httpAnalysisState]) error {
				value, ok := input.Value(symbol)
				if !ok {
					return nil
				}
				cfgBlock, ok := index.Block(block)
				if !ok {
					return nil
				}
				candidate, ok := blocks[cfgBlock.ID]
				if !ok || candidate.Statement == nil || candidate.Statement.Recovered {
					output.Set(symbol, value)
					return nil
				}
				value, _ = a.transferHTTPStatement(file, proc, *candidate.Statement, value, false)
				output.Set(symbol, value)
				return nil
			},
		},
	})
	if err != nil {
		return nil, err
	}
	result, err := solver.SolveContext(ctx)
	if err != nil {
		return nil, err
	}
	states := make(map[vbacfg.BlockID]httpAnalysisState, index.BlockCount())
	for _, block := range index.Blocks() {
		value, ok := result.State(block.Ordinal, lane).Value(symbol)
		if ok {
			states[block.ID] = value
		}
	}
	return states, nil
}

// httpStateLattice adapts the legacy HTTP state join to the indexed solver.
// Clone is the alias-safety boundary for the nested evidence sets still used
// by the HTTP domain during this migration step.
type httpStateLattice struct{}

func (httpStateLattice) Clone(value httpAnalysisState) httpAnalysisState {
	return cloneHTTPState(value)
}

func (httpStateLattice) Join(dst *httpAnalysisState, src httpAnalysisState) bool {
	if dst == nil {
		return false
	}
	merged, changed := joinHTTPState(*dst, src, true)
	if changed {
		*dst = merged
	}
	return changed
}

func (a Analyzer) transferHTTPStatement(file parsedFile, proc sourceProcedure, statement procedureir.Statement, state httpAnalysisState, collect bool) (httpAnalysisState, []httpFindingSpec) {
	state = cloneHTTPState(state)
	text := strings.TrimSpace(stripVBAFileComment(statement.Text))
	if text == "" {
		return state, nil
	}
	line := statement.Range.StartLine
	if line <= 0 {
		line = proc.StartLine
	}
	var findings []httpFindingSpec

	if match := httpObjectAssignmentRe.FindStringSubmatch(text); len(match) == 3 {
		target, rhs := strings.ToLower(match[1]), strings.TrimSpace(match[2])
		delete(state.launchers, target)
		if kind := httpKindFromConstruction(rhs); kind != httpUnknown {
			state.objects[target] = newHTTPObjectState(kind, fmt.Sprintf("%s@%d", target, statement.ID))
		} else if source := strings.ToLower(strings.TrimSpace(rhs)); state.objects[source].kind != httpUnknown {
			object := cloneHTTPObject(state.objects[source])
			if object.identity == "" {
				object.identity = source
			}
			state.objects[target] = object
		} else if launcher := httpLauncherFromConstruction(rhs); launcher != "" {
			delete(state.objects, target)
			state.launchers[target] = launcher
		} else if source := strings.ToLower(strings.TrimSpace(rhs)); state.launchers[source] != "" {
			delete(state.objects, target)
			state.launchers[target] = state.launchers[source]
		} else if !httpPartialNewMatches(state.objects[target].kind, rhs) {
			delete(state.objects, target)
		}
	}
	if name, expr, ok := fileAssignment(text); ok && !strings.HasPrefix(strings.ToLower(text), "set ") {
		key := strings.ToLower(name)
		if value, known := httpConstantString(expr, state); known {
			state.strings[key] = value
			state.known[key] = true
		} else {
			delete(state.strings, key)
			delete(state.known, key)
		}
		state.sensitive[key] = httpExpressionSensitive(expr, state)
	}

	if match := httpOptionAssignmentRe.FindStringSubmatch(text); len(match) == 5 {
		receiver := strings.ToLower(match[1])
		object := state.objects[receiver]
		option := strings.TrimSpace(match[2])
		if option == "" {
			option = strings.TrimSpace(match[3])
		}
		value := strings.TrimSpace(match[4])
		if object.kind == httpWinHTTP || object.kind == httpServerXML {
			if risk := httpTLSOptionRisk(object.kind, option, value, state); risk != "" && collect {
				findings = append(findings, httpFindingSpec{line: line, column: strings.Index(strings.ToLower(text), "option"), code: "VBA246", api: string(object.kind), risk: risk, redact: true})
			}
		}
	}
	if match := httpSetOptionCallRe.FindStringSubmatchIndex(text); match != nil {
		receiver := strings.ToLower(text[match[2]:match[3]])
		object := state.objects[receiver]
		args := httpMethodArgs(text, match[1])
		if (object.kind == httpServerXML || object.kind == httpWinHTTP) && len(args) >= 2 {
			if risk := httpTLSOptionRisk(object.kind, args[0], args[1], state); risk != "" && collect {
				findings = append(findings, httpFindingSpec{line: line, column: strings.Index(strings.ToLower(text), "setoption"), code: "VBA246", api: string(object.kind), risk: risk})
			}
		}
	}

	if httpLogRe.MatchString(text) && httpExpressionSensitive(text, state) && collect {
		findings = append(findings, httpFindingSpec{line: line, code: "VBA246", risk: "authorization_logging", redact: true})
	}

	if collect && isHTTPProcessLaunch(text, state) {
		for _, arg := range httpLaunchArgs(text) {
			path, known := httpConstantString(arg, state)
			if !known {
				continue
			}
			for _, candidate := range state.objects {
				if candidate.kind != httpADOStream {
					continue
				}
				launchKey := httpPathKey(path)
				for saved := range candidate.savedExecutable {
					if launchKey == saved || httpCommandContainsPath(launchKey, saved) {
						findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: "ADODB.Stream", risk: "download_and_execute", redact: true})
						break
					}
				}
			}
		}
	}
	match := httpMemberCallRe.FindStringSubmatchIndex(text)
	if match == nil {
		return state, findings
	}
	receiver := strings.ToLower(text[match[2]:match[3]])
	method := strings.ToLower(text[match[4]:match[5]])
	object, knownObject := state.objects[receiver]
	if !knownObject {
		return state, findings
	}
	args := httpMethodArgs(text, match[1])
	switch method {
	case "open":
		if object.kind != httpXML && object.kind != httpServerXML && object.kind != httpWinHTTP {
			break
		}
		if len(args) > 1 {
			object.url, object.urlKnown = httpConstantString(args[1], state)
			if object.urlKnown && httpURLHasCredentials(object.url) && httpIsWebURL(object.url) && collect {
				findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "credentials_in_url", origin: httpRedactedOrigin(object.url), redact: true})
			}
		}
		if len(args) > 3 && !httpKnownEmpty(args[3], state) {
			object.hasCredentials = true
			object.credentialSinks[statement.Range.StartByte] = true
		}
		if len(args) > 4 && !httpKnownEmpty(args[4], state) {
			object.hasCredentials = true
			object.credentialSinks[statement.Range.StartByte] = true
		}
		setHTTPObject(state, receiver, object)
	case "setrequestheader":
		if object.kind == httpXML || object.kind == httpServerXML || object.kind == httpWinHTTP {
			header, headerKnown := "", false
			if len(args) > 0 {
				header, headerKnown = httpConstantString(args[0], state)
			}
			if headerKnown && httpSensitiveHeaderRe.MatchString(strings.TrimSpace(header)) {
				object.hasCredentials = true
				object.credentialSinks[statement.Range.StartByte] = true
				sensitiveModuleConstant := len(args) > 1 && httpExpressionUsesSensitiveModuleConstant(args[1], file, state)
				if len(args) > 1 {
					markHTTPExpressionSensitive(args[1], state)
				}
				if collect && sensitiveModuleConstant {
					findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "sensitive_module_constant", header: canonicalHTTPHeader(header), redact: true})
				}
			}
			setHTTPObject(state, receiver, object)
		}
	case "setcredentials":
		if object.kind == httpWinHTTP {
			object.hasCredentials = true
			object.credentialSinks[statement.Range.StartByte] = true
			setHTTPObject(state, receiver, object)
		}
	case "settimeouts":
		if object.kind == httpWinHTTP || object.kind == httpServerXML {
			object.timeoutConfigured = len(args) == 4
			object.timeoutInfinite = false
			for _, arg := range args {
				if n, ok := httpInteger(arg, state); ok && (n == 0 || n == -1) {
					object.timeoutInfinite = true
					object.timeoutConfigured = false
				}
			}
			setHTTPObject(state, receiver, object)
		}
	case "send":
		if object.kind != httpXML && object.kind != httpServerXML && object.kind != httpWinHTTP {
			break
		}
		if collect && object.urlKnown && strings.EqualFold(httpURLScheme(object.url), "http") && object.hasCredentials && !a.isDevelopmentHTTPURL(object.url) {
			findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "plain_http_credentials", origin: httpRedactedOrigin(object.url), redact: true, ownedSinks: cloneHTTPIntSet(object.credentialSinks)})
		}
		if collect && (object.kind == httpServerXML || object.kind == httpWinHTTP) && !object.timeoutConfigured {
			timeout := "missing"
			if object.timeoutInfinite {
				timeout = "unbounded"
			}
			findings = append(findings, httpFindingSpec{line: line, code: "VBA247", api: string(object.kind), risk: "missing_http_timeout", timeout: timeout})
		}
	case "write":
		if object.kind == httpADOStream && len(args) > 0 {
			object.downloaded = httpResponseExpression(args[0], state)
			object.savedExecutable = map[string]bool{}
			setHTTPObject(state, receiver, object)
		}
	case "savetofile":
		if object.kind == httpADOStream && object.downloaded && len(args) > 0 {
			if path, known := httpConstantString(args[0], state); known && httpExecutableExtRe.MatchString(strings.ToLower(path)) {
				if object.savedExecutable == nil {
					object.savedExecutable = map[string]bool{}
				}
				object.savedExecutable[httpPathKey(path)] = true
				setHTTPObject(state, receiver, object)
			}
		}
	}
	return state, findings
}

func (a Analyzer) httpFinding(file parsedFile, proc sourceProcedure, spec httpFindingSpec) Finding {
	var message, reason, suggestion string
	if spec.code == "VBA247" {
		message = "HTTP request may wait indefinitely because a finite timeout is not configured."
		reason = "The recognized HTTP client reaches Send without a finite SetTimeouts configuration on every control-flow path."
		suggestion = "Call SetTimeouts with positive finite values before Open and Send."
	} else {
		message = "Unsafe HTTP transport or credential handling detected (" + spec.risk + ")."
		reason = "The recognized HTTP API has a statically proven transport-security or secret-exposure configuration."
		suggestion = httpSecuritySuggestion(spec.risk)
	}
	finding := a.simpleFinding(file, proc, spec.line, spec.code, "warning", message, reason, suggestion)
	finding.httpOwnedSinks = spec.ownedSinks
	finding.Column = max(0, spec.column)
	// HTTP findings always redact nearby strings/comments because neighboring
	// statements can contain credentials or sensitive URL path/query data even
	// when the finding itself is a TLS or timeout observation.
	finding.NearbyCode = redactedNearbyCode(file.Lines, spec.line, spec.line, 2)
	if spec.code == "VBA247" {
		finding.HTTPReliability = &HTTPReliabilityContext{API: spec.api, RiskKind: spec.risk, TimeoutState: spec.timeout}
	} else {
		finding.HTTPSecurity = &HTTPSecurityContext{API: spec.api, RiskKind: spec.risk, Header: spec.header, Origin: spec.origin}
	}
	return finding
}

func httpSecuritySuggestion(risk string) string {
	switch risk {
	case "plain_http_credentials":
		return "Use HTTPS for credential-bearing requests, or restrict an intentional loopback development origin in configuration."
	case "credentials_in_url":
		return "Remove credentials from the URL and use the client's credential or Authorization-header API over HTTPS."
	case "authorization_logging":
		return "Do not log authorization values; log only a request identifier and non-sensitive metadata."
	case "certificate_validation_bypass":
		return "Keep certificate validation and revocation checks enabled; install a trusted development certificate instead of bypassing validation."
	case "obsolete_tls_protocol":
		return "Allow TLS 1.2 or newer and remove SSL 2.0, SSL 3.0, TLS 1.0, and TLS 1.1 flags."
	case "download_and_execute":
		return "Do not immediately execute downloaded content; verify an expected signature/hash and enforce an explicit trusted publisher policy first."
	default:
		return "Load sensitive header values from a protected store and keep them out of module-level constants and diagnostics."
	}
}

func httpKindFromText(text string) httpClientKind {
	lower := strings.ToLower(strings.ReplaceAll(text, " ", ""))
	switch {
	case httpServerXMLTypeRe.MatchString(lower):
		return httpServerXML
	case httpWinHTTPTypeRe.MatchString(lower):
		return httpWinHTTP
	case httpXMLTypeRe.MatchString(lower):
		return httpXML
	case lower == "adodb.stream":
		return httpADOStream
	default:
		return httpUnknown
	}
}

func httpLauncherFromConstruction(text string) string {
	match := httpCreateObjectRe.FindStringSubmatch(text)
	if len(match) != 2 {
		return ""
	}
	switch strings.ToLower(strings.TrimSpace(match[1])) {
	case "wscript.shell":
		return "wscript.shell"
	case "shell.application":
		return "shell.application"
	default:
		return ""
	}
}

func httpKindFromConstruction(text string) httpClientKind {
	if match := httpNewObjectRe.FindStringSubmatch(text); len(match) == 2 {
		return httpKindFromText(match[1])
	}
	if match := httpCreateObjectRe.FindStringSubmatch(text); len(match) == 2 {
		return httpKindFromText(match[1])
	}
	return httpUnknown
}

func httpMethodArgs(text string, end int) []string {
	rest := strings.TrimSpace(text[end:])
	rest = strings.TrimPrefix(rest, "(")
	rest = strings.TrimSuffix(rest, ")")
	return splitArgs(rest)
}

func httpConstantString(expr string, state httpAnalysisState) (string, bool) {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return "", false
	}
	parts := splitHTTPConcat(expr)
	var b strings.Builder
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if len(part) >= 2 && part[0] == '"' && part[len(part)-1] == '"' {
			b.WriteString(strings.ReplaceAll(part[1:len(part)-1], `""`, `"`))
			continue
		}
		key := strings.ToLower(strings.TrimRight(part, "$%&@!#^"))
		if state.known[key] {
			b.WriteString(state.strings[key])
			continue
		}
		return "", false
	}
	return b.String(), true
}

func splitHTTPConcat(text string) []string {
	var out []string
	start := 0
	quoted := false
	depth := 0
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if quoted && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			quoted = !quoted
		case '(':
			if !quoted {
				depth++
			}
		case ')':
			if !quoted && depth > 0 {
				depth--
			}
		case '&':
			if !quoted && depth == 0 {
				out = append(out, text[start:i])
				start = i + 1
			}
		}
	}
	out = append(out, text[start:])
	return out
}

func httpInteger(expr string, state httpAnalysisState) (int, bool) {
	text := strings.TrimSpace(expr)
	if state.known[strings.ToLower(text)] {
		text = state.strings[strings.ToLower(text)]
	}
	text = strings.ReplaceAll(text, "&H", "0x")
	text = strings.ReplaceAll(text, "&h", "0x")
	text = strings.TrimRight(text, "%&@!#$")
	n, err := strconv.ParseInt(text, 0, strconv.IntSize)
	return int(n), err == nil
}

func httpTLSOptionRisk(kind httpClientKind, option, value string, state httpAnalysisState) string {
	lower := strings.ToLower(option)
	n, known := httpInteger(value, state)
	if kind == httpServerXML && (lower == "2" || strings.Contains(lower, "ignore_server_ssl_cert")) && known && n != 0 {
		return "certificate_validation_bypass"
	}
	if kind != httpWinHTTP {
		return ""
	}
	if (lower == "4" || strings.Contains(lower, "sslerrorignoreflags")) && known && n != 0 {
		return "certificate_validation_bypass"
	}
	if lower == "18" || strings.Contains(lower, "enablecertificaterevocationcheck") {
		if (known && n == 0) || strings.EqualFold(strings.TrimSpace(value), "false") {
			return "certificate_validation_bypass"
		}
	}
	if lower == "9" || strings.Contains(lower, "secureprotocols") {
		if known && n&(0x8|0x20|0x80|0x200) != 0 {
			return "obsolete_tls_protocol"
		}
	}
	return ""
}

func httpExpressionSensitive(expr string, state httpAnalysisState) bool {
	if httpAuthLiteralRe.MatchString(expr) {
		return true
	}
	for name, sensitive := range state.sensitive {
		if sensitive && containsHTTPIdentifierToken(strings.ToLower(expr), name) {
			return true
		}
	}
	return false
}

func markHTTPExpressionSensitive(expr string, state httpAnalysisState) {
	for _, match := range httpIdentifierRe.FindAllStringIndex(expr, -1) {
		name := expr[match[0]:match[1]]
		rest := strings.TrimLeft(expr[match[1]:], " \t")
		if strings.HasPrefix(rest, "(") {
			continue
		}
		state.sensitive[strings.ToLower(name)] = true
	}
}

func httpExpressionUsesSensitiveModuleConstant(expr string, file parsedFile, state httpAnalysisState) bool {
	if facts := file.ModuleFacts; facts != nil {
		found := false
		facts.forEachConstant(func(constant moduleConstantFact) {
			if found || !constant.Module {
				return
			}
			name := strings.ToLower(constant.Name)
			if state.sensitive[name] && containsHTTPIdentifierToken(strings.ToLower(expr), name) {
				found = true
			}
		})
		return found
	}
	for _, declaration := range file.IR.Declarations {
		name := strings.ToLower(declaration.Name)
		if declaration.Scope == procedureir.ScopeModule && declaration.Kind == "const" && state.sensitive[name] && containsHTTPIdentifierToken(strings.ToLower(expr), name) {
			return true
		}
	}
	return false
}

func httpResponseExpression(expr string, state httpAnalysisState) bool {
	lower := strings.ToLower(expr)
	if !strings.Contains(lower, "responsebody") && !strings.Contains(lower, "responsestream") {
		return false
	}
	for name, object := range state.objects {
		if (object.kind == httpXML || object.kind == httpServerXML || object.kind == httpWinHTTP) && containsHTTPIdentifierToken(lower, name) {
			return true
		}
	}
	return false
}

func setHTTPObject(state httpAnalysisState, receiver string, object httpObjectState) {
	state.objects[receiver] = cloneHTTPObject(object)
	if object.identity == "" {
		return
	}
	for name, alias := range state.objects {
		if name != receiver && alias.identity == object.identity {
			state.objects[name] = cloneHTTPObject(object)
		}
	}
}

func cloneHTTPObject(in httpObjectState) httpObjectState {
	out := in
	out.savedExecutable = map[string]bool{}
	for k, v := range in.savedExecutable {
		out.savedExecutable[k] = v
	}
	out.credentialSinks = cloneHTTPIntSet(in.credentialSinks)
	return out
}
func newHTTPObjectState(kind httpClientKind, identity string) httpObjectState {
	return httpObjectState{kind: kind, identity: identity, savedExecutable: map[string]bool{}, credentialSinks: map[int]bool{}}
}
func cloneHTTPIntSet(in map[int]bool) map[int]bool {
	out := map[int]bool{}
	for line, value := range in {
		out[line] = value
	}
	return out
}
func cloneHTTPState(in httpAnalysisState) httpAnalysisState {
	out := httpAnalysisState{objects: map[string]httpObjectState{}, launchers: map[string]string{}, strings: map[string]string{}, known: map[string]bool{}, sensitive: map[string]bool{}}
	for k, v := range in.objects {
		out.objects[k] = cloneHTTPObject(v)
	}
	for k, v := range in.strings {
		out.strings[k] = v
	}
	for k, v := range in.launchers {
		out.launchers[k] = v
	}
	for k, v := range in.known {
		out.known[k] = v
	}
	for k, v := range in.sensitive {
		out.sensitive[k] = v
	}
	return out
}

func joinHTTPState(left, right httpAnalysisState, initialized bool) (httpAnalysisState, bool) {
	if !initialized {
		return cloneHTTPState(right), true
	}
	out := cloneHTTPState(left)
	for name, r := range right.objects {
		l, ok := left.objects[name]
		if !ok {
			continue
		}
		joined := l
		if l.kind != r.kind {
			joined.kind = httpUnknown
		}
		if l.identity != r.identity {
			joined.identity = ""
		}
		joined.urlKnown = l.urlKnown && r.urlKnown && l.url == r.url
		if !joined.urlKnown {
			joined.url = ""
		}
		joined.hasCredentials = l.hasCredentials || r.hasCredentials
		joined.timeoutConfigured = l.timeoutConfigured && r.timeoutConfigured
		joined.timeoutInfinite = l.timeoutInfinite || r.timeoutInfinite
		joined.downloaded = l.downloaded && r.downloaded
		joined.savedExecutable = map[string]bool{}
		for p := range l.savedExecutable {
			if r.savedExecutable[p] {
				joined.savedExecutable[p] = true
			}
		}
		joined.credentialSinks = cloneHTTPIntSet(l.credentialSinks)
		for sink := range r.credentialSinks {
			joined.credentialSinks[sink] = true
		}
		out.objects[name] = joined
	}
	for name := range out.objects {
		if _, ok := right.objects[name]; !ok {
			delete(out.objects, name)
		}
	}
	for name, v := range right.sensitive {
		out.sensitive[name] = out.sensitive[name] || v
	}
	for name, launcher := range left.launchers {
		if right.launchers[name] != launcher {
			delete(out.launchers, name)
		}
	}
	for name, value := range left.strings {
		if !left.known[name] || !right.known[name] || right.strings[name] != value {
			delete(out.strings, name)
			delete(out.known, name)
		}
	}
	for name := range out.known {
		if _, ok := left.strings[name]; !ok {
			delete(out.strings, name)
			delete(out.known, name)
		}
	}
	return out, !sameHTTPState(left, out)
}

func httpPartialNewMatches(kind httpClientKind, rhs string) bool {
	lower := strings.ToLower(strings.TrimSpace(rhs))
	if !strings.HasPrefix(lower, "new ") {
		return false
	}
	typePrefix := strings.TrimSpace(strings.TrimPrefix(lower, "new "))
	return kind != httpUnknown && typePrefix != "" && strings.HasPrefix(strings.ToLower(string(kind)), typePrefix)
}

func sameHTTPState(a, b httpAnalysisState) bool {
	if len(a.objects) != len(b.objects) || len(a.launchers) != len(b.launchers) || len(a.strings) != len(b.strings) || len(a.known) != len(b.known) || len(a.sensitive) != len(b.sensitive) {
		return false
	}
	for name, left := range a.objects {
		right, ok := b.objects[name]
		if !ok || left.kind != right.kind || left.identity != right.identity || left.url != right.url || left.urlKnown != right.urlKnown || left.hasCredentials != right.hasCredentials || left.timeoutConfigured != right.timeoutConfigured || left.timeoutInfinite != right.timeoutInfinite || left.downloaded != right.downloaded || !sameHTTPBoolSet(left.savedExecutable, right.savedExecutable) || !sameHTTPIntSet(left.credentialSinks, right.credentialSinks) {
			return false
		}
	}
	for name, value := range a.launchers {
		if b.launchers[name] != value {
			return false
		}
	}
	for name, value := range a.strings {
		if b.strings[name] != value {
			return false
		}
	}
	for name, value := range a.known {
		if b.known[name] != value {
			return false
		}
	}
	for name, value := range a.sensitive {
		if b.sensitive[name] != value {
			return false
		}
	}
	return true
}

func sameHTTPBoolSet(a, b map[string]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key, value := range a {
		if b[key] != value {
			return false
		}
	}
	return true
}

func sameHTTPIntSet(a, b map[int]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key, value := range a {
		if b[key] != value {
			return false
		}
	}
	return true
}

func httpURLHasCredentials(raw string) bool {
	u, err := url.Parse(raw)
	return err == nil && u.User != nil && u.User.String() != ""
}
func httpIsWebURL(raw string) bool {
	scheme := strings.ToLower(httpURLScheme(raw))
	return scheme == "http" || scheme == "https"
}
func httpURLScheme(raw string) string {
	u, err := url.Parse(raw)
	if err != nil {
		return ""
	}
	return u.Scheme
}
func httpRedactedOrigin(raw string) string {
	u, err := url.Parse(raw)
	if err != nil || u.Host == "" {
		return ""
	}
	host := strings.ToLower(u.Hostname())
	port := u.Port()
	if (u.Scheme == "http" && port == "80") || (u.Scheme == "https" && port == "443") {
		port = ""
	}
	if strings.Contains(host, ":") {
		host = "[" + host + "]"
	}
	if port != "" {
		host = net.JoinHostPort(strings.Trim(host, "[]"), port)
	}
	return strings.ToLower(u.Scheme) + "://" + host
}
func (a Analyzer) isDevelopmentHTTPURL(raw string) bool {
	origin := httpRedactedOrigin(raw)
	u, err := url.Parse(origin)
	if err != nil {
		return false
	}
	host := u.Hostname()
	if strings.EqualFold(host, "localhost") {
		return true
	}
	if ip := net.ParseIP(host); ip != nil && (ip.IsLoopback()) {
		return true
	}
	for _, allowed := range a.Config.Analyze.DevelopmentHTTPOrigins {
		normalized, err := config.NormalizeDevelopmentHTTPOrigin(strings.TrimSpace(allowed))
		if err == nil && strings.EqualFold(normalized, origin) {
			return true
		}
	}
	return false
}
func canonicalHTTPHeader(value string) string {
	parts := strings.Split(strings.ToLower(strings.TrimSpace(value)), "-")
	for i := range parts {
		if parts[i] != "" {
			parts[i] = strings.ToUpper(parts[i][:1]) + parts[i][1:]
		}
	}
	return strings.Join(parts, "-")
}
func httpKnownEmpty(expr string, state httpAnalysisState) bool {
	value, known := httpConstantString(expr, state)
	return known && value == ""
}
func httpPathKey(path string) string { return strings.ToLower(filepath.Clean(strings.TrimSpace(path))) }
func httpCommandContainsPath(command, saved string) bool {
	for offset := 0; offset <= len(command)-len(saved); {
		index := strings.Index(command[offset:], saved)
		if index < 0 {
			return false
		}
		start := offset + index
		end := start + len(saved)
		leftOK := start == 0 || isHTTPCommandBoundary(command[start-1])
		rightOK := end == len(command) || isHTTPCommandBoundary(command[end])
		if leftOK && rightOK {
			return true
		}
		offset = start + 1
	}
	return false
}
func isHTTPCommandBoundary(value byte) bool {
	return value == ' ' || value == '\t' || value == '"' || value == '\'' || value == ',' || value == '(' || value == ')'
}
func isHTTPProcessLaunch(text string, state httpAnalysisState) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(lower, "shell ") || httpWin32LauncherRe.MatchString(lower) {
		return true
	}
	match := httpObjectLauncherRe.FindStringSubmatchIndex(text)
	if match == nil {
		return false
	}
	dot := strings.Index(text, ".")
	if dot < 0 {
		return false
	}
	receiver := strings.ToLower(strings.TrimSpace(text[:dot]))
	matchedCall := strings.ToLower(text[match[0]:match[1]])
	method := ""
	for _, candidate := range []string{"shellexecute", "exec", "run"} {
		if strings.Contains(matchedCall, candidate) {
			method = candidate
			break
		}
	}
	launcher := state.launchers[receiver]
	return launcher == "wscript.shell" && (method == "run" || method == "exec") || launcher == "shell.application" && method == "shellexecute"
}
func httpLaunchArgs(text string) []string {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(lower, "shell ") {
		return splitArgs(strings.TrimSpace(text[len("shell "):]))
	}
	if match := httpWin32LauncherName.FindStringIndex(strings.TrimSpace(text)); match != nil {
		trimmed := strings.TrimSpace(text)
		return splitArgs(strings.Trim(strings.TrimSpace(trimmed[match[1]:]), "()"))
	}
	if dot := strings.Index(text, "."); dot >= 0 {
		fields := strings.Fields(text[dot+1:])
		if len(fields) > 0 {
			rest := strings.TrimSpace(text[dot+1+len(fields[0]):])
			return splitArgs(strings.Trim(rest, "()"))
		}
	}
	return nil
}

func containsHTTPIdentifierToken(text, token string) bool {
	text, token = strings.ToLower(text), strings.ToLower(token)
	for start := 0; start <= len(text)-len(token); start++ {
		if !strings.HasPrefix(text[start:], token) {
			continue
		}
		end := start + len(token)
		leftOK := start == 0 || !isHTTPIdentifierByte(text[start-1])
		rightOK := end == len(text) || !isHTTPIdentifierByte(text[end])
		if leftOK && rightOK {
			return true
		}
	}
	return false
}

func isHTTPIdentifierByte(value byte) bool {
	return value >= 'a' && value <= 'z' || value >= '0' && value <= '9' || value == '_'
}
