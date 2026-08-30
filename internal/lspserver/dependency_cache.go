package lspserver

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

// dependencyCache stores immutable project products by their content
// dependency key. Unlike revisionCache, adjacent workspace revisions may
// share entries when the inputs that product actually reads are unchanged.
// One build is shared per key and failed or canceled builds are retryable.
type dependencyCache[T any] struct {
	mu         sync.Mutex
	values     map[string]T
	order      []string
	pending    map[string]*dependencyCachePending
	maxEntries int
}

const (
	dependencyCacheMaxEntries     = 256
	projectEffectsCacheMaxEntries = 32
)

type dependencyCachePending struct {
	done chan struct{}
}

func (c *dependencyCache[T]) retentionLimit() int {
	if c.maxEntries > 0 {
		return c.maxEntries
	}
	return dependencyCacheMaxEntries
}

func (c *dependencyCache[T]) get(key string) (T, bool) {
	var zero T
	c.mu.Lock()
	value, ok := c.values[key]
	c.mu.Unlock()
	if !ok {
		return zero, false
	}
	return value, true
}

func (c *dependencyCache[T]) publish(key string, value T) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.values == nil {
		c.values = make(map[string]T)
	}
	if _, exists := c.values[key]; exists {
		return
	}
	c.values[key] = value
	c.order = append(c.order, key)
	for len(c.order) > c.retentionLimit() {
		oldest := c.order[0]
		c.order = c.order[1:]
		if oldest != key {
			delete(c.values, oldest)
		}
	}
}

// projectResolutionFingerprint covers the resolver's semantic symbol index,
// but deliberately excludes procedure bodies. Declaration locations remain in
// the key because resolver candidates carry source lines and effects matching
// uses those locations to identify a project procedure.
func projectResolutionFingerprint(project intel.ProjectAnalysisSnapshot, complete bool, typeLibSymbols []procedureir.ResolverSymbol) string {
	h := sha256.New()
	writeFingerprintText(h, "resolution-v2", boolFingerprint(complete))
	documents := append([]intel.ProjectAnalysisDocument(nil), project.Documents...)
	sort.Slice(documents, func(i, j int) bool {
		return symbolFileKey(documents[i].IR.Path) < symbolFileKey(documents[j].IR.Path)
	})
	for _, document := range documents {
		ir := document.IR
		writeFingerprintText(h, symbolFileKey(ir.Path), ir.ModuleName, ir.ModuleKind,
			hex.EncodeToString(document.ProcedureCatalog.ModuleContextHash[:]),
			hex.EncodeToString(document.ProcedureCatalog.ConditionalHash[:]))
		for _, declaration := range ir.Declarations {
			writeFingerprintText(h, declaration.Name, declaration.Type, declaration.Kind,
				declaration.Visibility, declaration.Parent, boolFingerprint(declaration.IsConst),
				boolFingerprint(declaration.IsArray), string(declaration.ValueShape),
				conditionalFingerprint(declaration.ConditionalBranches), boolFingerprint(declaration.Recovered),
				strconv.Itoa(declaration.Range.StartLine))
			for _, parameter := range declaration.Parameters {
				writeParameterFingerprint(h, parameter)
			}
		}
		for _, procedure := range ir.Procedures {
			symbol := procedure.Symbol
			writeFingerprintText(h, symbol.Name, symbol.QualifiedName, string(symbol.Kind),
				symbol.ReturnType, symbol.Visibility, boolFingerprint(symbol.IsArray),
				string(symbol.ValueShape), boolFingerprint(symbol.Recovered),
				boolFingerprint(symbol.IsEventHandler), symbol.EventKind,
				conditionalFingerprint(symbol.ConditionalBranches),
				strconv.Itoa(symbol.DeclarationRange.StartLine))
			for _, parameter := range symbol.Parameters {
				writeParameterFingerprint(h, parameter)
			}
			for _, bound := range symbol.ArrayBounds {
				writeFingerprintText(h, bound.Expression, bound.Lower, bound.Upper, boolFingerprint(bound.Recovered))
			}
		}
	}
	for _, symbol := range typeLibSymbols {
		writeFingerprintText(h, symbol.Name, symbol.Type, symbol.Module, symbol.ModuleKind,
			symbol.Kind, symbol.Visibility, symbol.File, symbol.Parent, boolFingerprint(symbol.IsConst),
			conditionalFingerprint(symbol.ConditionalBranches))
	}
	return "resolution:" + hex.EncodeToString(h.Sum(nil))
}

func writeParameterFingerprint(h interface{ Write([]byte) (int, error) }, parameter procedureir.Parameter) {
	writeFingerprintText(h, parameter.Name, parameter.Type, parameter.Passing,
		boolFingerprint(parameter.PassingExplicit), boolFingerprint(parameter.Optional),
		parameter.Default, boolFingerprint(parameter.HasDefault), boolFingerprint(parameter.ParamArray),
		boolFingerprint(parameter.IsArray), string(parameter.ArrayShape), string(parameter.ValueShape),
		boolFingerprint(parameter.Recovered))
	for _, bound := range parameter.ArrayBounds {
		writeFingerprintText(h, bound.Expression, bound.Lower, bound.Upper, boolFingerprint(bound.Recovered))
	}
}

func projectResolutionDocumentFingerprint(document intel.ProjectAnalysisDocument, resolverKey string, complete bool) string {
	h := sha256.New()
	version := projectDocumentContentFingerprint(document)
	writeFingerprintText(h, "resolution-document-v2", resolverKey, version,
		symbolFileKey(document.IR.Path), boolFingerprint(complete),
		hex.EncodeToString(document.ProcedureCatalog.ModuleContextHash[:]),
		hex.EncodeToString(document.ProcedureCatalog.ConditionalHash[:]))
	return "resolution-document:" + hex.EncodeToString(h.Sum(nil))
}

func projectDocumentContentFingerprint(document intel.ProjectAnalysisDocument) string {
	if strings.TrimSpace(document.Version) != "" {
		// Version is scoped to an editor lifecycle and may restart after a
		// close/reopen. Include the available content identity so a recycled
		// version cannot resurrect an overlay from the previous lifecycle.
		h := sha256.New()
		writeFingerprintText(h, "version", document.Version)
		if document.Source != "" {
			sourceHash := sourceVersion([]byte(document.Source))
			if document.Version == sourceHash {
				return document.Version
			}
			writeFingerprintText(h, sourceHash)
		} else {
			writeFingerprintText(h, projectProcedureIRFingerprint(document))
		}
		return "version:" + hex.EncodeToString(h.Sum(nil))
	}
	if document.Source != "" {
		return sourceVersion([]byte(document.Source))
	}
	// Compatibility snapshots may omit both the published source and version.
	// Hash the procedure-local semantic fields needed by resolution rather than
	// serializing the complete IR on the normal LSP path.
	return "ir:" + projectProcedureIRFingerprint(document)
}

func projectProcedureIRFingerprint(document intel.ProjectAnalysisDocument) string {
	// Compatibility snapshots may omit both Source and Version. Walk the full
	// IR explicitly so cloned pointer fields contribute their values rather than
	// allocation addresses. The document path is normalized before the walk so
	// equivalent Windows spellings share one fingerprint.
	h := sha256.New()
	ir := document.IR
	ir.Path = symbolFileKey(ir.Path)
	writeDeterministicFingerprintValue(h, reflect.ValueOf(ir))
	return hex.EncodeToString(h.Sum(nil))
}

func writeDeterministicFingerprintValue(w interface{ Write([]byte) (int, error) }, value reflect.Value) {
	if !value.IsValid() {
		writeFingerprintText(w, "invalid")
		return
	}
	switch value.Kind() {
	case reflect.Interface:
		if value.IsNil() {
			writeFingerprintText(w, "interface", "nil")
			return
		}
		writeFingerprintText(w, "interface", value.Type().String())
		writeDeterministicFingerprintValue(w, value.Elem())
	case reflect.Pointer:
		if value.IsNil() {
			writeFingerprintText(w, "pointer", value.Type().String(), "nil")
			return
		}
		writeFingerprintText(w, "pointer", value.Type().String())
		writeDeterministicFingerprintValue(w, value.Elem())
	case reflect.Struct:
		writeFingerprintText(w, "struct", value.Type().String())
		for index := 0; index < value.NumField(); index++ {
			field := value.Type().Field(index)
			writeFingerprintText(w, "field", field.Name)
			writeDeterministicFingerprintValue(w, value.Field(index))
		}
	case reflect.Slice, reflect.Array:
		// nil and empty slices carry the same semantic information in the IR;
		// procedureir.Clone may materialize one form from the other.
		writeFingerprintText(w, "sequence", value.Type().String(), strconv.Itoa(value.Len()))
		for index := 0; index < value.Len(); index++ {
			writeDeterministicFingerprintValue(w, value.Index(index))
		}
	case reflect.Map:
		if value.IsNil() {
			writeFingerprintText(w, "map", value.Type().String(), "nil")
			return
		}
		type mapEntry struct {
			sortKey string
			key     reflect.Value
			value   reflect.Value
		}
		entries := make([]mapEntry, 0, value.Len())
		for _, key := range value.MapKeys() {
			var keyText strings.Builder
			writeDeterministicFingerprintValue(&keyText, key)
			entries = append(entries, mapEntry{sortKey: keyText.String(), key: key, value: value.MapIndex(key)})
		}
		sort.Slice(entries, func(i, j int) bool { return entries[i].sortKey < entries[j].sortKey })
		writeFingerprintText(w, "map", value.Type().String(), strconv.Itoa(len(entries)))
		for _, entry := range entries {
			writeDeterministicFingerprintValue(w, entry.key)
			writeDeterministicFingerprintValue(w, entry.value)
		}
	case reflect.String:
		writeFingerprintText(w, "string", value.String())
	case reflect.Bool:
		writeFingerprintText(w, "bool", strconv.FormatBool(value.Bool()))
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		writeFingerprintText(w, "integer", value.Type().String(), strconv.FormatInt(value.Int(), 10))
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		writeFingerprintText(w, "unsigned", value.Type().String(), strconv.FormatUint(value.Uint(), 10))
	case reflect.Float32, reflect.Float64:
		writeFingerprintText(w, "float", value.Type().String(), strconv.FormatFloat(value.Float(), 'g', -1, value.Type().Bits()))
	case reflect.Complex64, reflect.Complex128:
		complexValue := value.Complex()
		writeFingerprintText(w, "complex", value.Type().String(),
			strconv.FormatFloat(real(complexValue), 'g', -1, value.Type().Bits()/2),
			strconv.FormatFloat(imag(complexValue), 'g', -1, value.Type().Bits()/2))
	default:
		// ProcedureIR currently contains no channels, functions, or unsafe
		// pointers. Keep the kind marker deterministic if that changes later.
		writeFingerprintText(w, "unsupported", value.Type().String(), value.Kind().String())
	}
}

func resolverSymbolsFingerprint(symbols []procedureir.ResolverSymbol, complete bool) string {
	ordered := append([]procedureir.ResolverSymbol(nil), symbols...)
	sort.Slice(ordered, func(i, j int) bool {
		left := ordered[i]
		right := ordered[j]
		leftKey := strings.Join([]string{left.Module, left.Name, left.Kind, left.File, left.Parent, strconv.Itoa(left.Line)}, "\x00")
		rightKey := strings.Join([]string{right.Module, right.Name, right.Kind, right.File, right.Parent, strconv.Itoa(right.Line)}, "\x00")
		return leftKey < rightKey
	})
	h := sha256.New()
	writeFingerprintText(h, "resolver-symbols-v2", boolFingerprint(complete))
	for _, symbol := range ordered {
		writeFingerprintText(h, symbol.Name, symbol.Type, symbol.Module, symbol.ModuleKind,
			symbol.Kind, symbol.Visibility, symbol.File, symbol.Parent, boolFingerprint(symbol.Recovered),
			boolFingerprint(symbol.IsArray), boolFingerprint(symbol.IsConst), string(symbol.ValueShape),
			conditionalFingerprint(symbol.ConditionalBranches), strconv.Itoa(symbol.Line))
	}
	return "resolver-symbols:" + hex.EncodeToString(h.Sum(nil))
}

func workspaceResolutionDocumentFingerprint(path, version, resolverKey string, complete bool) string {
	h := sha256.New()
	writeFingerprintText(h, "workspace-resolution-document-v2", symbolFileKey(path), version, resolverKey, boolFingerprint(complete))
	return "workspace-resolution-document:" + hex.EncodeToString(h.Sum(nil))
}

func boolFingerprint(value bool) string {
	if value {
		return "1"
	}
	return "0"
}

func conditionalFingerprint(branches []procedureir.ConditionalBranch) string {
	if len(branches) == 0 {
		return ""
	}
	var builder strings.Builder
	for _, branch := range branches {
		builder.WriteString(branch.Group)
		builder.WriteByte('\x00')
		builder.WriteString(branch.Condition)
		builder.WriteByte('\x00')
		builder.WriteString(strconv.Itoa(branch.Branch))
		builder.WriteByte('\x00')
	}
	return builder.String()
}

func projectEffectsDependencyFingerprint(versions map[string]string, complete bool, resolutionKey string) string {
	h := sha256.New()
	writeFingerprintText(h, "effects-v2", boolFingerprint(complete), resolutionKey)
	paths := make([]string, 0, len(versions))
	for path := range versions {
		paths = append(paths, path)
	}
	sort.Strings(paths)
	for _, path := range paths {
		writeFingerprintText(h, path, versions[path])
	}
	return "effects:" + hex.EncodeToString(h.Sum(nil))
}

func projectCapabilityDependencyFingerprint(project intel.ProjectAnalysisSnapshot, complete bool, resolutionKey string, cfg config.AnalyzeConfig) string {
	h := sha256.New()
	writeFingerprintText(h, "capabilities-v2", boolFingerprint(complete), resolutionKey)
	if encoded, err := json.Marshal(cfg); err == nil {
		writeFingerprintText(h, string(encoded))
	}
	documents := append([]intel.ProjectAnalysisDocument(nil), project.Documents...)
	sort.Slice(documents, func(i, j int) bool { return symbolFileKey(documents[i].IR.Path) < symbolFileKey(documents[j].IR.Path) })
	for _, document := range documents {
		writeFingerprintText(h, symbolFileKey(document.IR.Path), projectDocumentContentFingerprint(document),
			document.IR.ModuleName, document.IR.ModuleKind,
			hex.EncodeToString(document.ProcedureCatalog.ModuleContextHash[:]),
			hex.EncodeToString(document.ProcedureCatalog.ConditionalHash[:]))
	}
	return "capabilities:" + hex.EncodeToString(h.Sum(nil))
}

func countProjectCapabilities(requirements analyze.ProjectCapabilityRequirements) int {
	count := 0
	for _, enabled := range []bool{
		requirements.TypeDB, requirements.Resolution, requirements.ProjectConstants,
		requirements.ByRefSymbols, requirements.Effects, requirements.ObjectFlow,
		requirements.ArrayInterprocedural, requirements.DataFlowInputs,
		requirements.DictionaryCollection, requirements.ApplicationState,
		requirements.EventReentry, requirements.PublicAPITypeIndex,
		requirements.ExcelLoopSymbols, requirements.ExcelAPIHelpers, requirements.ModuleState,
	} {
		if enabled {
			count++
		}
	}
	return count
}

func cloneProjectEffectVersions(in map[string]string) map[string]string {
	if len(in) == 0 {
		return nil
	}
	out := make(map[string]string, len(in))
	for key, value := range in {
		out[key] = value
	}
	return out
}

func changedProjectEffectFiles(previous, current map[string]string, previousValid bool) map[string]struct{} {
	changed := make(map[string]struct{})
	if !previousValid {
		for path := range current {
			changed[path] = struct{}{}
		}
		return changed
	}
	for path, version := range previous {
		if current[path] != version {
			changed[path] = struct{}{}
		}
	}
	for path := range current {
		if _, ok := previous[path]; !ok {
			changed[path] = struct{}{}
		}
	}
	return changed
}

func projectConstantsDependencyFingerprint(project intel.ProjectAnalysisSnapshot, complete bool, typeDB *vbadb.DB) string {
	h := sha256.New()
	writeFingerprintText(h, "constants-v2", boolFingerprint(complete))
	documents := append([]intel.ProjectAnalysisDocument(nil), project.Documents...)
	sort.Slice(documents, func(i, j int) bool { return symbolFileKey(documents[i].IR.Path) < symbolFileKey(documents[j].IR.Path) })
	for _, document := range documents {
		writeFingerprintText(h, symbolFileKey(document.IR.Path), document.IR.ModuleName, document.IR.ModuleKind,
			hex.EncodeToString(document.ProcedureCatalog.ModuleContextHash[:]),
			hex.EncodeToString(document.ProcedureCatalog.ConditionalHash[:]))
		for _, declaration := range document.IR.Declarations {
			writeFingerprintText(h, declaration.Name, declaration.Type, declaration.Kind,
				declaration.Visibility, declaration.Parent, boolFingerprint(declaration.IsConst),
				boolFingerprint(declaration.IsArray), string(declaration.ValueShape),
				conditionalFingerprint(declaration.ConditionalBranches), boolFingerprint(declaration.Recovered),
				strconv.Itoa(declaration.Range.StartLine))
		}
		// Constant expressions live in the module preamble. A procedure body
		// edit therefore does not alter this input, while a Const/Enum edit does.
		preamble := document.Source
		if len(document.ProcedureCatalog.Entries) > 0 {
			start := document.ProcedureCatalog.Entries[0].StartByte
			if start >= 0 && start <= len(preamble) {
				preamble = preamble[:start]
			}
		}
		writeFingerprintText(h, preamble)
	}
	if typeDB != nil {
		for _, constant := range typeDB.AllConstantsList() {
			writeFingerprintText(h, constant.Name, constant.Type, constant.Value, constant.EnumGroup, constant.Library)
		}
	}
	return "constants:" + hex.EncodeToString(h.Sum(nil))
}

func (c *dependencyCache[T]) getOrBuildContext(ctx context.Context, key string, build func() (T, error)) (T, error, bool) {
	var zero T
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return zero, err, false
	}

	for {
		c.mu.Lock()
		if c.values == nil {
			c.values = make(map[string]T)
		}
		if value, ok := c.values[key]; ok {
			c.mu.Unlock()
			return value, nil, true
		}
		if c.pending == nil {
			c.pending = make(map[string]*dependencyCachePending)
		}
		if pending := c.pending[key]; pending != nil {
			done := pending.done
			c.mu.Unlock()
			select {
			case <-done:
				if err := ctx.Err(); err != nil {
					return zero, err, false
				}
				continue
			case <-ctx.Done():
				return zero, ctx.Err(), false
			}
		}
		pending := &dependencyCachePending{done: make(chan struct{})}
		c.pending[key] = pending
		c.mu.Unlock()

		var value T
		var err error
		var panicValue any
		func() {
			defer func() { panicValue = recover() }()
			value, err = build()
		}()
		c.mu.Lock()
		delete(c.pending, key)
		close(pending.done)
		if panicValue == nil && err == nil && ctx.Err() == nil {
			c.values[key] = value
			c.order = append(c.order, key)
			for len(c.order) > c.retentionLimit() {
				oldest := c.order[0]
				c.order = c.order[1:]
				if oldest != key {
					delete(c.values, oldest)
				}
			}
		}
		c.mu.Unlock()
		if panicValue != nil {
			panic(panicValue)
		}
		if err != nil {
			return zero, err, false
		}
		if err := ctx.Err(); err != nil {
			return zero, err, false
		}
		return value, nil, false
	}
}
