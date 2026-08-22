package analysisstats

import (
	"context"
	"sort"
	"sync"
	"time"
)

type Stage struct {
	Name        string
	Elapsed     time.Duration
	Wait        time.Duration
	Outcome     string
	ResultCount int
	Calls       int
}

// Domain identifies one of the semantic areas of procedure-local analysis.
//
// Domain values are deliberately a small integer enum. A DomainAggregate uses
// the value as an index into a fixed-size array, so recording a domain does not
// require a map lookup or an allocation. Unknown values are attributed to
// DomainOther.
type Domain uint8

const (
	DomainSourceScan Domain = iota
	DomainRuntime
	DomainArray
	DomainObject
	DomainDictionary
	DomainError
	DomainDataflow
	DomainResource
	DomainExcel
	DomainApplicationState
	DomainOther
)

const domainCount = int(DomainOther) + 1

const (
	ProcedureLocalSourceScan       = "procedure_local/source_scan"
	ProcedureLocalRuntime          = "procedure_local/runtime"
	ProcedureLocalArray            = "procedure_local/array"
	ProcedureLocalObject           = "procedure_local/object"
	ProcedureLocalDictionary       = "procedure_local/dictionary"
	ProcedureLocalError            = "procedure_local/error"
	ProcedureLocalDataflow         = "procedure_local/dataflow"
	ProcedureLocalResource         = "procedure_local/resource"
	ProcedureLocalExcel            = "procedure_local/excel"
	ProcedureLocalApplicationState = "procedure_local/application_state"
	ProcedureLocalOther            = "procedure_local/other"
)

var domainNames = [...]string{
	ProcedureLocalSourceScan,
	ProcedureLocalRuntime,
	ProcedureLocalArray,
	ProcedureLocalObject,
	ProcedureLocalDictionary,
	ProcedureLocalError,
	ProcedureLocalDataflow,
	ProcedureLocalResource,
	ProcedureLocalExcel,
	ProcedureLocalApplicationState,
	ProcedureLocalOther,
}

// These paired array declarations fail compilation if a domain is added
// without adding its canonical output name (or vice versa).
var (
	_ [len(domainNames) - domainCount]struct{}
	_ [domainCount - len(domainNames)]struct{}
)

// String returns the stable stage name used by performance-log output.
func (d Domain) String() string {
	if int(d) >= 0 && int(d) < len(domainNames) {
		return domainNames[d]
	}
	return ProcedureLocalOther
}

// WorkCounter identifies a fixed workload counter. Counter values are kept in
// a fixed-size array by DomainAggregate; the string names are emitted only
// when an aggregate is merged into a Recorder.
type WorkCounter uint8

const (
	CounterRuntimeCandidateProcedures WorkCounter = iota
	CounterArrayCandidateProcedures
	CounterObjectCandidateProcedures
	CounterDictionaryCandidateProcedures
	CounterErrorCandidateProcedures
	CounterDataflowCandidateProcedures
	CounterResourceCandidateProcedures
	CounterExcelCandidateProcedures
	CounterApplicationStateCandidateProcedures
	CounterSourceLineScans
	CounterRuntimeCFGWalks
	CounterArrayCFGWalks
	CounterDictionaryCFGWalks
	CounterErrorCFGWalks
	CounterDataflowCFGWalks
	CounterResourceCFGWalks
	CounterExcelCFGWalks
	CounterSemanticKernelRuns
	CounterRuntimePlannedRuns
	CounterRuntimeSkippedRuns
	CounterArrayPlannedRuns
	CounterArraySkippedRuns
	CounterObjectPlannedRuns
	CounterObjectSkippedRuns
	CounterDictionaryPlannedRuns
	CounterDictionarySkippedRuns
	CounterErrorPlannedRuns
	CounterErrorSkippedRuns
	CounterDataflowPlannedRuns
	CounterDataflowSkippedRuns
	CounterResourcePlannedRuns
	CounterResourceSkippedRuns
	CounterExcelPlannedRuns
	CounterExcelSkippedRuns
	CounterApplicationStatePlannedRuns
	CounterApplicationStateSkippedRuns
)

const counterCount = int(CounterApplicationStateSkippedRuns) + 1

// WorkCounterCount is the fixed number of counter slots used by
// DomainAggregate. It is exposed for compile-time guards in lightweight
// callers that use a bitset for per-procedure de-duplication.
const WorkCounterCount = counterCount

const (
	RuntimeCandidateProceduresCounter          = "runtime_candidate_procedures"
	ArrayCandidateProceduresCounter            = "array_candidate_procedures"
	ObjectCandidateProceduresCounter           = "object_candidate_procedures"
	DictionaryCandidateProceduresCounter       = "dictionary_candidate_procedures"
	ErrorCandidateProceduresCounter            = "error_candidate_procedures"
	DataflowCandidateProceduresCounter         = "dataflow_candidate_procedures"
	ResourceCandidateProceduresCounter         = "resource_candidate_procedures"
	ExcelCandidateProceduresCounter            = "excel_candidate_procedures"
	ApplicationStateCandidateProceduresCounter = "application_state_candidate_procedures"
	SourceLineScansCounter                     = "source_line_scans"
	RuntimeCFGWalksCounter                     = "runtime_cfg_walks"
	ArrayCFGWalksCounter                       = "array_cfg_walks"
	DictionaryCFGWalksCounter                  = "dictionary_cfg_walks"
	ErrorCFGWalksCounter                       = "error_cfg_walks"
	DataflowCFGWalksCounter                    = "dataflow_cfg_walks"
	ResourceCFGWalksCounter                    = "resource_cfg_walks"
	ExcelCFGWalksCounter                       = "excel_cfg_walks"
	SemanticKernelRunsCounter                  = "semantic_kernel_runs"
	RuntimePlannedRunsCounter                  = "planned_runtime_runs"
	RuntimeSkippedRunsCounter                  = "skipped_runtime_runs"
	ArrayPlannedRunsCounter                    = "planned_array_runs"
	ArraySkippedRunsCounter                    = "skipped_array_runs"
	ObjectPlannedRunsCounter                   = "planned_object_runs"
	ObjectSkippedRunsCounter                   = "skipped_object_runs"
	DictionaryPlannedRunsCounter               = "planned_dictionary_runs"
	DictionarySkippedRunsCounter               = "skipped_dictionary_runs"
	ErrorPlannedRunsCounter                    = "planned_error_runs"
	ErrorSkippedRunsCounter                    = "skipped_error_runs"
	DataflowPlannedRunsCounter                 = "planned_dataflow_runs"
	DataflowSkippedRunsCounter                 = "skipped_dataflow_runs"
	ResourcePlannedRunsCounter                 = "planned_resource_runs"
	ResourceSkippedRunsCounter                 = "skipped_resource_runs"
	ExcelPlannedRunsCounter                    = "planned_excel_runs"
	ExcelSkippedRunsCounter                    = "skipped_excel_runs"
	ApplicationStatePlannedRunsCounter         = "planned_application_state_runs"
	ApplicationStateSkippedRunsCounter         = "skipped_application_state_runs"
)

var counterNames = [...]string{
	RuntimeCandidateProceduresCounter,
	ArrayCandidateProceduresCounter,
	ObjectCandidateProceduresCounter,
	DictionaryCandidateProceduresCounter,
	ErrorCandidateProceduresCounter,
	DataflowCandidateProceduresCounter,
	ResourceCandidateProceduresCounter,
	ExcelCandidateProceduresCounter,
	ApplicationStateCandidateProceduresCounter,
	SourceLineScansCounter,
	RuntimeCFGWalksCounter,
	ArrayCFGWalksCounter,
	DictionaryCFGWalksCounter,
	ErrorCFGWalksCounter,
	DataflowCFGWalksCounter,
	ResourceCFGWalksCounter,
	ExcelCFGWalksCounter,
	SemanticKernelRunsCounter,
	RuntimePlannedRunsCounter,
	RuntimeSkippedRunsCounter,
	ArrayPlannedRunsCounter,
	ArraySkippedRunsCounter,
	ObjectPlannedRunsCounter,
	ObjectSkippedRunsCounter,
	DictionaryPlannedRunsCounter,
	DictionarySkippedRunsCounter,
	ErrorPlannedRunsCounter,
	ErrorSkippedRunsCounter,
	DataflowPlannedRunsCounter,
	DataflowSkippedRunsCounter,
	ResourcePlannedRunsCounter,
	ResourceSkippedRunsCounter,
	ExcelPlannedRunsCounter,
	ExcelSkippedRunsCounter,
	ApplicationStatePlannedRunsCounter,
	ApplicationStateSkippedRunsCounter,
}

// These paired array declarations fail compilation if a counter is added
// without adding its canonical output name (or vice versa).
var (
	_ [len(counterNames) - counterCount]struct{}
	_ [counterCount - len(counterNames)]struct{}
)

// String returns the stable counter name used by performance-log output.
func (c WorkCounter) String() string {
	if int(c) >= 0 && int(c) < len(counterNames) {
		return counterNames[c]
	}
	return ""
}

// Fact-build counters are additive observations. They deliberately live in
// analysisstats rather than in an analyzer rule package so batch, realtime,
// and benchmark callers use the same stable names when reporting the amount
// of shared-facts preparation work.
const (
	ModuleFactBuildsCounter    = "module_fact_builds"
	ProcedureFactBuildsCounter = "procedure_fact_builds"
)

type Recorder struct {
	mu       sync.Mutex
	stages   []Stage
	counters map[string]uint64
}

func NewRecorder() *Recorder { return &Recorder{counters: make(map[string]uint64)} }

type aggregateStage struct {
	elapsed     time.Duration
	resultCount int
	calls       int
	outcome     string
}

// DomainAggregate collects procedure-local measurements before publishing
// them to a Recorder. It is intended to be owned by one analyzer worker (or
// one file analysis) and merged once when that work completes. Its fixed-size
// arrays avoid per-procedure maps, slices, recorder calls, and lock
// contention. Multiple aggregates may be merged concurrently into a Recorder.
type DomainAggregate struct {
	recorder *Recorder
	stages   [domainCount]aggregateStage
	counters [counterCount]uint64
}

// NewAggregate returns a local aggregate connected to recorder. When recorder
// is nil it returns nil, allowing callers to keep the disabled profiling path
// allocation-free. A nil *Recorder is also safe when called as a method.
func NewAggregate(recorder *Recorder) *DomainAggregate {
	if recorder == nil {
		return nil
	}
	return &DomainAggregate{recorder: recorder}
}

// Start begins timing one aggregate domain. The zero Measurement returned for
// a nil aggregate is a no-op, so the normal non-profiled path does not call
// time.Now.
func (a *DomainAggregate) Start(domain Domain) Measurement {
	if a == nil {
		return Measurement{}
	}
	return Measurement{aggregate: a, domain: normalizeDomain(domain), started: time.Now()}
}

// Record adds an already measured domain observation to this local aggregate.
// No operation is performed for a nil aggregate.
func (a *DomainAggregate) Record(domain Domain, elapsed time.Duration, resultCount int, outcome string) {
	if a == nil {
		return
	}
	stage := &a.stages[normalizeDomain(domain)]
	stage.elapsed += elapsed
	stage.resultCount += resultCount
	stage.calls++
	stage.outcome = combinedOutcome(stage.outcome, outcome)
}

// AddCounter adds an observation to one of the fixed workload counters. An
// invalid counter is ignored, which keeps callers safe if they pass a value
// from a future or external enum without introducing a dynamic map.
func (a *DomainAggregate) AddCounter(counter WorkCounter, value uint64) {
	if a == nil || int(counter) < 0 || int(counter) >= len(a.counters) {
		return
	}
	a.counters[counter] += value
}

// Merge publishes this local aggregate to its Recorder with one recorder lock
// acquisition. The local values are retained, so callers may inspect or merge
// the same aggregate again when that is useful for a test or a staged caller.
func (a *DomainAggregate) Merge() {
	if a == nil || a.recorder == nil {
		return
	}
	a.recorder.MergeAggregate(a)
}

// Measurement is a stack-friendly timer returned by DomainAggregate.Start.
// It intentionally does not use a closure, which avoids one allocation for
// every measured procedure/domain.
type Measurement struct {
	aggregate *DomainAggregate
	domain    Domain
	started   time.Time
}

// Finish records a measured domain and derives the outcome from err. A nil
// error is "ok"; a non-nil error is "error". Context cancellation can be
// represented explicitly with FinishOutcome when the caller has that state.
func (m Measurement) Finish(resultCount int, err error) {
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	m.FinishOutcome(resultCount, outcome)
}

// FinishOutcome records a measured domain with an explicit outcome.
func (m Measurement) FinishOutcome(resultCount int, outcome string) {
	if m.aggregate == nil {
		return
	}
	m.aggregate.Record(m.domain, time.Since(m.started), resultCount, outcome)
}

// MergeAggregate publishes all non-empty domain stages and counters in their
// canonical order. The single lock acquisition is the important property for
// large modules: a worker can record many procedures locally without touching
// the shared recorder until this method is called.
func (r *Recorder) MergeAggregate(a *DomainAggregate) {
	if r == nil || a == nil {
		return
	}
	r.mu.Lock()
	if r.counters == nil {
		r.counters = make(map[string]uint64)
	}
	for index, stage := range a.stages {
		if stage.calls == 0 {
			continue
		}
		r.stages = append(r.stages, Stage{
			Name:        domainNames[index],
			Elapsed:     stage.elapsed,
			Outcome:     stage.outcome,
			ResultCount: stage.resultCount,
			Calls:       stage.calls,
		})
	}
	for index, value := range a.counters {
		if value != 0 {
			r.counters[counterNames[index]] += value
		}
	}
	r.mu.Unlock()
}

func normalizeDomain(domain Domain) Domain {
	if int(domain) >= 0 && int(domain) < domainCount {
		return domain
	}
	return DomainOther
}

func (r *Recorder) Record(stage Stage) {
	if r == nil {
		return
	}
	r.mu.Lock()
	if stage.Calls == 0 {
		stage.Calls = 1
	}
	r.stages = append(r.stages, stage)
	r.mu.Unlock()
}

// Totals returns stages aggregated by name and counters sorted by name. Stage
// order follows the first observation so callers can render the analysis
// pipeline in execution order while repeated per-file measurements collapse
// into one stable record.
func (r *Recorder) Totals() ([]Stage, []Counter) {
	stages, counters := r.Snapshot()
	if len(stages) == 0 && len(counters) == 0 {
		return nil, nil
	}
	indexes := make(map[string]int, len(stages))
	totals := make([]Stage, 0, len(stages))
	for _, stage := range stages {
		index, ok := indexes[stage.Name]
		if !ok {
			indexes[stage.Name] = len(totals)
			if stage.Calls == 0 {
				stage.Calls = 1
			}
			totals = append(totals, stage)
			continue
		}
		total := &totals[index]
		total.Elapsed += stage.Elapsed
		total.Wait += stage.Wait
		total.ResultCount += stage.ResultCount
		if stage.Calls == 0 {
			stage.Calls = 1
		}
		total.Calls += stage.Calls
		total.Outcome = combinedOutcome(total.Outcome, stage.Outcome)
	}
	names := make([]string, 0, len(counters))
	for name := range counters {
		names = append(names, name)
	}
	sort.Strings(names)
	orderedCounters := make([]Counter, 0, len(names))
	for _, name := range names {
		orderedCounters = append(orderedCounters, Counter{Name: name, Value: counters[name]})
	}
	return totals, orderedCounters
}

type Counter struct {
	Name  string
	Value uint64
}

func combinedOutcome(current, next string) string {
	rank := func(outcome string) int {
		switch outcome {
		case "canceled":
			return 3
		case "error":
			return 2
		case "ok":
			return 1
		default:
			return 0
		}
	}
	if rank(next) > rank(current) {
		return next
	}
	return current
}

func (r *Recorder) Add(name string, value uint64) {
	r.AddSum(name, value)
}

// AddSum adds value to an additive counter. Add is retained as a compatibility
// alias for existing instrumentation callers.
func (r *Recorder) AddSum(name string, value uint64) {
	if r == nil {
		return
	}
	r.mu.Lock()
	r.counters[name] += value
	r.mu.Unlock()
}

// AddMax records the greatest value observed for a maximum counter. Maximum
// workload dimensions are intentionally separate from additive counters so
// repeated observations do not turn a per-file or per-procedure maximum into
// a sum.
func (r *Recorder) AddMax(name string, value uint64) {
	if r == nil {
		return
	}
	r.mu.Lock()
	current, ok := r.counters[name]
	if !ok || value > current {
		r.counters[name] = value
	}
	r.mu.Unlock()
}

// RecordModuleFactBuild records one construction of the immutable file-level
// facts for an analysis revision. A nil recorder is intentionally a no-op.
func (r *Recorder) RecordModuleFactBuild() {
	r.AddSum(ModuleFactBuildsCounter, 1)
}

// RecordProcedureFactBuild records one construction of the immutable
// procedure-level facts for an analysis revision. A nil recorder is
// intentionally a no-op.
func (r *Recorder) RecordProcedureFactBuild() {
	r.RecordProcedureFactBuilds(1)
}

// RecordProcedureFactBuilds records multiple procedure-fact constructions in
// one counter update. Callers that already know the procedure count should use
// this form so instrumentation does not take one recorder lock per procedure
// in a large-module benchmark.
func (r *Recorder) RecordProcedureFactBuilds(count uint64) {
	r.AddSum(ProcedureFactBuildsCounter, count)
}

func (r *Recorder) Snapshot() ([]Stage, map[string]uint64) {
	if r == nil {
		return nil, nil
	}
	r.mu.Lock()
	defer r.mu.Unlock()
	stages := append([]Stage(nil), r.stages...)
	counters := make(map[string]uint64, len(r.counters))
	for key, value := range r.counters {
		counters[key] = value
	}
	return stages, counters
}

type recorderKey struct{}

func WithRecorder(ctx context.Context, recorder *Recorder) context.Context {
	if recorder == nil {
		return ctx
	}
	return context.WithValue(ctx, recorderKey{}, recorder)
}

func FromContext(ctx context.Context) *Recorder {
	if ctx == nil {
		return nil
	}
	recorder, _ := ctx.Value(recorderKey{}).(*Recorder)
	return recorder
}

func Measure(ctx context.Context, name string) func(int, error) {
	recorder := FromContext(ctx)
	if recorder == nil {
		return func(int, error) {}
	}
	started := time.Now()
	return func(count int, err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		recorder.Record(Stage{Name: name, Elapsed: time.Since(started), Outcome: outcome, ResultCount: count})
	}
}

// MeasureWait records time spent waiting for a shared artifact owned by
// another analysis request. It is separate from compute elapsed time so stage
// logs can distinguish contention from actual analysis work.
func MeasureWait(ctx context.Context, name string) func(error) {
	recorder := FromContext(ctx)
	if recorder == nil {
		return func(error) {}
	}
	started := time.Now()
	return func(err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		recorder.Record(Stage{Name: name, Wait: time.Since(started), Outcome: outcome})
	}
}
