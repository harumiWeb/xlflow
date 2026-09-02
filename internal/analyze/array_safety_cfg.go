package analyze

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// walkArrayCFG owns the common allocation-state worklist used by both the
// procedure findings pass and the array-return summary pass. Exceptional and
// uncertain edges retain the predecessor's input state because the statement
// may not have completed before control leaves the block.
func walkArrayCFG(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState) {
	walkArrayCFGWorklist(graph, lines, initial, visit, nil, nil, false)
}

func walkArrayCFGWithEdgesStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stats *arrayInterproceduralStats) {
	walkArrayCFGWithStopStats(graph, lines, initial, visit, edgeState, nil, stats)
}

func walkArrayCFGWithStopStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStats(graph, lines, initial, visit, edgeState, stop, false, stats)
}

func walkArrayCFGWithSourceLinesReliableStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliable(graph, lines, initial, visit, edgeState, nil, true, stats, reliableExceptional)
}

type arrayCFGBlockVisit func(block vbacfg.Block, text string, line int, in arrayFlowState) arrayFlowState

// arrayCFGWorklistReachable follows the same edge set as the array worklist.
// CFGView.IsReachable also expands unknown-flow sources for conservative
// diagnostics, but those synthetic reachability results do not cause the
// worklist to visit a disconnected nested statement block. The distinction is
// needed when source-line recovery has to attribute a call to its container.
func arrayCFGWorklistReachable(graph *vbacfg.CFGView) map[vbacfg.BlockID]bool {
	reachable := map[vbacfg.BlockID]bool{}
	if graph == nil {
		return reachable
	}
	entry := graph.Entry()
	reachable[entry] = true
	queue := []vbacfg.BlockID{entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if reachable[edge.To] {
				return true
			}
			reachable[edge.To] = true
			queue = append(queue, edge.To)
			return true
		})
	}
	return reachable
}

func arrayCFGBlockOwnsNestedStatements(block vbacfg.Block) bool {
	if block.Statement == nil {
		return false
	}
	switch block.Statement.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementElse,
		procedureir.StatementSelect, procedureir.StatementCase,
		procedureir.StatementFor, procedureir.StatementForEach,
		procedureir.StatementDo, procedureir.StatementWhile, procedureir.StatementWith:
		return true
	default:
		return false
	}
}

func walkArrayCFGWithSourceLinesReliableStatsAndBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliableAndBlock(graph, lines, initial, visit, visitBlock, edgeState, nil, true, stats, reliableExceptional)
}

func walkArrayCFGWorklist(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	walkArrayCFGWorklistStats(graph, lines, initial, visit, edgeState, stop, sourceLines, nil)
}

func walkArrayCFGWorklistStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliable(graph, lines, initial, visit, edgeState, stop, sourceLines, stats, nil)
}

func walkArrayCFGWorklistStatsWithReliable(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool) {
	walkArrayCFGWorklistStatsWithReliableAndBlock(graph, lines, initial, visit, nil, edgeState, stop, sourceLines, stats, reliableExceptional)
}

func walkArrayCFGWorklistStatsWithReliableAndBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool) {
	strategy := arrayCFGStrategyAuto
	if stats != nil {
		strategy = stats.strategy
	}
	if strategy == arrayCFGStrategyLegacy {
		if stats != nil {
			stats.addLegacyWalk()
		}
		walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, visitBlock, edgeState, stop, sourceLines)
		return
	}
	// The Array adapter owns the policy-specific source-line, edge, and stop
	// semantics at the indexed solver boundary. Only index/solver construction
	// incompatibility falls back to the legacy map worklist; a solve-time error
	// is not retried after analysis has begun.
	// A zero-slot state cannot carry the declaration-only scalar/object checks
	// that still consult the variables side table from the transfer callback.
	// Keep those paths on the legacy adapter; this is an intentional
	// compatibility boundary rather than an attempt to enlarge the compact
	// lattice with non-array declarations.
	if graph != nil && len(initial) > 0 {
		if _, err := semanticstate.NewIndexView(*graph); err == nil {
			if stats != nil {
				stats.addCompactWalk()
			}
			if stop == nil && !sourceLines && edgeState == nil && reliableExceptional == nil {
				_ = walkArrayCFGCompact(context.Background(), graph, lines, initial, visit, nil, false)
			} else {
				_ = walkArrayCFGCompactAdvancedWithBlock(context.Background(), graph, lines, initial, visit, visitBlock, edgeState, stop, reliableExceptional, sourceLines)
			}
			return
		}
	}
	if stats != nil {
		stats.addLegacyWalk()
		if graph != nil {
			// A forced compact run is a test/benchmark oracle. Unsupported graph
			// or participant layouts use this same compatibility fallback; do not
			// retry the compact solver after callbacks have started.
			reason := arrayFallbackIndex
			if len(initial) == 0 {
				reason = arrayFallbackEmptyState
			}
			stats.addFallbackReason(reason)
		}
	}
	walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, visitBlock, edgeState, stop, sourceLines)
}

func walkArrayCFGWorklistLegacy(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, nil, edgeState, stop, sourceLines)
}

func walkArrayCFGWorklistLegacyWithBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	if graph == nil {
		return
	}
	inStates := map[vbacfg.BlockID]arrayFlowState{graph.Entry(): initial}
	queued := map[vbacfg.BlockID]bool{graph.Entry(): true}
	for len(queued) > 0 {
		// Block IDs are ordered so the worklist cannot inherit Go map iteration
		// order and change the fixed-point path through the array state lattice.
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)
		in := cloneArrayState(inStates[id])
		block, ok := graph.BlockByID(id)
		if !ok {
			continue
		}
		// Keep the predecessor state intact for exceptional/uncertain edges.
		// Transfer functions mutate their input state, so give the current
		// block its own copy once instead of cloning again in every transfer.
		out := cloneArrayState(in)
		visitLine := func(text string, line int, state arrayFlowState) arrayFlowState {
			if visitBlock != nil {
				return visitBlock(block, text, line, state)
			}
			return visit(text, line, state)
		}
		if block.Statement != nil {
			if !sourceLines {
				line := block.Statement.Range.StartLine
				if line == 0 {
					line = block.Range.StartLine
				}
				text := block.Statement.Text
				if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
					text = normalizedCodeLine(lines[line-1])
				}
				// The transfer callback owns the block-local copy.  Keeping the
				// predecessor input untouched is required for exceptional and
				// uncertain edges, which deliberately propagate `in` below.
				out = visitLine(text, line, out)
				if stop != nil && stop(text, line) {
					continue
				}
			} else {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				end := block.Statement.Range.EndLine
				if end < start {
					end = start
				}
				stopped := false
				if (block.Statement.Kind == procedureir.StatementSelect || block.Statement.Kind == procedureir.StatementCase) && start >= 1 && start <= len(lines) {
					// Select Case and Case own separate CFG blocks for each branch.
					// Visiting a Case's whole source range here would scan nested
					// statements once before their branch edge facts are applied,
					// making a branch-local allocation fact appear to be absent.
					// The nested blocks below own the remaining physical lines.
					text := normalizedCodeLine(lines[start-1])
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				} else if start == end && start >= 1 && start <= len(lines) {
					// A single physical line can still have multiple logical CFG
					// statements, for example `If ... Then ReDim ...`. Preserve the
					// block's own text so the ReDim block is not mistaken for the
					// surrounding If statement.
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" {
						text = normalizedCodeLine(lines[start-1])
					}
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				} else if start >= 1 && end <= len(lines) {
					// CFG blocks may contain an entire multi-line loop or conditional
					// statement. Process its physical source lines in order so an
					// allocation earlier in the block is visible to later bounds or
					// element accesses. The CFG still owns path joins and exceptional
					// edges; this only restores statement order within a block.
					for line := start; line <= end; line++ {
						text := normalizedCodeLine(lines[line-1])
						if strings.TrimSpace(text) == "" {
							continue
						}
						out = visitLine(text, line, out)
						if stop != nil && stop(text, line) {
							stopped = true
							break
						}
					}
				} else {
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
						text = normalizedCodeLine(lines[start-1])
					}
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				}
				if stopped {
					continue
				}
			}
		}
		graph.ForEachOutgoing(id, func(edge vbacfg.Edge) bool {
			next := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				next = in
				// The source-line VBA227 pass still models exceptional edges
				// conservatively, except when the statement itself establishes a
				// deterministic allocation. A valid plain ReDim (or a recognized
				// built-in array factory assignment) cannot leave the target in its
				// pre-allocation state merely because the procedure has
				// `On Error Resume Next`; keeping that predecessor state would make
				// every following indexed write in the same branch look unsafe.
				if sourceLines && edge.Class == vbacfg.EdgeExceptional && arrayAllocationTransferIsReliable(block.Statement, in, out) {
					next = out
				}
			} else if edgeState != nil {
				next = edgeState(block, edge, next)
			}
			if mergeArrayState(inStates, edge.To, next) {
				queued[edge.To] = true
			}
			return true
		})
	}
}

// arrayCFGWorklistLane describes one array state policy that can be advanced
// by walkArrayCFGCombined.  Lanes deliberately own their state and transfer
// callbacks: sharing a queue must not force the block-level, source-line, and
// runtime policies to share a merge or edge interpretation.
//
// Graph is normally the same graph for every lane.  It is allowed to be a
// policy-specific copy, however (for example arrayVBA227Graph removes
// impossible normal continuations).  Block IDs are stable across those copies
// and the worklist remains shared even when an edge is absent from one lane.
type arrayCFGWorklistLane struct {
	Graph   *vbacfg.CFGView
	Initial arrayFlowState
	Stats   *arrayInterproceduralStats

	// Visit receives a private copy of the lane's block input state and returns
	// the state to propagate to outgoing edges.  In source-line mode it is
	// called once for each physical line owned by the block.
	Visit func(text string, line int, in arrayFlowState) arrayFlowState

	// EdgeState applies a lane-specific normal-edge refinement (guards,
	// Select Case facts, module configuration, and so on). It is not called for
	// exceptional or uncertain edges, which retain the predecessor state.
	EdgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState

	// Stop can terminate propagation for this lane after visiting a statement.
	// A stopped lane does not prevent other lanes from processing the same block.
	Stop func(text string, line int) bool

	// ReliableExceptional permits a source-line lane to carry a deterministic
	// allocation across an exceptional edge. The historical default is the
	// predecessor state; callers should set this only for the existing narrow
	// reliable-allocation contract.
	ReliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool

	// SourceLines restores physical source order inside a CFG block.  False
	// retains the historical single-statement block semantics.
	SourceLines bool
}

// walkArrayCFGCombined advances multiple array policies with one deterministic
// block/edge index and one queue. Each lane still has an independent state
// lattice and callback policy, preserving semantic compatibility while
// avoiding repeated graph scheduling and indexing work.
//
// Exceptional and uncertain edges retain each lane's predecessor input state.
// A lane may opt into the existing reliable-allocation exception through
// ReliableExceptional; this is intentionally independent from SourceLines so
// runtime or other block-level lanes cannot accidentally inherit it.
func walkArrayCFGCombined(ctx context.Context, lines []string, lanes []arrayCFGWorklistLane) error {
	if len(lanes) == 0 {
		return ctx.Err()
	}
	strategy := arrayCFGStrategyAuto
	for _, lane := range lanes {
		if lane.Stats == nil || lane.Stats.strategy == arrayCFGStrategyAuto {
			continue
		}
		strategy = lane.Stats.strategy
		break
	}
	if strategy != arrayCFGStrategyLegacy {
		if handled, err := walkArrayCFGCombinedCompact(ctx, lines, lanes); handled {
			return err
		}
	}
	// A forced compact run remains observable through the same compatibility
	// fallback when a lane cannot be indexed; no second compact attempt is
	// made after callbacks have started.
	fallbackReason := arrayFallbackUnsupported
	for _, lane := range lanes {
		if len(lane.Initial) == 0 {
			fallbackReason = arrayFallbackEmptyState
			break
		}
		if lane.Graph != nil {
			if _, err := semanticstate.NewIndexView(*lane.Graph); err != nil {
				fallbackReason = arrayFallbackIndex
				break
			}
		}
	}
	seenStats := map[*arrayInterproceduralStats]bool{}
	for _, lane := range lanes {
		if lane.Stats == nil || seenStats[lane.Stats] {
			continue
		}
		seenStats[lane.Stats] = true
		lane.Stats.addLegacyWalk()
		if strategy != arrayCFGStrategyLegacy {
			lane.Stats.addFallbackReason(fallbackReason)
		}
	}

	type graphIndex struct {
		graph *vbacfg.CFGView
	}
	type laneIndex struct {
		graphIndex
		inStates map[vbacfg.BlockID]arrayFlowState
	}
	indexes := make([]laneIndex, len(lanes))
	graphIndexes := make(map[*vbacfg.CFGView]graphIndex, len(lanes))
	queued := map[vbacfg.BlockID]bool{}
	for index, lane := range lanes {
		if lane.Graph == nil {
			continue
		}
		shared, ok := graphIndexes[lane.Graph]
		if !ok {
			shared = graphIndex{graph: lane.Graph}
			graphIndexes[lane.Graph] = shared
		}
		inStates := map[vbacfg.BlockID]arrayFlowState{
			lane.Graph.Entry(): cloneArrayState(lane.Initial),
		}
		indexes[index] = laneIndex{graphIndex: shared, inStates: inStates}
		queued[lane.Graph.Entry()] = true
	}

	for len(queued) > 0 {
		if err := ctx.Err(); err != nil {
			return err
		}
		// Block IDs are ordered so convergence does not depend on Go map
		// iteration order. A shared queue is safe even when a policy-specific
		// graph omits an edge: that lane simply has no outgoing edge to merge.
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)

		for index, lane := range lanes {
			stateIndex := &indexes[index]
			if lane.Graph == nil || lane.Visit == nil {
				continue
			}
			in, ok := stateIndex.inStates[id]
			if !ok {
				continue
			}
			block, ok := stateIndex.graph.BlockByID(id)
			if !ok {
				continue
			}

			in = cloneArrayState(in)
			out := cloneArrayState(in)
			stopped := false
			if block.Statement != nil {
				if !lane.SourceLines {
					line := block.Statement.Range.StartLine
					if line == 0 {
						line = block.Range.StartLine
					}
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
						text = normalizedCodeLine(lines[line-1])
					}
					// Keep the predecessor input immutable to the block transfer;
					// exceptional and uncertain edges must receive that input state.
					out = lane.Visit(text, line, out)
					stopped = lane.Stop != nil && lane.Stop(text, line)
				} else {
					start := block.Statement.Range.StartLine
					if start == 0 {
						start = block.Range.StartLine
					}
					end := block.Statement.Range.EndLine
					if end < start {
						end = start
					}
					if (block.Statement.Kind == procedureir.StatementSelect || block.Statement.Kind == procedureir.StatementCase) && start >= 1 && start <= len(lines) {
						// Select Case and Case own separate CFG blocks for each
						// branch; do not scan all clause lines before applying the
						// edge fact.
						text := normalizedCodeLine(lines[start-1])
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					} else if start == end && start >= 1 && start <= len(lines) {
						text := block.Statement.Text
						if strings.TrimSpace(text) == "" {
							text = normalizedCodeLine(lines[start-1])
						}
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					} else if start >= 1 && end <= len(lines) {
						for line := start; line <= end; line++ {
							if line&255 == 0 {
								if err := ctx.Err(); err != nil {
									return err
								}
							}
							text := normalizedCodeLine(lines[line-1])
							if strings.TrimSpace(text) == "" {
								continue
							}
							out = lane.Visit(text, line, out)
							if lane.Stop != nil && lane.Stop(text, line) {
								stopped = true
								break
							}
						}
					} else {
						text := block.Statement.Text
						if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
							text = normalizedCodeLine(lines[start-1])
						}
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					}
				}
			}
			if stopped {
				continue
			}

			stateIndex.graph.ForEachOutgoing(id, func(edge vbacfg.Edge) bool {
				next := out
				if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
					next = in
					if lane.SourceLines && edge.Class == vbacfg.EdgeExceptional && lane.ReliableExceptional != nil && lane.ReliableExceptional(block.Statement, in, out) {
						next = out
					}
				} else if lane.EdgeState != nil {
					next = lane.EdgeState(block, edge, next)
				}
				if mergeArrayState(stateIndex.inStates, edge.To, next) {
					queued[edge.To] = true
				}
				return true
			})
		}
	}
	return nil
}
