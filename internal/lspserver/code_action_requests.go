package lspserver

import (
	"context"
	"encoding/json"
	"errors"
	"sync"

	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

const maxPendingCodeActions = 16

// Only code actions leave the ordered transport handler. One worker and a
// bounded queue prevent automatic editor requests from multiplying analysis.
type codeActionRequests struct {
	mu      sync.Mutex
	pending []*codeActionRequest
	active  *codeActionRequest
	stopped bool
	wake    chan struct{}
	done    chan struct{}
}

type codeActionRequest struct {
	id            jsonrpc2.ID
	uri           string
	automatic     bool
	ctx           context.Context
	cancelContext context.CancelFunc
	publishMu     sync.Mutex
	run           func(context.Context) (any, error)
	reply         func(any, error)
}

// Cancellation and final validation/publication share one linearization
// boundary. Cancellation may interrupt computation, but cannot slip between
// the final ctx/revision check in reply and the actual response write.
func (r *codeActionRequest) cancel() {
	r.publishMu.Lock()
	defer r.publishMu.Unlock()
	r.cancelContext()
}

func (r *codeActionRequest) respond(result any, err error) {
	r.publishMu.Lock()
	defer r.publishMu.Unlock()
	if canceled := r.ctx.Err(); canceled != nil {
		result, err = nil, canceled
	}
	r.reply(result, err)
}

func newCodeActionRequests() *codeActionRequests {
	q := &codeActionRequests{wake: make(chan struct{}, 1), done: make(chan struct{})}
	go q.work()
	return q
}

func (q *codeActionRequests) enqueue(request *codeActionRequest) {
	q.mu.Lock()
	var canceled []*codeActionRequest
	if request.automatic {
		canceled = q.cancelLocked(func(old *codeActionRequest) bool {
			return old.automatic && old.uri == request.uri
		})
	}
	rejected := q.stopped || len(q.pending) >= maxPendingCodeActions
	if !rejected {
		q.pending = append(q.pending, request)
	}
	q.mu.Unlock()
	replyCanceledCodeActions(canceled)
	if rejected {
		request.cancel()
		request.respond(nil, context.Canceled)
		return
	}
	select {
	case q.wake <- struct{}{}:
	default:
	}
}

func (q *codeActionRequests) work() {
	defer close(q.done)
	for range q.wake {
		for {
			q.mu.Lock()
			if len(q.pending) == 0 {
				stopped := q.stopped
				q.mu.Unlock()
				if stopped {
					return
				}
				break
			}
			request := q.pending[0]
			q.pending[0] = nil
			q.pending = q.pending[1:]
			q.active = request
			q.mu.Unlock()
			var result any
			err := request.ctx.Err()
			if err == nil {
				result, err = request.run(request.ctx)
			}
			request.respond(result, err)
			request.cancel()
			q.mu.Lock()
			q.active = nil
			q.mu.Unlock()
		}
	}
}

// cancelLocked removes queued work immediately, so a stream of obsolete
// requests cannot consume all queue slots while the current scan unwinds.
func (q *codeActionRequests) cancelLocked(matches func(*codeActionRequest) bool) []*codeActionRequest {
	if q.active != nil && matches(q.active) {
		q.active.cancel()
	}
	var canceled []*codeActionRequest
	kept := q.pending[:0]
	for _, request := range q.pending {
		if matches(request) {
			request.cancel()
			canceled = append(canceled, request)
		} else {
			kept = append(kept, request)
		}
	}
	clear(q.pending[len(kept):])
	q.pending = kept
	return canceled
}

func replyCanceledCodeActions(requests []*codeActionRequest) {
	for _, request := range requests {
		request.respond(nil, context.Canceled)
	}
}

func (q *codeActionRequests) cancelMatching(matches func(*codeActionRequest) bool) {
	if q == nil {
		return
	}
	q.mu.Lock()
	canceled := q.cancelLocked(matches)
	q.mu.Unlock()
	replyCanceledCodeActions(canceled)
}

func (q *codeActionRequests) stop() {
	if q == nil {
		return
	}
	q.mu.Lock()
	q.stopped = true
	canceled := q.cancelLocked(func(*codeActionRequest) bool { return true })
	q.mu.Unlock()
	replyCanceledCodeActions(canceled)
	select {
	case q.wake <- struct{}{}:
	default:
	}
	<-q.done
}

func (s *Server) cancelDocumentCodeActions(uri string) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return
	}
	key := normalizePathKey(path)
	s.codeActions.cancelMatching(func(request *codeActionRequest) bool { return request.uri == key })
}

func (s *Server) dispatchCodeAction(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) bool {
	if req.Method == "$/cancelRequest" {
		var params struct {
			ID *jsonrpc2.ID `json:"id"`
		}
		if req.Params != nil && json.Unmarshal(*req.Params, &params) == nil && params.ID != nil {
			s.codeActions.cancelMatching(func(request *codeActionRequest) bool { return request.id == *params.ID })
		}
		return true
	}
	if req.Method != "textDocument/codeAction" || req.Notif || !s.handler.IsInitialized() {
		return false
	}
	var params protocol.CodeActionParams
	if req.Params == nil || json.Unmarshal(*req.Params, &params) != nil || params.TextDocument.URI == "" {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid code action params"})
		return true
	}
	// GLSP's 3.16 context predates triggerKind. Read the optional field without
	// making protocol-neutral analysis depend on a newer protocol package.
	var trigger struct {
		Context struct {
			TriggerKind int `json:"triggerKind"`
		} `json:"context"`
	}
	_ = json.Unmarshal(*req.Params, &trigger)
	measurement := s.startPerformanceURI("textDocument/codeAction", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInternalError, Message: err.Error()})
		return true
	}
	measurement.setDocument(doc)
	requestCtx, cancel := context.WithCancel(ctx)
	s.codeActions.enqueue(&codeActionRequest{
		id: req.ID, uri: normalizePathKey(doc.Path), automatic: trigger.Context.TriggerKind == 2,
		ctx: requestCtx, cancelContext: cancel,
		run: func(runCtx context.Context) (any, error) {
			if s.performanceHook != nil {
				s.performanceHook("codeAction", doc.Path)
			}
			return s.codeActionForDocument(runCtx, doc, &params)
		},
		reply: func(result any, err error) {
			// Publish under the lifecycle lock: a didChange/close/reopen cannot
			// retire this revision between validation and sending its edits.
			unlock := s.lockDocumentLifecycle(doc.URI)
			defer unlock()
			if requestCtx.Err() != nil || !s.codeActionDocumentCurrent(doc) {
				result, err = nil, context.Canceled
			}
			count := 0
			if actions, ok := result.([]protocol.CodeAction); ok {
				count = len(actions)
			}
			measurement.finish(count, err)
			if err != nil {
				code := int64(jsonrpc2.CodeInternalError)
				if errors.Is(err, context.Canceled) || errors.Is(err, context.DeadlineExceeded) {
					code = -32800 // LSP RequestCancelled
				}
				_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: code, Message: err.Error()})
				return
			}
			_ = conn.Reply(ctx, req.ID, result)
		},
	})
	return true
}

func (s *Server) codeActionDocumentCurrent(doc intel.Document) bool {
	s.docs.mu.RLock()
	defer s.docs.mu.RUnlock()
	entry, ok := s.docs.docs[normalizePathKey(doc.Path)]
	return !s.docs.closed && ok && entry.snapshot == doc.Snapshot && entry.currentVersion() == doc.Version
}
