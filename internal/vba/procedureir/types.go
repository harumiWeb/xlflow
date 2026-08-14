// Package procedureir normalizes VBA procedure syntax into protocol-neutral
// Go values. Values returned by this package never retain tree-sitter nodes or
// source byte slices.
package procedureir

import vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"

type BuildOptions struct {
	RootDir    string
	Path       string
	ModuleName string
	ModuleKind string
}

type ParseSummary struct {
	HasError   bool `json:"hasError"`
	HasMissing bool `json:"hasMissing"`
}

type DocumentIR struct {
	Path             string            `json:"path"`
	ModuleName       string            `json:"moduleName"`
	ModuleKind       string            `json:"moduleKind"`
	ModuleAttributes []ModuleAttribute `json:"moduleAttributes,omitempty"`
	Parse            ParseSummary      `json:"parse"`
	Declarations     []Declaration     `json:"declarations"`
	Procedures       []ProcedureIR     `json:"procedures"`
	TypeReferences   []TypeReference   `json:"-"`
}

// ModuleAttribute is a raw exported-module attribute such as VB_Exposed.
// Keeping the name and value in the IR avoids coupling consumers to the
// parser's attribute-node representation.
type ModuleAttribute struct {
	Name  string `json:"name"`
	Value string `json:"value"`
}

type ProcedureIR struct {
	Symbol       ProcedureSymbol       `json:"symbol"`
	Declarations []Declaration         `json:"declarations"`
	Statements   []Statement           `json:"statements"`
	Expressions  []Expression          `json:"expressions"`
	Calls        []CallSite            `json:"calls"`
	RaiseEvents  []RaiseEventReference `json:"raiseEvents,omitempty"`
	Accesses     []VariableAccess      `json:"accesses"`
}

type ProcedureKind string

const (
	ProcedureSub         ProcedureKind = "sub"
	ProcedureFunction    ProcedureKind = "function"
	ProcedureProperty    ProcedureKind = "property"
	ProcedurePropertyGet ProcedureKind = "property_get"
	ProcedurePropertyLet ProcedureKind = "property_let"
	ProcedurePropertySet ProcedureKind = "property_set"
)

type ProcedureSymbol struct {
	Name             string         `json:"name"`
	QualifiedName    string         `json:"qualifiedName"`
	Kind             ProcedureKind  `json:"kind"`
	Visibility       string         `json:"visibility,omitempty"`
	Parameters       []Parameter    `json:"parameters"`
	ReturnType       string         `json:"returnType,omitempty"`
	IsArray          bool           `json:"isArray,omitempty"`
	ValueShape       ValueShapeKind `json:"valueShape,omitempty"`
	ArrayBounds      []ArrayBound   `json:"arrayBounds,omitempty"`
	DeclarationRange vbaast.Range   `json:"declarationRange"`
	BodyRange        vbaast.Range   `json:"bodyRange"`
	IsEventHandler   bool           `json:"isEventHandler"`
	EventKind        string         `json:"eventKind,omitempty"`
	Recovered        bool           `json:"recovered,omitempty"`
	// ConditionalBranches records the mutually-exclusive compilation branches
	// that contribute this procedure header.  The first (consequence) branch
	// remains the canonical signature in Parameters; consumers that require a
	// single active declaration must treat a multi-branch symbol conservatively.
	ConditionalBranches []ConditionalBranch `json:"conditionalBranches,omitempty"`
}

// ArrayShape describes the syntactic shape of a parameter array declarator.
// It intentionally distinguishes an omitted bounds clause from an explicit
// dynamic () clause, and preserves malformed parser-recovery input as
// invalid instead of silently treating it as dynamic.
type ArrayShape string

const (
	ArrayShapeNone    ArrayShape = "none"
	ArrayShapeDynamic ArrayShape = "dynamic"
	ArrayShapeBounded ArrayShape = "bounded"
	ArrayShapeInvalid ArrayShape = "invalid"
)

// ValueShapeKind is the declaration-level value shape shared by consumers
// that need to distinguish scalar values from array-capable values.  It is
// intentionally separate from ArrayShape, which describes only the syntax of
// a procedure parameter's bounds clause.
type ValueShapeKind string

const (
	ValueShapeUnknown      ValueShapeKind = "unknown"
	ValueShapeScalar       ValueShapeKind = "scalar"
	ValueShapeVariant      ValueShapeKind = "variant"
	ValueShapeFixedArray   ValueShapeKind = "fixed-array"
	ValueShapeDynamicArray ValueShapeKind = "dynamic-array"
)

// ArrayBound is one dimension in a parameter's bounds clause.  A bound with
// only Expression set is the shorthand form (for example, `(10)`); Lower and
// Upper are populated for an explicit `lower To upper` form.
type ArrayBound struct {
	Expression string        `json:"expression,omitempty"`
	Lower      string        `json:"lower,omitempty"`
	Upper      string        `json:"upper,omitempty"`
	Range      vbaast.Range  `json:"range"`
	LowerRange *vbaast.Range `json:"lowerRange,omitempty"`
	UpperRange *vbaast.Range `json:"upperRange,omitempty"`
	Recovered  bool          `json:"recovered,omitempty"`
}

// ConditionalBranch identifies one possible #If/#ElseIf/#Else declaration
// branch. Group is stable within a document build (the source offset of the
// conditional node); Branch is zero-based with the consequence as branch 0.
// Condition contains the source expression for #If and #ElseIf branches. It is
// empty for #Else because that branch has no standalone expression; consumers
// should use Group and Branch to determine mutual exclusivity.
type ConditionalBranch struct {
	Group     string       `json:"group"`
	Condition string       `json:"condition,omitempty"`
	Branch    int          `json:"branch"`
	Range     vbaast.Range `json:"range"`
}

type Parameter struct {
	Name            string        `json:"name"`
	Type            string        `json:"type,omitempty"`
	Passing         string        `json:"passing,omitempty"`
	PassingExplicit bool          `json:"passingExplicit,omitempty"`
	Optional        bool          `json:"optional,omitempty"`
	ParamArray      bool          `json:"paramArray,omitempty"`
	Default         string        `json:"default,omitempty"`
	HasDefault      bool          `json:"hasDefault,omitempty"`
	DefaultRange    *vbaast.Range `json:"defaultRange,omitempty"`
	Range           vbaast.Range  `json:"range"`
	IsArray         bool          `json:"isArray,omitempty"`
	ArrayShape      ArrayShape    `json:"arrayShape,omitempty"`
	BoundsRange     *vbaast.Range `json:"boundsRange,omitempty"`
	ArrayBounds     []ArrayBound  `json:"arrayBounds,omitempty"`
	// ValueShape is the normalized declaration shape. ArrayShape remains the
	// source-compatible syntax fact used by signature validation.
	ValueShape ValueShapeKind `json:"valueShape,omitempty"`
	Recovered  bool           `json:"recovered,omitempty"`
}

type Declaration struct {
	ID                  int                 `json:"id"`
	Name                string              `json:"name"`
	Type                string              `json:"type,omitempty"`
	Parameters          []Parameter         `json:"parameters,omitempty"`
	Scope               SymbolScope         `json:"scope"`
	Visibility          string              `json:"visibility,omitempty"`
	Kind                string              `json:"kind"`
	Parent              string              `json:"parent,omitempty"`
	IsArray             bool                `json:"isArray,omitempty"`
	IsObject            bool                `json:"isObject,omitempty"`
	IsConst             bool                `json:"isConst,omitempty"`
	ParamArray          bool                `json:"paramArray,omitempty"`
	ValueShape          ValueShapeKind      `json:"valueShape,omitempty"`
	ArrayBounds         []ArrayBound        `json:"arrayBounds,omitempty"`
	Range               vbaast.Range        `json:"range"`
	Recovered           bool                `json:"recovered,omitempty"`
	ConditionalBranches []ConditionalBranch `json:"conditionalBranches,omitempty"`
}

type StatementKind string

const (
	StatementDeclaration StatementKind = "declaration"
	StatementAssignment  StatementKind = "assignment"
	StatementSet         StatementKind = "set_assignment"
	StatementCall        StatementKind = "call"
	StatementRaiseEvent  StatementKind = "raise_event"
	StatementIf          StatementKind = "if"
	StatementElseIf      StatementKind = "elseif"
	StatementElse        StatementKind = "else"
	StatementSelect      StatementKind = "select"
	StatementCase        StatementKind = "case"
	StatementFor         StatementKind = "for"
	StatementForEach     StatementKind = "for_each"
	StatementDo          StatementKind = "do"
	StatementWhile       StatementKind = "while"
	StatementWith        StatementKind = "with"
	StatementReDim       StatementKind = "redim"
	StatementExit        StatementKind = "exit"
	StatementLabel       StatementKind = "label"
	StatementGoTo        StatementKind = "goto"
	StatementOnError     StatementKind = "on_error"
	StatementResume      StatementKind = "resume"
	StatementEnd         StatementKind = "end"
	StatementUnknown     StatementKind = "unknown"
	StatementRecovered   StatementKind = "recovered"
)

type BranchRole string

const (
	BranchThen BranchRole = "then"
	BranchElse BranchRole = "else"
)

type LoopTest string

const (
	LoopPreWhile  LoopTest = "pre_while"
	LoopPreUntil  LoopTest = "pre_until"
	LoopPostWhile LoopTest = "post_while"
	LoopPostUntil LoopTest = "post_until"
)

type TransferKind string

const (
	TransferGoto              TransferKind = "goto"
	TransferExitSub           TransferKind = "exit_sub"
	TransferExitFunction      TransferKind = "exit_function"
	TransferExitProperty      TransferKind = "exit_property"
	TransferExitFor           TransferKind = "exit_for"
	TransferExitDo            TransferKind = "exit_do"
	TransferOnErrorGoto       TransferKind = "on_error_goto"
	TransferOnErrorResumeNext TransferKind = "on_error_resume_next"
	TransferOnErrorDisable    TransferKind = "on_error_disable"
	TransferResumeRetry       TransferKind = "resume_retry"
	TransferResumeNext        TransferKind = "resume_next"
	TransferResumeLabel       TransferKind = "resume_label"
	TransferTerminate         TransferKind = "terminate"
)

type ControlFlowMetadata struct {
	Branch   BranchRole   `json:"branch,omitempty"`
	CaseElse bool         `json:"caseElse,omitempty"`
	Loop     LoopTest     `json:"loop,omitempty"`
	Transfer TransferKind `json:"transfer,omitempty"`
	Target   string       `json:"target,omitempty"`
	// TargetRange identifies only the label token. Range continues to cover
	// the complete control-flow statement.
	TargetRange vbaast.Range `json:"targetRange,omitempty"`
	// LoopVariable is the control variable declared by a For/For Each header.
	LoopVariable      string       `json:"loopVariable,omitempty"`
	LoopVariableRange vbaast.Range `json:"loopVariableRange,omitempty"`
	// NextVariables are the optional variables on the matching Next closer.
	// NextVariableRanges is parallel to this slice.
	NextVariables      []string       `json:"nextVariables,omitempty"`
	NextVariableRanges []vbaast.Range `json:"nextVariableRanges,omitempty"`
	Range              vbaast.Range   `json:"range"`
}

type Statement struct {
	ID                  int                  `json:"id"`
	ParentID            int                  `json:"parentId,omitempty"`
	Kind                StatementKind        `json:"kind"`
	SyntaxKind          string               `json:"syntaxKind"`
	Text                string               `json:"text"`
	Range               vbaast.Range         `json:"range"`
	Recovered           bool                 `json:"recovered,omitempty"`
	Label               string               `json:"label,omitempty"`
	LabelRange          vbaast.Range         `json:"labelRange,omitempty"`
	Target              *Expression          `json:"target,omitempty"`
	Value               *Expression          `json:"value,omitempty"`
	Condition           *Expression          `json:"condition,omitempty"`
	TargetID            int                  `json:"targetId,omitempty"`
	ValueID             int                  `json:"valueId,omitempty"`
	ConditionID         int                  `json:"conditionId,omitempty"`
	ExpressionIDs       []int                `json:"expressionIds,omitempty"`
	ConditionalBranches []ConditionalBranch  `json:"conditionalBranches,omitempty"`
	Control             *ControlFlowMetadata `json:"control,omitempty"`
}

type ExpressionKind string

const (
	ExpressionIdentifier  ExpressionKind = "identifier"
	ExpressionLiteral     ExpressionKind = "literal"
	ExpressionMember      ExpressionKind = "member_access"
	ExpressionCall        ExpressionKind = "call"
	ExpressionNew         ExpressionKind = "new"
	ExpressionUnary       ExpressionKind = "unary"
	ExpressionBinary      ExpressionKind = "binary"
	ExpressionParentheses ExpressionKind = "parenthesized"
	ExpressionUnknown     ExpressionKind = "unknown"
)

type Expression struct {
	ID          int            `json:"id"`
	ParentID    int            `json:"parentId,omitempty"`
	StatementID int            `json:"statementId,omitempty"`
	Kind        ExpressionKind `json:"kind"`
	SyntaxKind  string         `json:"syntaxKind"`
	Text        string         `json:"text"`
	Range       vbaast.Range   `json:"range"`
	Children    []int          `json:"children,omitempty"`
	Recovered   bool           `json:"recovered,omitempty"`
}

type AccessMode string

const (
	AccessRead      AccessMode = "read"
	AccessWrite     AccessMode = "write"
	AccessReadWrite AccessMode = "read_write"
)

type SymbolScope string

const (
	ScopeParameter  SymbolScope = "parameter"
	ScopeLocal      SymbolScope = "local"
	ScopeModule     SymbolScope = "module"
	ScopeProject    SymbolScope = "project"
	ScopeUnresolved SymbolScope = "unresolved"
)

type VariableAccess struct {
	Name         string           `json:"name"`
	Mode         AccessMode       `json:"mode"`
	Scope        SymbolScope      `json:"scope"`
	Range        vbaast.Range     `json:"range"`
	StatementID  int              `json:"statementId,omitempty"`
	ExpressionID int              `json:"expressionId,omitempty"`
	Resolution   SymbolResolution `json:"resolution"`
}

type Callee struct {
	Text     string  `json:"text"`
	BaseName string  `json:"baseName"`
	Receiver *string `json:"receiver,omitempty"`
	Member   string  `json:"member"`
}

type Arguments struct {
	Count         int             `json:"count"`
	Named         []NamedArgument `json:"named"`
	ExpressionIDs []int           `json:"-"`
}

type NamedArgument struct {
	Name         string `json:"name"`
	ValueText    string `json:"valueText"`
	ExpressionID int    `json:"-"`
}

type CallSite struct {
	ID           int            `json:"id"`
	File         string         `json:"file"`
	Module       string         `json:"module"`
	Caller       ProcedureRef   `json:"caller"`
	Callee       Callee         `json:"callee"`
	Arguments    Arguments      `json:"arguments"`
	Range        vbaast.Range   `json:"range"`
	StatementID  int            `json:"statementId,omitempty"`
	ExpressionID int            `json:"expressionId,omitempty"`
	IsRaiseEvent bool           `json:"isRaiseEvent,omitempty"`
	Resolution   CallResolution `json:"resolution"`
	// NonCallableNames carries lexical declarations that shadow project
	// procedures (for example `Dim Run` followed by `Run()`). It is an
	// internal resolver fact and never changes wire/inspect output.
	NonCallableNames []string `json:"-"`
}

type ProcedureRef struct {
	Name          string        `json:"name"`
	Kind          ProcedureKind `json:"kind"`
	QualifiedName string        `json:"qualifiedName"`
}

type ResolutionStatus string

const (
	ResolutionNotAttempted ResolutionStatus = "not_attempted"
	ResolutionMatched      ResolutionStatus = "matched"
	ResolutionNonCallable  ResolutionStatus = "non_callable"
	ResolutionAmbiguous    ResolutionStatus = "ambiguous"
	ResolutionUnresolved   ResolutionStatus = "unresolved"
	ResolutionExternal     ResolutionStatus = "external"
	ResolutionBuiltinLike  ResolutionStatus = "builtin_like"
	ResolutionMemberCall   ResolutionStatus = "member_call"
	ResolutionDynamic      ResolutionStatus = "dynamic"
	ResolutionIncomplete   ResolutionStatus = "incomplete"
)

type Candidate struct {
	QualifiedName string `json:"qualifiedName"`
	Kind          string `json:"kind"`
	File          string `json:"file"`
	Line          int    `json:"line"`
}

type CallResolution struct {
	Status     ResolutionStatus `json:"status"`
	Candidates []Candidate      `json:"candidates,omitempty"`
	// ProjectLocal is set only when the resolver can prove that a negative
	// result belongs to this project (for example Module.Missing or
	// Me.Missing). It is intentionally not serialized: inspect-call JSON is a
	// compatibility surface and should continue to expose only status and
	// candidates.
	ProjectLocal bool `json:"-"`
}

type SymbolReference struct {
	Name   string
	Module string
	Caller ProcedureRef
	Range  vbaast.Range
}

// EnumMemberReference describes an enum constant lookup. Enum is empty for a
// bare member reference and populated for an explicitly qualified reference.
// Keeping this separate from CallSite lets callers use the same visibility and
// completeness policy for expression-level enum resolution.
type EnumMemberReference struct {
	Name   string
	Enum   string
	Module string
	Caller ProcedureRef
	Range  vbaast.Range
}

type EnumResolution struct {
	Status     ResolutionStatus `json:"status"`
	Candidates []Candidate      `json:"candidates,omitempty"`
}

// RaiseEventReference is the syntax-local fact for a RaiseEvent statement.
// Range covers only the event identifier, making it suitable for a precise
// diagnostic while the surrounding CallSite retains legacy call-graph data.
type RaiseEventReference struct {
	Name                string              `json:"name"`
	Module              string              `json:"module"`
	Caller              ProcedureRef        `json:"caller"`
	Range               vbaast.Range        `json:"range"`
	Arguments           Arguments           `json:"arguments"`
	Recovered           bool                `json:"recovered,omitempty"`
	ConditionalBranches []ConditionalBranch `json:"conditionalBranches,omitempty"`
	Resolution          SymbolResolution    `json:"resolution"`
}

type SymbolResolution struct {
	Status     ResolutionStatus `json:"status,omitempty"`
	Scope      SymbolScope      `json:"scope"`
	Candidates []Candidate      `json:"candidates,omitempty"`
}

// Resolver provides project-dependent overlays. Implementations must be safe
// to call repeatedly; Resolve never mutates either its input or the resolver.
type Resolver interface {
	ResolveCall(CallSite) CallResolution
	ResolveSymbol(SymbolReference) SymbolResolution
}

type ResolverSymbol struct {
	Name       string
	Type       string
	Module     string
	ModuleKind string
	Kind       string
	Visibility string
	File       string
	Line       int
	// Parent identifies the containing Enum for enum members. It is kept
	// generic because the symbols package also uses Parent for procedure/local
	// ownership.
	Parent              string
	Recovered           bool
	ConditionalBranches []ConditionalBranch
	IsArray             bool
	IsConst             bool
	ValueShape          ValueShapeKind
}

type TypeReference struct {
	Kind   string        `json:"kind"`
	File   string        `json:"file"`
	Module string        `json:"module"`
	Caller *ProcedureRef `json:"caller,omitempty"`
	Target string        `json:"target"`
	Range  vbaast.Range  `json:"range"`
	Parse  ParseSummary  `json:"parse"`
}
