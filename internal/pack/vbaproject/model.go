package vbaproject

import (
	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

// ModuleType is the kind of a module. It distinguishes the editable kinds
// (Std/Class/Document) from Form, which is detected but not edited.
type ModuleType int

const (
	ModuleStd      ModuleType = iota // .bas, source is replaceable
	ModuleClass                      // .cls, source is replaceable
	ModuleDocument                   // ThisWorkbook/Sheet: identity preserved, source replaceable
	ModuleForm                       // .frm, detection only (not supported in v1)
)

// Module is a single module. Source is the plain source after skipping the
// p-code, decompressing, and decoding from the CODEPAGE. The remaining fields
// carry the module-level metadata of the dir PROJECTMODULES record so a
// read-modify-write round trip preserves them.
type Module struct {
	Name        string
	StreamName  string
	Type        ModuleType
	Source      string
	DocString   string                   // MODULEDOCSTRING (module description)
	HelpContext uint32                   // MODULEHELPCONTEXT
	ReadOnly    bool                     // MODULEREADONLY record present
	Private     bool                     // MODULEPRIVATE record present
	Extra       []ovba.ModuleExtraRecord // uninterpreted module records, preserved verbatim
}

// Reference is best-effort display metadata for a project reference.
// Verbatim preservation is handled by Project.ReferencesRaw, so this holds the
// name only.
type Reference struct {
	Name string
}

// ProjectProps holds project attributes derived from the dir and PROJECT streams.
type ProjectProps struct {
	ProjectID string
	Name      string
	SysKind   uint32
	LCID      uint32
	CodePage  uint16
}

// Protection holds CMG/DPB/GC verbatim (it does not decrypt them).
type Protection struct {
	CMG, DPB, GC string
	IsProtected  bool
}

// Project is the model of an entire vbaProject.bin and the central type for
// read-modify-write.
type Project struct {
	CFBFormat        cfb.Format
	Modules          []Module
	References       []Reference // best-effort, for display (not edited in v1)
	ReferencesRaw    []byte      // verbatim byte span of the dir references section (written back unchanged)
	ProjectInfoRaw   []byte      // verbatim span of dir PROJECTINFORMATION (written back unchanged)
	ProjectStreamRaw []byte      // template PROJECT stream; unrelated content is preserved while component declarations are rebuilt
	// GenerateProjectMetadata marks a fresh project whose PROJECT and PROJECTwm
	// streams must be rebuilt from the final component set at write time.
	GenerateProjectMetadata bool
	// RawStreams holds every stream the writer does not own, keyed by full
	// "/"-separated path, captured verbatim at read time and re-emitted on write.
	// Membership is structural: the first path segment is neither "VBA" (owned and
	// regenerated) nor "PROJECT" (re-emitted verbatim). This carries root-level
	// designer storages (UserForm1/f, o, \x01CompObj, \x03VBFrame), PROJECTwm, and
	// any other opaque payload through a round-trip without modeling it.
	RawStreams map[string][]byte
	Props      ProjectProps
	Protection Protection
}
