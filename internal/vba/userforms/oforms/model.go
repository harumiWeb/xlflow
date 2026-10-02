// Package oforms reads the binary Microsoft Forms persistence streams stored
// beside VBA modules in vbaProject.bin.
package oforms

import "github.com/harumiWeb/xlflow/internal/pack/cfb"

// Form is one lossless UserForm designer model. It is deliberately separate
// from the user-facing forms.FormSpec representation.
type Form struct {
	Name           string
	DesignerSource StoredText
	CompObj        CompObj
	Levels         []*Level
	Controls       []*Control

	// sourceSignature binds the decoded model to the exact persistence state
	// accepted by ReadForm. Issue #881 is intentionally a no-op serializer:
	// callers must not mutate either decoded fields or retained raw bytes until
	// the explicit mutation layer is added.
	sourceSignature [32]byte
}

// Level is one parent-control storage: the form itself or a nested container.
type Level struct {
	Path        string
	StorageMeta cfb.StorageMeta
	Record      *Record
	Sites       []*Site
	Controls    []*Control

	FRaw         []byte
	ORaw         []byte
	XRaw         []byte
	HasXStream   bool
	CompObj      *CompObj
	CompObjRaw   []byte
	HasCompObj   bool
	VBFrameRaw   []byte
	HasVBFrame   bool
	ExtraStreams map[string][]byte

	MouseIconRaw  []byte
	FontRaw       []byte
	PictureRaw    []byte
	ClassTable    []ClassInfo
	ClassTableRaw []byte
	DepthsRaw     []byte
	TrailingRaw   []byte
}

// Control is an embedded control associated with one OleSite record.
type Control struct {
	Name             string
	Kind             string
	ID               int32
	CLSIDCacheIndex  uint16
	TabIndex         *int16
	ObjectStreamSize uint32
	Depth            uint8
	SiteType         uint8
	Site             *Site
	Record           *Record
	OpaqueRaw        []byte
	Children         []*Control
	Level            *Level
}

// Site is an OleSiteConcreteControl and its raw persistence details.
type Site struct {
	Version  uint16
	Mask     uint32
	Values   map[string]int64
	Strings  map[string]StoredString
	Position *Position
	Padding  map[string][]byte
	Raw      []byte
}

// Record is a FormControl, embedded-control, or TextProps property record.
type Record struct {
	Type      string
	Minor     uint8
	Major     uint8
	Mask      uint64
	MaskWidth int
	Values    map[string]int64
	Strings   map[string]StoredString
	Sizes     map[string]Size
	Arrays    map[string][]byte
	Pictures  map[string][]byte
	Padding   map[string][]byte
	TextProps *Record
	TailRaw   []byte
	Raw       []byte
}

// StoredString retains both decoded text and its exact persisted bytes.
type StoredString struct {
	Text       string
	Compressed bool
	Raw        []byte
}

// StoredText is an MBCS text stream together with its source bytes.
type StoredText struct {
	Text string
	Raw  []byte
}

// Size stores an MSForms size in HIMETRIC units.
type Size struct {
	Width  int32
	Height int32
}

// Position stores an MSForms parent-relative position in HIMETRIC units.
type Position struct {
	Left int32
	Top  int32
}

// ClassInfo retains one SiteClassInfo record from a FormSiteData class table.
type ClassInfo struct {
	Version uint16
	Raw     []byte
}

// CompObj is the parsed prefix of the OLE CompObj stream plus all original
// bytes. TailRaw retains fields outside the identity strings understood here.
type CompObj struct {
	ByteOrderVersion uint32
	FormatVersion    uint32
	Header           []byte
	UserType         StoredString
	ClipboardFormat  StoredString
	TailRaw          []byte
	Raw              []byte
}
