package oforms

import (
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"maps"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

// ErrUnsupportedMutation identifies a decoded or retained persistence change.
// Issue #881 supports lossless no-op replay only; controlled model mutation is
// introduced by the later FormSpec compiler.
var ErrUnsupportedMutation = errors.New("oforms: model mutation is not supported")

// SerializedForm is one validated Designer storage ready to add to a CFB
// writer. Paths are absolute within vbaProject.bin and byte slices are owned by
// this value.
type SerializedForm struct {
	Streams  map[string][]byte
	Storages map[string]cfb.StorageMeta
}

// SerializeForm validates and replays one Form accepted by ReadForm. It does
// not encode semantic edits: both decoded state and retained persistence bytes
// must still match the read-time model.
func SerializeForm(form *Form, codePage uint16) (*SerializedForm, error) {
	if form == nil {
		return nil, parseError("", "", 0, "Form", "nil form")
	}
	if len(form.Levels) == 0 || form.Levels[0] == nil {
		return nil, parseError(form.Name, "", 0, "Form", "form has no root level")
	}
	root := form.Levels[0].Path
	if root == "" || strings.Contains(root, "/") || form.Name != root {
		return nil, parseError(root, "", 0, "Form", "root storage %q does not match form name %q", root, form.Name)
	}

	serialized := &SerializedForm{
		Streams:  make(map[string][]byte),
		Storages: make(map[string]cfb.StorageMeta),
	}
	for _, level := range form.Levels {
		if err := appendLevel(serialized, root, level); err != nil {
			return nil, err
		}
	}

	writer := cfb.NewWriter()
	for _, path := range slices.Sorted(maps.Keys(serialized.Storages)) {
		writer.AddStorage(splitStoragePath(path), serialized.Storages[path])
	}
	for _, path := range slices.Sorted(maps.Keys(serialized.Streams)) {
		writer.AddStream(strings.Split(path, "/"), serialized.Streams[path])
	}
	body, err := writer.Bytes()
	if err != nil {
		return nil, parseError(root, "", 0, "container", "%v", err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		return nil, parseError(root, "", 0, "container", "%v", err)
	}
	reparsed, err := ReadForm(container, root, codePage)
	if err != nil {
		return nil, err
	}

	currentSignature, err := modelSignature(form)
	if err != nil {
		return nil, fmt.Errorf("oforms: signature form %q: %w", form.Name, err)
	}
	if form.sourceSignature == ([32]byte{}) || currentSignature != form.sourceSignature {
		return nil, fmt.Errorf("%w: form %q differs from its read-time state", ErrUnsupportedMutation, form.Name)
	}
	if reparsed.sourceSignature != form.sourceSignature {
		return nil, parseError(root, "", 0, "serialization", "replayed model differs from the accepted source model")
	}
	return serialized, nil
}

func appendLevel(serialized *SerializedForm, root string, level *Level) error {
	if level == nil {
		return parseError(root, "", 0, "container", "nil level")
	}
	if level.Path != root && !strings.HasPrefix(level.Path, root+"/") {
		return parseError(level.Path, "", 0, "container", "storage is outside root %q", root)
	}
	if _, exists := serialized.Storages[level.Path]; exists {
		return parseError(level.Path, "", 0, "container", "duplicate storage path")
	}
	serialized.Storages[level.Path] = level.StorageMeta

	if err := addSerializedStream(serialized, level.Path+"/"+level.FStreamName, level.FRaw); err != nil {
		return err
	}
	if err := addSerializedStream(serialized, level.Path+"/"+level.OStreamName, level.ORaw); err != nil {
		return err
	}
	if level.HasXStream {
		if err := addSerializedStream(serialized, level.Path+"/"+level.XStreamName, level.XRaw); err != nil {
			return err
		}
	}
	if level.HasCompObj {
		if err := addSerializedStream(serialized, level.Path+"/"+level.CompObjName, level.CompObjRaw); err != nil {
			return err
		}
	}
	if level.HasVBFrame {
		if err := addSerializedStream(serialized, level.Path+"/"+level.VBFrameName, level.VBFrameRaw); err != nil {
			return err
		}
	}
	for _, name := range slices.Sorted(maps.Keys(level.ExtraStreams)) {
		if name == "" || strings.Contains(name, "/") {
			return parseError(level.Path, name, 0, "opaque stream", "invalid direct-child stream name")
		}
		if err := addSerializedStream(serialized, level.Path+"/"+name, level.ExtraStreams[name]); err != nil {
			return err
		}
	}
	return nil
}

func addSerializedStream(serialized *SerializedForm, path string, body []byte) error {
	storagePath, stream := "", path
	if index := strings.LastIndexByte(path, '/'); index >= 0 {
		storagePath, stream = path[:index], path[index+1:]
	}
	if len(body) > maxDesignerStreamSize {
		return parseError(storagePath, stream, 0, "stream", "stream exceeds %d-byte limit", maxDesignerStreamSize)
	}
	if _, exists := serialized.Streams[path]; exists {
		return parseError(path, "", 0, "stream", "duplicate stream path")
	}
	serialized.Streams[path] = bytes.Clone(body)
	return nil
}

func splitStoragePath(path string) []string {
	if path == "" {
		return nil
	}
	return strings.Split(path, "/")
}

type formSignatureView struct {
	Name           string
	DesignerSource StoredText
	CompObj        compObjSignatureView
	Levels         []levelSignatureView
	Controls       []controlSignatureView
}

type levelSignatureView struct {
	Path          string
	StorageMeta   cfb.StorageMeta
	Record        *recordSignatureView
	Sites         []siteSignatureView
	Controls      []controlSignatureView
	FRaw          []byte
	FStreamName   string
	ORaw          []byte
	OStreamName   string
	XRaw          []byte
	XStreamName   string
	HasXStream    bool
	CompObj       *compObjSignatureView
	CompObjRaw    []byte
	CompObjName   string
	HasCompObj    bool
	VBFrameRaw    []byte
	VBFrameName   string
	HasVBFrame    bool
	ExtraStreams  map[string][]byte
	MouseIconRaw  []byte
	FontRaw       []byte
	PictureRaw    []byte
	ClassTable    []ClassInfo
	ClassTableRaw []byte
	DepthsRaw     []byte
	TrailingRaw   []byte
}

type controlSignatureView struct {
	Name             string
	Kind             string
	ID               int32
	CLSIDCacheIndex  uint16
	TabIndex         *int16
	ObjectStreamSize uint32
	Depth            uint8
	SiteType         uint8
	Site             *siteSignatureView
	Record           *recordSignatureView
	OpaqueRaw        []byte
	Children         []controlSignatureView
	LevelPath        *string
	TabStrip         *TabStrip
	MultiPage        *multiPageSignatureView
}

type multiPageSignatureView struct {
	HiddenID       int32
	PageIDs        []int32
	Reserved       []byte
	PageProperties map[int32][]byte
	Properties     *recordSignatureView
}

type siteSignatureView struct {
	Version  uint16
	Mask     uint32
	Values   map[string]int64
	Strings  map[string]StoredString
	Position *Position
	Padding  map[string][]byte
	Raw      []byte
}

type recordSignatureView struct {
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
	TextProps *recordSignatureView
	TailRaw   []byte
	Raw       []byte
}

type compObjSignatureView struct {
	ByteOrderVersion uint32
	FormatVersion    uint32
	Header           []byte
	UserType         StoredString
	ClipboardFormat  StoredString
	TailRaw          []byte
	Raw              []byte
}

func modelSignature(form *Form) ([32]byte, error) {
	// Derived references must still point into the signed ownership tree.
	// Hash their order and semantic state without introducing reference cycles.
	for _, level := range form.Levels {
		if level == nil {
			continue
		}
		for _, control := range level.Controls {
			if control == nil || control.MultiPage == nil {
				continue
			}
			state := control.MultiPage
			children := make(map[*Control]bool, len(control.Children))
			for _, child := range control.Children {
				children[child] = true
			}
			if state.Hidden == nil || !children[state.Hidden] {
				return [32]byte{}, ErrUnsupportedMutation
			}
			for _, page := range state.Pages {
				if page == nil || !children[page] {
					return [32]byte{}, ErrUnsupportedMutation
				}
			}
		}
	}
	view := formSignatureView{
		Name: form.Name, DesignerSource: form.DesignerSource,
		CompObj:  compObjSignature(form.CompObj),
		Levels:   make([]levelSignatureView, 0, len(form.Levels)),
		Controls: controlSignatures(form.Controls),
	}
	for _, level := range form.Levels {
		if level == nil {
			view.Levels = append(view.Levels, levelSignatureView{})
			continue
		}
		levelView := levelSignatureView{
			Path: level.Path, StorageMeta: level.StorageMeta,
			Record: recordSignature(level.Record), Sites: siteSignatures(level.Sites),
			Controls: controlSignatures(level.Controls), FRaw: level.FRaw, FStreamName: level.FStreamName,
			ORaw: level.ORaw, OStreamName: level.OStreamName,
			XRaw: level.XRaw, XStreamName: level.XStreamName, HasXStream: level.HasXStream,
			CompObjRaw: level.CompObjRaw, CompObjName: level.CompObjName, HasCompObj: level.HasCompObj,
			VBFrameRaw: level.VBFrameRaw, VBFrameName: level.VBFrameName, HasVBFrame: level.HasVBFrame,
			ExtraStreams: level.ExtraStreams, MouseIconRaw: level.MouseIconRaw,
			FontRaw: level.FontRaw, PictureRaw: level.PictureRaw, ClassTable: level.ClassTable,
			ClassTableRaw: level.ClassTableRaw, DepthsRaw: level.DepthsRaw, TrailingRaw: level.TrailingRaw,
		}
		if level.CompObj != nil {
			value := compObjSignature(*level.CompObj)
			levelView.CompObj = &value
		}
		view.Levels = append(view.Levels, levelView)
	}
	body, err := json.Marshal(view)
	if err != nil {
		return [32]byte{}, err
	}
	return sha256.Sum256(body), nil
}

func controlSignatures(controls []*Control) []controlSignatureView {
	if controls == nil {
		return nil
	}
	result := make([]controlSignatureView, 0, len(controls))
	for _, control := range controls {
		if control == nil {
			result = append(result, controlSignatureView{})
			continue
		}
		view := controlSignatureView{
			Name: control.Name, Kind: control.Kind, ID: control.ID,
			CLSIDCacheIndex: control.CLSIDCacheIndex, TabIndex: control.TabIndex,
			ObjectStreamSize: control.ObjectStreamSize, Depth: control.Depth,
			SiteType: control.SiteType, Record: recordSignature(control.Record),
			OpaqueRaw: control.OpaqueRaw, Children: controlSignatures(control.Children),
			TabStrip: control.TabStrip,
		}
		if state := control.MultiPage; state != nil {
			multi := &multiPageSignatureView{Reserved: state.Reserved, PageProperties: state.PageProperties, Properties: recordSignature(state.Properties)}
			if state.Hidden != nil {
				multi.HiddenID = state.Hidden.ID
			}
			for _, page := range state.Pages {
				if page != nil {
					multi.PageIDs = append(multi.PageIDs, page.ID)
				}
			}
			view.MultiPage = multi
		}
		if control.Site != nil {
			site := siteSignature(control.Site)
			view.Site = &site
		}
		if control.Level != nil {
			view.LevelPath = new(control.Level.Path)
		}
		result = append(result, view)
	}
	return result
}

func siteSignatures(sites []*Site) []siteSignatureView {
	if sites == nil {
		return nil
	}
	result := make([]siteSignatureView, 0, len(sites))
	for _, site := range sites {
		result = append(result, siteSignature(site))
	}
	return result
}

func siteSignature(site *Site) siteSignatureView {
	if site == nil {
		return siteSignatureView{}
	}
	return siteSignatureView{
		Version: site.Version, Mask: site.Mask, Values: site.Values,
		Strings: site.Strings, Position: site.Position, Padding: site.Padding, Raw: site.Raw,
	}
}

func recordSignature(record *Record) *recordSignatureView {
	if record == nil {
		return nil
	}
	return &recordSignatureView{
		Type: record.Type, Minor: record.Minor, Major: record.Major,
		Mask: record.Mask, MaskWidth: record.MaskWidth, Values: record.Values,
		Strings: record.Strings, Sizes: record.Sizes, Arrays: record.Arrays,
		Pictures: record.Pictures, Padding: record.Padding,
		TextProps: recordSignature(record.TextProps), TailRaw: record.TailRaw, Raw: record.Raw,
	}
}

func compObjSignature(compObj CompObj) compObjSignatureView {
	return compObjSignatureView(compObj)
}
