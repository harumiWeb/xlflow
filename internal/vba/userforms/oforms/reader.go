package oforms

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"maps"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

const (
	maxSitesPerForm = 65_535
	maxNestingDepth = 64

	formMaskMouseIcon = uint64(1 << 15)
	formMaskFont      = uint64(1 << 20)
	formMaskPicture   = uint64(1 << 21)

	guidStdFont   = uint32(0x0be35203)
	guidTextProps = uint32(0xafc20920)
)

type depthType struct {
	depth uint8
	typ   uint8
}

type readState struct {
	container  *cfb.Container
	codePage   uint16
	form       *Form
	totalSites int
}

// DiscoverForms returns root-level storages that contain an f stream. The
// caller remains responsible for matching them with ModuleForm entries.
func DiscoverForms(container *cfb.Container) []string {
	if container == nil {
		return nil
	}
	var forms []string
	for _, path := range container.StoragePaths() {
		if path == "" || strings.Contains(path, "/") || strings.EqualFold(path, "VBA") {
			continue
		}
		if _, ok := container.Stream(path + "/f"); ok {
			forms = append(forms, path)
		}
	}
	return forms
}

// ReadForm parses one UserForm designer storage without Excel, COM, or VBIDE.
func ReadForm(container *cfb.Container, storagePath string, codePage uint16) (*Form, error) {
	if container == nil {
		return nil, parseError(storagePath, "", 0, "Form", "nil CFB container")
	}
	storagePath = strings.Trim(storagePath, "/")
	if storagePath == "" {
		return nil, parseError(storagePath, "", 0, "Form", "empty storage path")
	}
	if _, ok := container.Storage(storagePath); !ok {
		return nil, parseError(storagePath, "", 0, "Form", "storage does not exist")
	}
	if _, err := ovba.DecodeMBCS(nil, codePage); err != nil {
		return nil, parseError(storagePath, "", 0, "Form", "%v", err)
	}
	form := &Form{Name: pathBase(storagePath)}
	state := &readState{container: container, codePage: codePage, form: form}
	controls, level, err := state.readLevel(storagePath, 0, true)
	if err != nil {
		return nil, err
	}
	form.Controls = controls
	form.CompObj = *level.CompObj
	designer, err := ovba.DecodeMBCS(level.VBFrameRaw, codePage)
	if err != nil {
		return nil, parseError(storagePath, "\x03VBFrame", 0, "VBFrame", "%v", err)
	}
	form.DesignerSource = StoredText{Text: designer, Raw: bytes.Clone(level.VBFrameRaw)}
	return form, nil
}

func (s *readState) readLevel(path string, depth int, root bool) ([]*Control, *Level, error) {
	if depth >= maxNestingDepth {
		return nil, nil, parseError(path, "", 0, "container", "nesting exceeds %d levels", maxNestingDepth)
	}
	fRaw, ok := s.container.Stream(path + "/f")
	if !ok {
		return nil, nil, parseError(path, "f", 0, "FormControl", "required stream is missing")
	}
	oRaw, ok := s.container.Stream(path + "/o")
	if !ok {
		return nil, nil, parseError(path, "o", 0, "Object stream", "required stream is missing")
	}
	if len(fRaw) > maxDesignerStreamSize || len(oRaw) > maxDesignerStreamSize {
		return nil, nil, parseError(path, "", 0, "container", "designer stream exceeds %d-byte limit", maxDesignerStreamSize)
	}
	level, err := parseFormStream(fRaw, path, s.codePage)
	if err != nil {
		return nil, nil, err
	}
	level.Path = path
	level.FRaw = bytes.Clone(fRaw)
	level.ORaw = bytes.Clone(oRaw)
	level.StorageMeta, _ = s.container.Storage(path)
	if xRaw, found := s.container.Stream(path + "/x"); found {
		if len(xRaw) > maxDesignerStreamSize {
			return nil, nil, parseError(path, "x", 0, "auxiliary stream", "stream exceeds %d-byte limit", maxDesignerStreamSize)
		}
		level.XRaw = bytes.Clone(xRaw)
	}
	if compObj, found := s.container.Stream(path + "/\x01CompObj"); found {
		parsed, err := parseCompObj(compObj, path, s.codePage)
		if err != nil {
			return nil, nil, err
		}
		level.CompObj = &parsed
		// parseCompObj enforces the stream cap before cloning. Share that
		// retained clone rather than allocating a second raw copy for the level.
		level.CompObjRaw = parsed.Raw
	} else if root {
		return nil, nil, parseError(path, "\x01CompObj", 0, "CompObj", "required stream is missing")
	}
	if vbFrame, found := s.container.Stream(path + "/\x03VBFrame"); found {
		if len(vbFrame) > maxDesignerStreamSize {
			return nil, nil, parseError(path, "\x03VBFrame", 0, "VBFrame", "stream exceeds %d-byte limit", maxDesignerStreamSize)
		}
		level.VBFrameRaw = bytes.Clone(vbFrame)
	} else if root {
		return nil, nil, parseError(path, "\x03VBFrame", 0, "VBFrame", "required stream is missing")
	}
	level.ExtraStreams, err = s.extraStreams(path)
	if err != nil {
		return nil, nil, err
	}

	s.totalSites += len(level.Sites)
	if s.totalSites > maxSitesPerForm {
		return nil, nil, parseError(path, "f", 0, "FormSiteData", "form exceeds %d sites", maxSitesPerForm)
	}
	s.form.Levels = append(s.form.Levels, level)

	childrenByID, err := s.childStorages(path)
	if err != nil {
		return nil, nil, err
	}
	var objectOffset uint64
	controls := make([]*Control, 0, len(level.Sites))
	for _, site := range level.Sites {
		objectSize := uint32(site.Values["ObjectStreamSize"])
		if objectOffset+uint64(objectSize) > uint64(len(oRaw)) {
			return nil, nil, parseError(path, "o", int(objectOffset), "ObjectStreamSize", "site %q size %d exceeds stream boundary", siteName(site), objectSize)
		}
		id := int32(site.Values["ID"])
		cacheIndex := uint16(site.Values["ClsidCacheIndex"])
		control := &Control{
			Name: siteName(site), Kind: controlKind(cacheIndex, nil), ID: id,
			CLSIDCacheIndex: cacheIndex, ObjectStreamSize: objectSize,
			Depth: uint8(site.Values["_Depth"]), SiteType: uint8(site.Values["_Type"]), Site: site,
		}
		if value, found := site.Values["TabIndex"]; found {
			tabIndex := int16(value)
			control.TabIndex = &tabIndex
		}
		childPath, hasChild := childrenByID[id]
		_, isContainer := containerCacheIndices[cacheIndex]
		switch {
		case hasChild && !isContainer:
			return nil, nil, parseError(path, "f", 0, "container ownership", "site %q with class index %d owns nested storage %q", control.Name, cacheIndex, childPath)
		case isContainer && !hasChild:
			return nil, nil, parseError(path, "f", 0, "container ownership", "container site %q (ID %d) has no matching child storage", control.Name, id)
		case hasChild:
			if objectSize != 0 {
				return nil, nil, parseError(path, "o", int(objectOffset), "container ownership", "container site %q has nonzero ObjectStreamSize %d", control.Name, objectSize)
			}
			childControls, childLevel, err := s.readLevel(childPath, depth+1, false)
			if err != nil {
				return nil, nil, err
			}
			control.Children = childControls
			control.Level = childLevel
			control.Record = childLevel.Record
			delete(childrenByID, id)
		default:
			payload := oRaw[objectOffset : objectOffset+uint64(objectSize)]
			if spec := specsByCacheIndex[cacheIndex]; spec != nil && len(payload) > 0 {
				record, consumed, err := parseRecord(payload, spec, s.codePage, len(payload))
				if err != nil {
					return nil, nil, parseError(path, "o", int(objectOffset)+consumed, spec.typeName, "%v", err)
				}
				control.Record = record
				control.Kind = controlKind(cacheIndex, record)
			} else {
				control.OpaqueRaw = bytes.Clone(payload)
			}
			objectOffset += uint64(objectSize)
		}
		controls = append(controls, control)
	}
	if objectOffset != uint64(len(oRaw)) {
		return nil, nil, parseError(path, "o", int(objectOffset), "Object stream", "sites account for %d bytes but stream has %d", objectOffset, len(oRaw))
	}
	if len(childrenByID) != 0 {
		names := slices.Sorted(maps.Values(childrenByID))
		return nil, nil, parseError(path, "", 0, "container ownership", "child storages claimed by no site: %s", strings.Join(names, ", "))
	}
	level.Controls = controls
	return controls, level, nil
}

func (s *readState) extraStreams(path string) (map[string][]byte, error) {
	prefix := path + "/"
	known := map[string]struct{}{
		"f": {}, "o": {}, "x": {}, "\x01compobj": {}, "\x03vbframe": {},
	}
	extra := map[string][]byte{}
	for _, candidate := range s.container.Paths() {
		rel, ok := strings.CutPrefix(candidate, prefix)
		if !ok || rel == "" || strings.Contains(rel, "/") {
			continue
		}
		if _, found := known[strings.ToLower(rel)]; found {
			continue
		}
		body, _ := s.container.Stream(candidate)
		if len(body) > maxDesignerStreamSize {
			return nil, parseError(path, rel, 0, "opaque stream", "stream exceeds %d-byte limit", maxDesignerStreamSize)
		}
		extra[rel] = bytes.Clone(body)
	}
	if len(extra) == 0 {
		return nil, nil
	}
	return extra, nil
}

func (s *readState) childStorages(path string) (map[int32]string, error) {
	prefix := path + "/"
	children := map[int32]string{}
	for _, candidate := range s.container.StoragePaths() {
		rel, ok := strings.CutPrefix(candidate, prefix)
		if !ok || rel == "" || strings.Contains(rel, "/") {
			continue
		}
		if len(rel) < 2 || !strings.EqualFold(rel[:1], "i") {
			return nil, parseError(path, "", 0, "container ownership", "unrecognized child storage %q", candidate)
		}
		id64, err := strconv.ParseInt(rel[1:], 10, 32)
		if err != nil || id64 < 0 {
			return nil, parseError(path, "", 0, "container ownership", "invalid child storage ID in %q", candidate)
		}
		id := int32(id64)
		if previous, exists := children[id]; exists {
			return nil, parseError(path, "", 0, "container ownership", "storages %q and %q have duplicate ID %d", previous, candidate, id)
		}
		children[id] = candidate
	}
	return children, nil
}

func parseFormStream(data []byte, path string, codePage uint16) (*Level, error) {
	if len(data) > maxDesignerStreamSize {
		return nil, parseError(path, "f", 0, "FormControl", "stream exceeds %d-byte limit", maxDesignerStreamSize)
	}
	if len(data) < 4 {
		return nil, parseError(path, "f", 0, "FormControl", "stream is shorter than its header")
	}
	cbForm := int(binary.LittleEndian.Uint16(data[2:4]))
	formEnd := 4 + cbForm
	if formEnd > len(data) {
		return nil, parseError(path, "f", 2, "FormControl", "cbForm %d exceeds stream size %d", cbForm, len(data))
	}
	record, consumed, err := parseRecord(data, &formSpec, codePage, formEnd)
	if err != nil {
		return nil, parseError(path, "f", consumed, "FormControl", "%v", err)
	}
	r := newByteReader(data)
	r.pos = formEnd
	level := &Level{Record: record}
	if record.Mask&formMaskMouseIcon != 0 {
		level.MouseIconRaw, err = readGUIDAndPicture(r)
		if err != nil {
			return nil, parseError(path, "f", r.pos, "FormStreamData.MouseIcon", "%v", err)
		}
	}
	if record.Mask&formMaskFont != 0 {
		level.FontRaw, err = readFontBlob(r)
		if err != nil {
			return nil, parseError(path, "f", r.pos, "FormStreamData.Font", "%v", err)
		}
	}
	if record.Mask&formMaskPicture != 0 {
		level.PictureRaw, err = readGUIDAndPicture(r)
		if err != nil {
			return nil, parseError(path, "f", r.pos, "FormStreamData.Picture", "%v", err)
		}
	}
	start := r.pos
	parsed, errStored := parseSiteData(data, start, codePage, true)
	if errStored != nil {
		parsed, err = parseSiteData(data, start, codePage, false)
		if err != nil {
			return nil, parseError(path, "f", start, "FormSiteData", "stored class-table layout: %v; omitted layout: %v", errStored, err)
		}
	}
	level.Sites = parsed.sites
	level.ClassTable = parsed.classTable
	level.ClassTableRaw = parsed.classTableRaw
	level.DepthsRaw = parsed.depthsRaw
	level.TrailingRaw = parsed.trailingRaw
	return level, nil
}

type parsedSiteData struct {
	sites         []*Site
	classTable    []ClassInfo
	classTableRaw []byte
	depthsRaw     []byte
	trailingRaw   []byte
}

func parseSiteData(data []byte, start int, codePage uint16, classTableStored bool) (parsedSiteData, error) {
	r := newByteReader(data)
	r.pos = start
	classStart := start
	var classTable []ClassInfo
	if classTableStored {
		count, err := r.uint16()
		if err != nil {
			return parsedSiteData{}, err
		}
		if int(count) > r.remaining()/4 {
			return parsedSiteData{}, fmt.Errorf("class count %d exceeds remaining bytes", count)
		}
		classTable = make([]ClassInfo, 0, count)
		for range int(count) {
			recordStart := r.pos
			version, err := r.uint16()
			if err != nil {
				return parsedSiteData{}, err
			}
			if version != 0 {
				return parsedSiteData{}, fmt.Errorf("unsupported SiteClassInfo version %d", version)
			}
			cb, err := r.uint16()
			if err != nil {
				return parsedSiteData{}, err
			}
			if _, err := r.take(int(cb)); err != nil {
				return parsedSiteData{}, err
			}
			classTable = append(classTable, ClassInfo{Version: version, Raw: bytes.Clone(data[recordStart:r.pos])})
		}
	}
	classRaw := bytes.Clone(data[classStart:r.pos])
	count, err := r.uint32()
	if err != nil {
		return parsedSiteData{}, err
	}
	if count > maxSitesPerForm {
		return parsedSiteData{}, fmt.Errorf("site count %d exceeds limit %d", count, maxSitesPerForm)
	}
	countOfBytes, err := r.uint32()
	if err != nil {
		return parsedSiteData{}, err
	}
	if uint64(countOfBytes) > uint64(r.remaining()) {
		return parsedSiteData{}, fmt.Errorf("CountOfBytes %d exceeds remaining %d", countOfBytes, r.remaining())
	}
	end := r.pos + int(countOfBytes)
	if end != len(data) && !isTrailingRecord(data, end) {
		return parsedSiteData{}, fmt.Errorf("CountOfBytes ends at %d, expected %d or one exact trailing record", end, len(data))
	}
	depthStart := r.pos
	depths := make([]depthType, 0, count)
	for len(depths) < int(count) {
		depth, err := r.uint8()
		if err != nil {
			return parsedSiteData{}, err
		}
		typeOrCount, err := r.uint8()
		if err != nil {
			return parsedSiteData{}, err
		}
		run := 1
		typ := typeOrCount
		if typeOrCount&0x80 != 0 {
			run = int(typeOrCount & 0x7f)
			if run == 0 {
				return parsedSiteData{}, fmt.Errorf("zero-length SiteDepthsAndTypes run")
			}
			typ, err = r.uint8()
			if err != nil {
				return parsedSiteData{}, err
			}
		}
		if len(depths)+run > int(count) {
			return parsedSiteData{}, fmt.Errorf("SiteDepthsAndTypes accounts for more than %d sites", count)
		}
		for range run {
			depths = append(depths, depthType{depth: depth, typ: typ})
		}
	}
	if _, err := r.align(depthStart, 4); err != nil {
		return parsedSiteData{}, err
	}
	depthsRaw := bytes.Clone(data[depthStart:r.pos])
	sites := make([]*Site, 0, count)
	for i := range int(count) {
		site, err := parseSite(r, codePage)
		if err != nil {
			return parsedSiteData{}, fmt.Errorf("site %d: %w", i, err)
		}
		site.Values["_Depth"] = int64(depths[i].depth)
		site.Values["_Type"] = int64(depths[i].typ)
		sites = append(sites, site)
	}
	if r.pos != end {
		return parsedSiteData{}, fmt.Errorf("sites end at %d, CountOfBytes boundary is %d", r.pos, end)
	}
	return parsedSiteData{
		sites: sites, classTable: classTable, classTableRaw: classRaw,
		depthsRaw: depthsRaw, trailingRaw: bytes.Clone(data[end:]),
	}, nil
}

func parseSite(r *byteReader, codePage uint16) (*Site, error) {
	start := r.pos
	version, err := r.uint16()
	if err != nil {
		return nil, err
	}
	if version != 0 {
		return nil, fmt.Errorf("unsupported OleSiteConcreteControl version %d", version)
	}
	cbSite, err := r.uint16()
	if err != nil {
		return nil, err
	}
	end := start + 4 + int(cbSite)
	if end > r.end {
		return nil, fmt.Errorf("cbSite %d exceeds enclosing boundary", cbSite)
	}
	mask, err := r.uint32()
	if err != nil {
		return nil, err
	}
	site := &Site{
		Version: version, Mask: mask, Values: map[string]int64{},
		Strings: map[string]StoredString{}, Padding: map[string][]byte{},
	}
	for _, field := range siteFields {
		if mask&(uint32(1)<<field.bit) == 0 {
			continue
		}
		pad, err := r.align(start, field.size)
		if err != nil {
			return nil, err
		}
		if len(pad) > 0 {
			site.Padding["before:"+field.name] = bytes.Clone(pad)
		}
		value, err := readInteger(r, field.size, field.signed)
		if err != nil {
			return nil, err
		}
		site.Values[field.name] = value
	}
	pad, err := r.align(start, 4)
	if err != nil {
		return nil, err
	}
	if len(pad) > 0 {
		site.Padding["data:end"] = bytes.Clone(pad)
	}
	readString := func(lengthField, name string) error {
		value, found := site.Values[lengthField]
		if !found {
			return nil
		}
		stored, err := readStoredString(r, uint32(value), codePage)
		if err != nil {
			return err
		}
		site.Strings[name] = stored
		padding, err := r.align(start, 4)
		if err != nil {
			return err
		}
		if len(padding) > 0 {
			site.Padding["str:"+name] = bytes.Clone(padding)
		}
		return nil
	}
	for _, item := range siteStringsBeforePosition {
		if err := readString(item[0], item[1]); err != nil {
			return nil, err
		}
	}
	if mask&(1<<8) != 0 {
		left, err := r.int32()
		if err != nil {
			return nil, err
		}
		top, err := r.int32()
		if err != nil {
			return nil, err
		}
		site.Position = &Position{Left: left, Top: top}
	}
	for _, item := range siteStringsAfterPosition {
		if err := readString(item[0], item[1]); err != nil {
			return nil, err
		}
	}
	if r.pos != end {
		return nil, fmt.Errorf("site %q consumed to %d, cbSite boundary is %d", siteName(site), r.pos, end)
	}
	site.Raw = bytes.Clone(r.data[start:end])
	return site, nil
}

type siteField struct {
	bit    uint8
	name   string
	size   int
	signed bool
}

var siteFields = []siteField{
	{0, "NameData", 4, false}, {1, "TagData", 4, false}, {2, "ID", 4, true},
	{3, "HelpContextID", 4, true}, {4, "BitFlags", 4, false},
	{5, "ObjectStreamSize", 4, false}, {6, "TabIndex", 2, true},
	{7, "ClsidCacheIndex", 2, false}, {9, "GroupID", 2, false},
	{11, "ControlTipTextData", 4, false}, {12, "RuntimeLicKeyData", 4, false},
	{13, "ControlSourceData", 4, false}, {14, "RowSourceData", 4, false},
}

var siteStringsBeforePosition = [][2]string{{"NameData", "Name"}, {"TagData", "Tag"}}
var siteStringsAfterPosition = [][2]string{
	{"ControlTipTextData", "ControlTipText"}, {"RuntimeLicKeyData", "RuntimeLicKey"},
	{"ControlSourceData", "ControlSource"}, {"RowSourceData", "RowSource"},
}

func readFontBlob(r *byteReader) ([]byte, error) {
	start := r.pos
	guid, err := r.take(16)
	if err != nil {
		return nil, err
	}
	switch binary.LittleEndian.Uint32(guid[:4]) {
	case guidStdFont:
		version, err := r.uint8()
		if err != nil {
			return nil, err
		}
		if version != 1 {
			return nil, fmt.Errorf("unsupported StdFont version %d", version)
		}
		if _, err := r.take(9); err != nil {
			return nil, err
		}
		nameLength, err := r.uint8()
		if err != nil {
			return nil, err
		}
		if _, err := r.take(int(nameLength)); err != nil {
			return nil, err
		}
	case guidTextProps:
		if _, err := r.take(2); err != nil {
			return nil, err
		}
		cb, err := r.uint16()
		if err != nil {
			return nil, err
		}
		if _, err := r.take(int(cb)); err != nil {
			return nil, err
		}
	default:
		return nil, fmt.Errorf("unknown font GUID %x", guid)
	}
	return bytes.Clone(r.data[start:r.pos]), nil
}

func isTrailingRecord(data []byte, offset int) bool {
	if offset < 0 || offset+4 > len(data) || data[offset] != 0 || data[offset+1] != 2 {
		return false
	}
	return offset+4+int(binary.LittleEndian.Uint16(data[offset+2:offset+4])) == len(data)
}

func parseCompObj(data []byte, path string, codePage uint16) (CompObj, error) {
	if len(data) > maxDesignerStreamSize {
		return CompObj{}, parseError(path, "\x01CompObj", 0, "CompObj", "stream exceeds %d-byte limit", maxDesignerStreamSize)
	}
	if len(data) < 28 {
		return CompObj{}, parseError(path, "\x01CompObj", 0, "CompObj", "stream is shorter than 28-byte header")
	}
	r := newByteReader(data)
	header, _ := r.take(28)
	byteOrderVersion := binary.LittleEndian.Uint32(header[0:4])
	formatVersion := binary.LittleEndian.Uint32(header[4:8])
	if byteOrderVersion != 0xfffe0001 || formatVersion != 0x00000a03 {
		return CompObj{}, parseError(path, "\x01CompObj", 0, "CompObj", "unsupported header %08x/%08x", byteOrderVersion, formatVersion)
	}
	readLPString := func(name string) (StoredString, error) {
		length, err := r.uint32()
		if err != nil {
			return StoredString{}, fmt.Errorf("%s length: %w", name, err)
		}
		if uint64(length) > uint64(r.remaining()) {
			return StoredString{}, fmt.Errorf("%s length %d exceeds remaining %d", name, length, r.remaining())
		}
		raw, err := r.take(int(length))
		if err != nil {
			return StoredString{}, err
		}
		textRaw := bytes.TrimSuffix(raw, []byte{0})
		text, err := ovba.DecodeMBCS(textRaw, codePage)
		if err != nil {
			return StoredString{}, fmt.Errorf("%s: %w", name, err)
		}
		return StoredString{Text: text, Compressed: true, Raw: bytes.Clone(raw)}, nil
	}
	userType, err := readLPString("UserType")
	if err != nil {
		return CompObj{}, parseError(path, "\x01CompObj", r.pos, "CompObj", "%v", err)
	}
	clipboard, err := readLPString("ClipboardFormat")
	if err != nil {
		return CompObj{}, parseError(path, "\x01CompObj", r.pos, "CompObj", "%v", err)
	}
	return CompObj{
		ByteOrderVersion: byteOrderVersion, FormatVersion: formatVersion,
		Header: bytes.Clone(header), UserType: userType, ClipboardFormat: clipboard,
		TailRaw: bytes.Clone(data[r.pos:]), Raw: bytes.Clone(data),
	}, nil
}

func siteName(site *Site) string {
	if value, ok := site.Strings["Name"]; ok {
		return value.Text
	}
	return ""
}

func controlKind(index uint16, record *Record) string {
	if index >= 0x8000 {
		return "ActiveX.Control"
	}
	kind, ok := controlKinds[index]
	if !ok {
		return "MSForms.Control"
	}
	if index == 15 && record != nil {
		switch record.Values["DisplayStyle"] {
		case 0, 1:
			return "MSForms.TextBox"
		case 2:
			return "MSForms.ListBox"
		case 3, 7:
			return "MSForms.ComboBox"
		case 4:
			return "MSForms.CheckBox"
		case 5:
			return "MSForms.OptionButton"
		case 6:
			return "MSForms.ToggleButton"
		default:
			return "MSForms.Control"
		}
	}
	return kind
}

func pathBase(path string) string {
	if i := strings.LastIndexByte(path, '/'); i >= 0 {
		return path[i+1:]
	}
	return path
}
