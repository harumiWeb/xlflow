package cfb

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"slices"
	"sort"
	"strings"
	"unicode"
	"unicode/utf16"
)

// Writer assembles the streams registered via AddStream into a single CFB file.
type Writer struct {
	streams  []streamSpec
	storages []storageSpec
	format   Format
}

type streamSpec struct {
	path []string
	data []byte
}

type storageSpec struct {
	path []string
	meta StorageMeta
}

// NewWriter returns an empty v3 Writer.
func NewWriter() *Writer { return &Writer{format: FormatV3} }

// NewWriterForFormat returns an empty Writer for the requested CFB version.
func NewWriterForFormat(format Format) (*Writer, error) {
	if _, err := geometryFor(format); err != nil {
		return nil, err
	}
	return &Writer{format: format}, nil
}

// AddStream registers a stream. The last path element becomes the stream name
// and the intermediate elements become storages. Multiple streams that share
// an intermediate storage produce only one storage. The arguments are copied,
// so the caller may reuse them.
func (w *Writer) AddStream(path []string, data []byte) {
	w.streams = append(w.streams, streamSpec{slices.Clone(path), bytes.Clone(data)})
}

// AddStorage registers a storage and its directory metadata. An empty path
// identifies the root directory entry. Repeating an identical definition is
// allowed; conflicting definitions are rejected when Bytes is called. The
// path is copied, so the caller may reuse it.
func (w *Writer) AddStorage(path []string, meta StorageMeta) {
	w.storages = append(w.storages, storageSpec{slices.Clone(path), meta})
}

// node is one element of the directory tree (root / storage / stream).
type node struct {
	name     string
	objType  byte
	data     []byte           // stream only
	children map[string]*node // storage / root only
	meta     StorageMeta      // storage / root only
	metaSet  bool

	id          uint32 // index in the directory array
	left, right uint32 // sibling red-black tree links (all black)
	child       uint32 // root of the child-element tree
	startSector uint32
	size        uint64
}

func newNode(name string, objType byte) *node {
	return &node{
		name:     name,
		objType:  objType,
		children: map[string]*node{},
		left:     noStream,
		right:    noStream,
		child:    noStream,
	}
}

// Bytes assembles the entire CFB file in memory and returns it.
func (w *Writer) Bytes() ([]byte, error) {
	root, all, err := w.buildTree()
	if err != nil {
		return nil, err
	}
	format := w.format
	if format == 0 {
		format = FormatV3
	}
	return assemble(root, all, format)
}

// buildTree constructs the directory tree from the AddStream calls, assigns ids,
// and links the sibling red-black tree. all[0] is always the root.
func (w *Writer) buildTree() (*node, []*node, error) {
	root := newNode("Root Entry", objRoot)

	for _, s := range w.storages {
		if err := validatePath(s.path, true); err != nil {
			return nil, nil, err
		}
		if len(s.path) == 0 {
			if s.meta.Created != 0 {
				return nil, nil, fmt.Errorf("cfb: root storage has a nonzero creation FILETIME")
			}
			if root.metaSet && root.meta != s.meta {
				return nil, nil, fmt.Errorf("cfb: conflicting metadata for root storage")
			}
			root.meta = s.meta
			root.metaSet = true
			continue
		}
		cur := root
		for i, part := range s.path {
			ch, ok := cur.children[part]
			if !ok {
				ch = newNode(part, objStorage)
				cur.children[part] = ch
			} else if ch.objType != objStorage {
				return nil, nil, fmt.Errorf("cfb: %q already exists as a stream but was registered as a storage", part)
			}
			if i == len(s.path)-1 {
				if ch.metaSet && ch.meta != s.meta {
					return nil, nil, fmt.Errorf("cfb: conflicting metadata for storage %q", strings.Join(s.path, "/"))
				}
				ch.meta = s.meta
				ch.metaSet = true
			}
			cur = ch
		}
	}

	for _, s := range w.streams {
		if err := validatePath(s.path, false); err != nil {
			return nil, nil, err
		}
		cur := root
		for i, part := range s.path {
			last := i == len(s.path)-1
			ch, ok := cur.children[part]
			if !ok {
				if last {
					ch = newNode(part, objStream)
					ch.data = s.data
				} else {
					ch = newNode(part, objStorage)
				}
				cur.children[part] = ch
			} else {
				// Detect a type conflict with an existing element.
				if last && ch.objType != objStream {
					return nil, nil, fmt.Errorf("cfb: %q already exists as a storage but was registered as a stream", part)
				}
				if !last && ch.objType != objStorage {
					return nil, nil, fmt.Errorf("cfb: %q already exists as a stream but was used as a storage", part)
				}
				if last {
					return nil, nil, fmt.Errorf("cfb: duplicate stream %q", part)
				}
			}
			cur = ch
		}
	}

	// id assignment: root=0, then each level in child-name order.
	all := []*node{root}
	var assign func(n *node)
	assign = func(n *node) {
		kids := sortedChildren(n)
		for _, k := range kids {
			k.id = uint32(len(all))
			all = append(all, k)
		}
		for _, k := range kids {
			assign(k)
		}
	}
	assign(root)

	// Sibling red-black tree links (all black, balanced BST).
	var link func(n *node)
	link = func(n *node) {
		kids := sortedChildren(n)
		n.child = buildBST(kids)
		for _, k := range kids {
			link(k)
		}
	}
	link(root)

	return root, all, nil
}

func validatePath(path []string, allowRoot bool) error {
	if len(path) == 0 && !allowRoot {
		return fmt.Errorf("cfb: empty stream path is not allowed")
	}
	for _, part := range path {
		if part == "" {
			return fmt.Errorf("cfb: directory name must not be empty")
		}
		units := utf16.Encode([]rune(part))
		if len(units) > 31 {
			return fmt.Errorf("cfb: name %q is too long (must fit in 31 UTF-16 code units)", part)
		}
		for _, unit := range units {
			if forbiddenDirectoryNameUnit(unit) {
				return fmt.Errorf("cfb: name %q contains forbidden character %q", part, rune(unit))
			}
		}
	}
	return nil
}

// sortedChildren returns the child elements sorted by the [MS-CFB] §2.6.4 name order.
func sortedChildren(n *node) []*node {
	kids := make([]*node, 0, len(n.children))
	for _, k := range n.children {
		kids = append(kids, k)
	}
	sort.Slice(kids, func(i, j int) bool {
		return nameLess(kids[i].name, kids[j].name)
	})
	return kids
}

// buildBST builds a balanced BST from a name-sorted array and returns the root's id.
// The middle becomes the root; the left half -> left, the right half -> right. Empty yields NOSTREAM.
func buildBST(kids []*node) uint32 {
	if len(kids) == 0 {
		return noStream
	}
	mid := len(kids) / 2
	k := kids[mid]
	k.left = buildBST(kids[:mid])
	k.right = buildBST(kids[mid+1:])
	return k.id
}

// nameLess implements [MS-CFB] §2.6.4: a shorter UTF-16 code-unit length sorts first;
// if equal length, upcase and compare each code unit.
func nameLess(a, b string) bool {
	ua := utf16.Encode([]rune(a))
	ub := utf16.Encode([]rune(b))
	if len(ua) != len(ub) {
		return len(ua) < len(ub)
	}
	for i := range ua {
		ca := upcase(ua[i])
		cb := upcase(ub[i])
		if ca != cb {
			return ca < cb
		}
	}
	return false
}

func upcase(u uint16) uint16 {
	return uint16(unicode.ToUpper(rune(u)))
}

// DirectoryNameKey returns the case-insensitive identity key a CFB directory
// applies to an entry name: its UTF-16 code units, each upcased per
// [MS-CFB] \u00a72.6.4. Two names are equivalent in the container (and
// therefore collide) exactly when their keys are equal.
func DirectoryNameKey(name string) string {
	units := utf16.Encode([]rune(name))
	key := make([]byte, 0, len(units)*2)
	for _, u := range units {
		u = upcase(u)
		key = append(key, byte(u), byte(u>>8))
	}
	return string(key)
}

// assemble lays out the sectors in two passes and builds the CFB byte stream.
func assemble(root *node, all []*node, format Format) ([]byte, error) {
	g, err := geometryFor(format)
	if err != nil {
		return nil, err
	}
	// --- Pass 1: decide the layout ---
	// regular streams (>=cutoff) -> mini stream container -> mini FAT -> directory -> FAT.
	var large, mini []*node
	for _, n := range all[1:] {
		if n.objType != objStream {
			continue
		}
		if len(n.data) >= cutoff {
			large = append(large, n)
		} else {
			mini = append(mini, n)
		}
	}

	sector := uint64(0)
	reserve := func(count uint64, label string) (uint32, error) {
		if count > uint64(^uint32(0))-sector || sector > uint64(^uint32(0)) {
			return 0, fmt.Errorf("cfb: %s exceeds the sector-number limit", label)
		}
		start := uint32(sector)
		sector += count
		return start, nil
	}

	// Regular streams: stored directly in sectors.
	for _, n := range large {
		count := ceilDiv64(uint64(len(n.data)), uint64(g.sectorSize))
		start, err := reserve(count, "regular stream")
		if err != nil {
			return nil, err
		}
		n.startSector = start
		n.size = uint64(len(n.data))
	}

	// Build the mini stream container and mini FAT.
	var miniStream []byte
	var miniFat []uint32
	for _, n := range mini {
		n.size = uint64(len(n.data))
		if len(n.data) == 0 {
			n.startSector = endOfChain
			continue
		}
		nMini64 := ceilDiv64(uint64(len(n.data)), miniSectorSize)
		if uint64(len(miniFat)) > uint64(^uint32(0)) || nMini64 > uint64(^uint32(0))-uint64(len(miniFat)) {
			return nil, fmt.Errorf("cfb: mini stream exceeds the mini-sector-number limit")
		}
		nMini, err := checkedInt(nMini64, "mini stream sector count")
		if err != nil {
			return nil, err
		}
		chunkSize, err := checkedInt(nMini64*miniSectorSize, "mini stream padded size")
		if err != nil {
			return nil, err
		}
		start := uint32(len(miniFat)) // mini sector number
		n.startSector = start
		chunk := make([]byte, chunkSize) // zero-pad to the 64B boundary
		copy(chunk, n.data)
		miniStream = append(miniStream, chunk...)
		for j := 0; j < nMini; j++ {
			if j == nMini-1 {
				miniFat = append(miniFat, endOfChain)
			} else {
				miniFat = append(miniFat, start+uint32(j)+1)
			}
		}
	}

	// Mini stream container (regular sectors, 512B boundary).
	containerStart := uint32(endOfChain)
	var containerSectors uint64
	if len(miniStream) > 0 {
		containerSectors = ceilDiv64(uint64(len(miniStream)), uint64(g.sectorSize))
		var err error
		containerStart, err = reserve(containerSectors, "mini stream container")
		if err != nil {
			return nil, err
		}
	}
	root.startSector = containerStart
	root.size = uint64(len(miniStream))

	// Mini FAT (regular sectors).
	miniFatStart := uint32(endOfChain)
	var miniFatSectors uint64
	if len(miniFat) > 0 {
		miniFatSectors = ceilDiv64(uint64(len(miniFat))*4, uint64(g.sectorSize))
		var err error
		miniFatStart, err = reserve(miniFatSectors, "mini-FAT")
		if err != nil {
			return nil, err
		}
	}

	// Directory.
	dirSectors := ceilDiv64(uint64(len(all)), uint64(g.entriesPerDirSector))
	dirStart, err := reserve(dirSectors, "directory")
	if err != nil {
		return nil, err
	}

	// --- Converge on FAT and DIFAT sector counts. ---
	nonAllocation := sector
	fatSectors, difatSectors := uint64(1), uint64(0)
	for {
		total := nonAllocation + fatSectors + difatSectors
		nextFat := ceilDiv64(total, uint64(g.entriesPerFatSector))
		nextDifat := uint64(0)
		if nextFat > difatHeaderLen {
			nextDifat = ceilDiv64(nextFat-difatHeaderLen, uint64(g.entriesPerDifatSector))
		}
		if nextFat == fatSectors && nextDifat == difatSectors {
			break
		}
		fatSectors, difatSectors = nextFat, nextDifat
	}
	fatStart, err := reserve(fatSectors, "FAT")
	if err != nil {
		return nil, err
	}
	difatStart := uint32(endOfChain)
	if difatSectors > 0 {
		difatStart, err = reserve(difatSectors, "DIFAT")
		if err != nil {
			return nil, err
		}
	}
	totalSectors := sector

	// --- Pass 2: fill the FAT array ---
	fatEntries, err := checkedInt(fatSectors*uint64(g.entriesPerFatSector), "writer FAT entry count")
	if err != nil {
		return nil, err
	}
	fat := make([]uint32, fatEntries)
	for i := range fat {
		fat[i] = freeSect
	}
	for _, n := range large {
		chainRun(fat, n.startSector, uint32(ceilDiv64(uint64(len(n.data)), uint64(g.sectorSize))))
	}
	if containerSectors > 0 {
		chainRun(fat, containerStart, uint32(containerSectors))
	}
	if miniFatSectors > 0 {
		chainRun(fat, miniFatStart, uint32(miniFatSectors))
	}
	chainRun(fat, dirStart, uint32(dirSectors))
	for i := uint32(0); i < uint32(fatSectors); i++ {
		fat[fatStart+i] = fatSect
	}
	for i := uint32(0); i < uint32(difatSectors); i++ {
		fat[difatStart+i] = difSect
	}

	// Pad the mini FAT array with FREESECT up to the 512B (128-entry) boundary.
	miniFatPadded := make([]uint32, int(miniFatSectors)*g.entriesPerFatSector)
	for i := range miniFatPadded {
		miniFatPadded[i] = freeSect
	}
	copy(miniFatPadded, miniFat)

	// --- Assemble the byte stream ---
	outputSize, err := checkedInt(uint64(g.headerSpan)+totalSectors*uint64(g.sectorSize), "output size")
	if err != nil {
		return nil, err
	}
	buf := make([]byte, outputSize)
	off := func(s uint32) int { return g.headerSpan + int(s)*g.sectorSize }

	// Header.
	writeHeader(buf, g, uint32(fatSectors), dirStart, uint32(dirSectors), miniFatStart, uint32(miniFatSectors), fatStart, difatStart, uint32(difatSectors))

	// Regular stream bodies.
	for _, n := range large {
		copy(buf[off(n.startSector):], n.data)
	}
	// Mini stream container.
	if containerSectors > 0 {
		copy(buf[off(containerStart):], miniStream)
	}
	// Mini FAT.
	for i, v := range miniFatPadded {
		binary.LittleEndian.PutUint32(buf[off(miniFatStart)+i*4:], v)
	}
	// Directory.
	writeDirectory(buf[off(dirStart):], all, uint32(dirSectors), g.entriesPerDirSector)
	// FAT.
	for i, v := range fat {
		binary.LittleEndian.PutUint32(buf[off(fatStart)+i*4:], v)
	}
	// DIFAT sectors contain FAT sector numbers beyond the 109 header entries.
	remainingFat := int(fatSectors) - difatHeaderLen
	for i := range int(difatSectors) {
		base := off(difatStart + uint32(i))
		for j := range g.entriesPerDifatSector {
			v := uint32(freeSect)
			index := i*g.entriesPerDifatSector + j
			if index < remainingFat {
				v = fatStart + uint32(difatHeaderLen+index)
			}
			binary.LittleEndian.PutUint32(buf[base+j*4:], v)
		}
		next := uint32(endOfChain)
		if i+1 < int(difatSectors) {
			next = difatStart + uint32(i+1)
		}
		binary.LittleEndian.PutUint32(buf[base+g.entriesPerDifatSector*4:], next)
	}

	return buf, nil
}

// chainRun chains count consecutive sectors in the FAT (ending with ENDOFCHAIN).
func chainRun(fat []uint32, start, count uint32) {
	for i := uint32(0); i < count; i++ {
		if i == count-1 {
			fat[start+i] = endOfChain
		} else {
			fat[start+i] = start + i + 1
		}
	}
}

// writeHeader writes the 512B header at the start of buf.
func writeHeader(buf []byte, g geometry, fatSectors, dirStart, dirSectors, miniFatStart, miniFatSectors, fatStart, difatStart, difatSectors uint32) {
	copy(buf[0:8], []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	// CLSID(8..24) zero
	binary.LittleEndian.PutUint16(buf[24:], minorVersion)
	binary.LittleEndian.PutUint16(buf[26:], uint16(g.format))
	binary.LittleEndian.PutUint16(buf[28:], byteOrderMark)
	binary.LittleEndian.PutUint16(buf[30:], g.sectorShift)
	binary.LittleEndian.PutUint16(buf[32:], miniSectorShift)
	// reserved(34..40) zero
	if g.format == FormatV4 {
		binary.LittleEndian.PutUint32(buf[40:], dirSectors)
	}
	binary.LittleEndian.PutUint32(buf[44:], fatSectors)
	binary.LittleEndian.PutUint32(buf[48:], dirStart)
	binary.LittleEndian.PutUint32(buf[52:], 0) // transactionSignature
	binary.LittleEndian.PutUint32(buf[56:], cutoff)
	binary.LittleEndian.PutUint32(buf[60:], miniFatStart)
	binary.LittleEndian.PutUint32(buf[64:], miniFatSectors)
	binary.LittleEndian.PutUint32(buf[68:], difatStart)
	binary.LittleEndian.PutUint32(buf[72:], difatSectors)

	// DIFAT array (76..512, 109 entries). FAT sector numbers first, the rest FREESECT.
	for i := 0; i < difatHeaderLen; i++ {
		v := uint32(freeSect)
		if uint32(i) < fatSectors {
			v = fatStart + uint32(i)
		}
		binary.LittleEndian.PutUint32(buf[76+i*4:], v)
	}
}

// writeDirectory writes the 128B entry array into dst (the start of the directory region).
// Leftover entries are filled as unused (objType=0, sibling/child=NOSTREAM).
func writeDirectory(dst []byte, all []*node, dirSectors uint32, entriesPerSector int) {
	totalEntries := int(dirSectors) * entriesPerSector
	for i := 0; i < totalEntries; i++ {
		e := dst[i*dirEntrySize : (i+1)*dirEntrySize]
		if i < len(all) {
			writeDirEntry(e, all[i])
		} else {
			// Unused entry.
			binary.LittleEndian.PutUint32(e[68:], noStream)
			binary.LittleEndian.PutUint32(e[72:], noStream)
			binary.LittleEndian.PutUint32(e[76:], noStream)
		}
	}
}

// writeDirEntry writes one 128B DirectoryEntry.
func writeDirEntry(e []byte, n *node) {
	// name: UTF-16LE + NUL terminator (64B field).
	u16 := utf16.Encode([]rune(n.name))
	for i, cu := range u16 {
		binary.LittleEndian.PutUint16(e[i*2:], cu)
	}
	nameLen := (len(u16) + 1) * 2 // byte length incl. NUL
	binary.LittleEndian.PutUint16(e[64:], uint16(nameLen))
	e[66] = n.objType
	e[67] = 1 // colorFlag: all black
	binary.LittleEndian.PutUint32(e[68:], n.left)
	binary.LittleEndian.PutUint32(e[72:], n.right)
	binary.LittleEndian.PutUint32(e[76:], n.child)
	copy(e[80:96], n.meta.CLSID[:])
	binary.LittleEndian.PutUint32(e[96:], n.meta.StateBits)
	binary.LittleEndian.PutUint64(e[100:], n.meta.Created)
	binary.LittleEndian.PutUint64(e[108:], n.meta.Modified)
	binary.LittleEndian.PutUint32(e[116:], n.startSector)
	binary.LittleEndian.PutUint64(e[120:], uint64(n.size)) // upper 4 bytes are 0 (v3)
}
