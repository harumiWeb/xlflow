package cfb

import (
	"encoding/binary"
	"fmt"
	"sort"
	"unicode"
	"unicode/utf16"
)

// Writer assembles the streams registered via AddStream into a single CFB file.
type Writer struct {
	streams []streamSpec
}

type streamSpec struct {
	path []string
	data []byte
}

// NewWriter returns an empty Writer.
func NewWriter() *Writer { return &Writer{} }

// AddStream registers a stream. The last path element becomes the stream name
// and the intermediate elements become storages. Multiple streams that share
// an intermediate storage produce only one storage. The arguments are copied,
// so the caller may reuse them.
func (w *Writer) AddStream(path []string, data []byte) {
	cp := make([]string, len(path))
	copy(cp, path)
	cd := make([]byte, len(data))
	copy(cd, data)
	w.streams = append(w.streams, streamSpec{cp, cd})
}

// node is one element of the directory tree (root / storage / stream).
type node struct {
	name     string
	objType  byte
	data     []byte           // stream only
	children map[string]*node // storage / root only

	id          uint32 // index in the directory array
	left, right uint32 // sibling red-black tree links (all black)
	child       uint32 // root of the child-element tree
	startSector uint32
	size        uint32
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
	return assemble(root, all)
}

// buildTree constructs the directory tree from the AddStream calls, assigns ids,
// and links the sibling red-black tree. all[0] is always the root.
func (w *Writer) buildTree() (*node, []*node, error) {
	root := newNode("Root Entry", objRoot)

	for _, s := range w.streams {
		if len(s.path) == 0 {
			return nil, nil, fmt.Errorf("cfb: empty stream path is not allowed")
		}
		for _, part := range s.path {
			// The name must fit in the 64B field (UTF-16 + NUL terminator).
			if len(utf16.Encode([]rune(part))) > 31 {
				return nil, nil, fmt.Errorf("cfb: name %q is too long (must fit in 31 UTF-16 code units)", part)
			}
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

func ceilDiv(a, b int) int { return (a + b - 1) / b }

// assemble lays out the sectors in two passes and builds the CFB byte stream.
func assemble(root *node, all []*node) ([]byte, error) {
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

	sector := uint32(0)

	// Regular streams: stored directly in 512B sectors.
	for _, n := range large {
		n.startSector = sector
		n.size = uint32(len(n.data))
		sector += uint32(ceilDiv(len(n.data), sectorSize))
	}

	// Build the mini stream container and mini FAT.
	var miniStream []byte
	var miniFat []uint32
	for _, n := range mini {
		n.size = uint32(len(n.data))
		if len(n.data) == 0 {
			n.startSector = endOfChain
			continue
		}
		start := uint32(len(miniFat)) // mini sector number
		n.startSector = start
		nMini := ceilDiv(len(n.data), miniSectorSize)
		chunk := make([]byte, nMini*miniSectorSize) // zero-pad to the 64B boundary
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
	var containerSectors uint32
	if len(miniStream) > 0 {
		containerStart = sector
		containerSectors = uint32(ceilDiv(len(miniStream), sectorSize))
		sector += containerSectors
	}
	root.startSector = containerStart
	root.size = uint32(len(miniStream))

	// Mini FAT (regular sectors).
	miniFatStart := uint32(endOfChain)
	var miniFatSectors uint32
	if len(miniFat) > 0 {
		miniFatStart = sector
		miniFatSectors = uint32(ceilDiv(len(miniFat)*4, sectorSize))
		sector += miniFatSectors
	}

	// Directory (regular sectors, 4 entries/sector).
	dirSectors := uint32(ceilDiv(len(all), entriesPerDirSector))
	dirStart := sector
	sector += dirSectors

	// --- Converge on the number of FAT sectors ---
	// The minimum number of sectors (at 128 entries/sector) able to describe the total sector count, FAT included.
	nonFat := sector
	fatSectors := uint32(1)
	for uint32(entriesPerFatSector)*fatSectors < nonFat+fatSectors {
		fatSectors++
	}
	if fatSectors > difatHeaderLen {
		return nil, fmt.Errorf("cfb: FAT needs %d sectors, which exceeds the 109 DIFAT header entries (outside the minimal profile)", fatSectors)
	}
	fatStart := sector
	sector += fatSectors
	totalSectors := sector

	// --- Pass 2: fill the FAT array ---
	fat := make([]uint32, fatSectors*uint32(entriesPerFatSector))
	for i := range fat {
		fat[i] = freeSect
	}
	for _, n := range large {
		chainRun(fat, n.startSector, uint32(ceilDiv(len(n.data), sectorSize)))
	}
	if containerSectors > 0 {
		chainRun(fat, containerStart, containerSectors)
	}
	if miniFatSectors > 0 {
		chainRun(fat, miniFatStart, miniFatSectors)
	}
	chainRun(fat, dirStart, dirSectors)
	for i := uint32(0); i < fatSectors; i++ {
		fat[fatStart+i] = fatSect
	}

	// Pad the mini FAT array with FREESECT up to the 512B (128-entry) boundary.
	miniFatPadded := make([]uint32, miniFatSectors*uint32(entriesPerFatSector))
	for i := range miniFatPadded {
		miniFatPadded[i] = freeSect
	}
	copy(miniFatPadded, miniFat)

	// --- Assemble the byte stream ---
	buf := make([]byte, headerSize+int(totalSectors)*sectorSize)
	off := func(s uint32) int { return headerSize + int(s)*sectorSize }

	// Header.
	writeHeader(buf, fatSectors, dirStart, miniFatStart, miniFatSectors, fatStart)

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
	writeDirectory(buf[off(dirStart):], all, dirSectors)
	// FAT.
	for i, v := range fat {
		binary.LittleEndian.PutUint32(buf[off(fatStart)+i*4:], v)
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
func writeHeader(buf []byte, fatSectors, dirStart, miniFatStart, miniFatSectors, fatStart uint32) {
	copy(buf[0:8], []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	// CLSID(8..24) zero
	binary.LittleEndian.PutUint16(buf[24:], minorVersion)
	binary.LittleEndian.PutUint16(buf[26:], majorVersion)
	binary.LittleEndian.PutUint16(buf[28:], byteOrderMark)
	binary.LittleEndian.PutUint16(buf[30:], sectorShift)
	binary.LittleEndian.PutUint16(buf[32:], miniSectorShift)
	// reserved(34..40) zero
	binary.LittleEndian.PutUint32(buf[40:], 0) // numDirSectors (0 in v3)
	binary.LittleEndian.PutUint32(buf[44:], fatSectors)
	binary.LittleEndian.PutUint32(buf[48:], dirStart)
	binary.LittleEndian.PutUint32(buf[52:], 0) // transactionSignature
	binary.LittleEndian.PutUint32(buf[56:], cutoff)
	binary.LittleEndian.PutUint32(buf[60:], miniFatStart)
	binary.LittleEndian.PutUint32(buf[64:], miniFatSectors)
	binary.LittleEndian.PutUint32(buf[68:], endOfChain) // firstDifatSectorLocation
	binary.LittleEndian.PutUint32(buf[72:], 0)          // numDifatSectors

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
func writeDirectory(dst []byte, all []*node, dirSectors uint32) {
	totalEntries := int(dirSectors) * entriesPerDirSector
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
	// CLSID(80..96) / stateBits(96..100) / 2x FILETIME(100..116) zero
	binary.LittleEndian.PutUint32(e[116:], n.startSector)
	binary.LittleEndian.PutUint64(e[120:], uint64(n.size)) // upper 4 bytes are 0 (v3)
}
