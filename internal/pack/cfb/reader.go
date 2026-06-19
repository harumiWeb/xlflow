package cfb

import (
	"encoding/binary"
	"errors"
	"fmt"
	"unicode/utf16"
)

// Container holds every stream read from a CFB file, keyed by path.
type Container struct {
	streams map[string][]byte // "/"-separated path (e.g. "VBA/dir") → contents
	order   []string          // discovery order (for deterministic iteration)
}

// Stream returns the contents of the stream at the given path. The second
// return value is false if no such stream exists.
func (c *Container) Stream(path string) ([]byte, bool) { d, ok := c.streams[path]; return d, ok }

// Paths returns the paths of all stored streams in discovery order.
func (c *Container) Paths() []string { return c.order }

type cfbHeader struct {
	firstDir     uint32
	miniCutoff   uint32
	firstMiniFat uint32
	firstDifat   uint32
	difat        []uint32 // array of FAT sector numbers (109 in the header + DIFAT sectors)
}

var cfbSig = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func parseHeader(d []byte) (*cfbHeader, error) {
	if len(d) < headerSize {
		return nil, errors.New("cfb: file shorter than 512 bytes")
	}
	for i, b := range cfbSig {
		if d[i] != b {
			return nil, fmt.Errorf("cfb: signature mismatch at offset %d", i)
		}
	}
	if v := binary.LittleEndian.Uint16(d[26:]); v != 3 {
		return nil, fmt.Errorf("cfb: unsupported major version %d (only v3 is supported)", v)
	}
	h := &cfbHeader{
		firstDir:     binary.LittleEndian.Uint32(d[48:]),
		miniCutoff:   binary.LittleEndian.Uint32(d[56:]),
		firstMiniFat: binary.LittleEndian.Uint32(d[60:]),
		firstDifat:   binary.LittleEndian.Uint32(d[68:]),
	}
	for i := 0; i < difatHeaderLen; i++ {
		h.difat = append(h.difat, binary.LittleEndian.Uint32(d[76+i*4:]))
	}
	return h, nil
}

// sectorSlice returns the 512B slice of sector n within data (0-based, starting right after the header).
func sectorSlice(data []byte, n uint32) []byte {
	off := headerSize + int(n)*sectorSize
	if off+sectorSize > len(data) {
		return nil
	}
	return data[off : off+sectorSize]
}

// buildFAT collects the FAT sector numbers from the DIFAT entries and returns the FAT array.
func buildFAT(data []byte, h *cfbHeader) []uint32 {
	// Collect sector numbers from the in-header DIFAT.
	var fatSectorNums []uint32
	for _, s := range h.difat {
		if s == freeSect || s == endOfChain {
			break
		}
		fatSectorNums = append(fatSectorNums, s)
	}
	// Follow the DIFAT chain (for large files).
	// seenDifat detects self-referencing/cyclic chains to prevent an infinite loop.
	seenDifat := map[uint32]bool{}
difatChain:
	for next := h.firstDifat; next != endOfChain && next != freeSect; {
		if seenDifat[next] {
			break // cyclic DIFAT chain
		}
		seenDifat[next] = true
		sec := sectorSlice(data, next)
		if sec == nil {
			break
		}
		// DIFAT sector: entriesPerDifatSector(127) FAT sector numbers + the last 4 bytes are the next DIFAT sector number.
		for i := 0; i < entriesPerDifatSector; i++ {
			s := binary.LittleEndian.Uint32(sec[i*4:])
			if s == freeSect || s == endOfChain {
				break difatChain
			}
			fatSectorNums = append(fatSectorNums, s)
		}
		next = binary.LittleEndian.Uint32(sec[entriesPerDifatSector*4:])
	}

	// Concatenate the FAT sectors into the array.
	var fat []uint32
	for _, sn := range fatSectorNums {
		sec := sectorSlice(data, sn)
		if sec == nil {
			continue
		}
		for i := 0; i < entriesPerFatSector; i++ {
			fat = append(fat, binary.LittleEndian.Uint32(sec[i*4:]))
		}
	}
	return fat
}

// readChain follows the FAT from start to endOfChain, concatenating the sectors and returning them.
func readChain(data []byte, fat []uint32, start uint32) []byte {
	var out []byte
	seen := make([]bool, len(fat))
	for s := start; s != endOfChain && s != freeSect && int(s) < len(fat); s = fat[s] {
		if seen[s] {
			break // cyclic FAT chain
		}
		seen[s] = true
		sec := sectorSlice(data, s)
		if sec == nil {
			break
		}
		out = append(out, sec...)
	}
	return out
}

// readMiniChain assembles data from the mini-stream container following the mini-FAT chain.
func readMiniChain(miniData []byte, miniFAT []uint32, start uint32, size uint32) []byte {
	var out []byte
	seen := make([]bool, len(miniFAT))
	for s := start; s != endOfChain && s != freeSect && int(s) < len(miniFAT); s = miniFAT[s] {
		if seen[s] {
			break // cyclic mini-FAT chain
		}
		seen[s] = true
		off := int(s) * miniSectorSize
		if off+miniSectorSize > len(miniData) {
			break
		}
		out = append(out, miniData[off:off+miniSectorSize]...)
	}
	if uint32(len(out)) > size {
		out = out[:size]
	}
	return out
}

type dirEntry struct {
	name               string
	objType            byte
	left, right, child uint32
	start              uint32
	size               uint64
}

// parseDirEntry parses one 128B directory entry.
func parseDirEntry(e []byte) dirEntry {
	nameLen := int(binary.LittleEndian.Uint16(e[64:]))
	if nameLen > dirNameMaxBytes { // the name field is a fixed 64B; an excess value means a corrupt entry, so clamp it (prevents out-of-range reads)
		nameLen = dirNameMaxBytes
	}
	var name string
	if nameLen >= 2 {
		// nameLen is a byte length that includes the NUL terminator, so subtract 2 to get the number of code units
		nUnits := (nameLen - 2) / 2
		if nUnits > 0 {
			u16 := make([]uint16, nUnits)
			for i := range u16 {
				u16[i] = binary.LittleEndian.Uint16(e[i*2:])
			}
			name = string(utf16.Decode(u16))
		}
	}
	return dirEntry{
		name:    name,
		objType: e[66],
		left:    binary.LittleEndian.Uint32(e[68:]),
		right:   binary.LittleEndian.Uint32(e[72:]),
		child:   binary.LittleEndian.Uint32(e[76:]),
		start:   binary.LittleEndian.Uint32(e[116:]),
		size:    binary.LittleEndian.Uint64(e[120:]),
	}
}

// Open reads a CFB byte slice and returns a Container with all streams reconstructed.
func Open(data []byte) (*Container, error) {
	h, err := parseHeader(data)
	if err != nil {
		return nil, err
	}

	fat := buildFAT(data, h)

	// Read the directory entry array.
	dirData := readChain(data, fat, h.firstDir)
	nEntries := len(dirData) / dirEntrySize
	if nEntries == 0 {
		return nil, errors.New("cfb: no directory entries")
	}
	entries := make([]dirEntry, nEntries)
	for i := 0; i < nEntries; i++ {
		entries[i] = parseDirEntry(dirData[i*dirEntrySize:])
	}

	root := entries[0]
	if root.objType != objRoot {
		return nil, fmt.Errorf("cfb: entry 0 is not the Root Entry (objType=%d)", root.objType)
	}

	// Build the mini-stream container (the Root Entry's stream).
	var miniData []byte
	if root.start != endOfChain && root.start != freeSect && root.size > 0 {
		raw := readChain(data, fat, root.start)
		if root.size < uint64(len(raw)) {
			raw = raw[:root.size]
		}
		miniData = raw
	}

	// Build the mini-FAT.
	var miniFAT []uint32
	if h.firstMiniFat != endOfChain && h.firstMiniFat != freeSect {
		miniFATData := readChain(data, fat, h.firstMiniFat)
		miniFAT = make([]uint32, len(miniFATData)/4)
		for i := range miniFAT {
			miniFAT[i] = binary.LittleEndian.Uint32(miniFATData[i*4:])
		}
	}

	c := &Container{streams: map[string][]byte{}}

	// DFS the directory red-black tree to collect paths and stream contents.
	// An entry's left/right are BST siblings at the same level; child is the BST root of the child storage.
	visited := make([]bool, len(entries))
	var walkDir func(idx uint32, prefix string)
	walkDir = func(idx uint32, prefix string) {
		if idx == noStream || int(idx) >= len(entries) || visited[idx] {
			return // prevent infinite recursion on a cyclic directory BST
		}
		visited[idx] = true
		e := entries[idx]

		// Visit all BST nodes in order (left -> self -> right).
		walkDir(e.left, prefix)

		switch e.objType {
		case objStream:
			path := prefix + e.name
			var content []byte
			if e.size == 0 {
				content = []byte{}
			} else if e.size < uint64(h.miniCutoff) && len(miniData) > 0 {
				content = readMiniChain(miniData, miniFAT, e.start, uint32(e.size))
			} else if e.start != endOfChain && e.start != freeSect {
				raw := readChain(data, fat, e.start)
				sz := e.size
				if sz > uint64(len(raw)) {
					sz = uint64(len(raw))
				}
				content = raw[:sz]
			} else {
				content = []byte{}
			}
			c.streams[path] = content
			c.order = append(c.order, path)

		case objStorage:
			walkDir(e.child, prefix+e.name+"/")
		}

		walkDir(e.right, prefix)
	}

	walkDir(root.child, "")

	return c, nil
}
