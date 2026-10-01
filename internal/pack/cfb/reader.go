package cfb

import (
	"encoding/binary"
	"errors"
	"fmt"
	"slices"
	"unicode/utf16"
)

// StorageMeta is the metadata stored in a CFB storage directory entry. CLSID
// preserves the 16 bytes exactly as they appear on disk; Created and Modified
// are raw Windows FILETIME values.
type StorageMeta struct {
	CLSID     [16]byte
	StateBits uint32
	Created   uint64
	Modified  uint64
}

// Container holds every stream and storage read from a CFB file, keyed by
// path. The empty storage path identifies the root directory entry.
type Container struct {
	streams      map[string][]byte
	order        []string
	storages     map[string]StorageMeta
	storageOrder []string
	format       Format
}

// Stream returns the contents of the stream at path. The second result reports
// whether the stream exists.
func (c *Container) Stream(path string) ([]byte, bool) { d, ok := c.streams[path]; return d, ok }

// Paths returns all stream paths in discovery order.
func (c *Container) Paths() []string {
	return slices.Clone(c.order)
}

// Storage returns the metadata for the storage at path. The empty path
// identifies the root directory entry. The second result reports whether the
// storage exists.
func (c *Container) Storage(path string) (StorageMeta, bool) {
	meta, ok := c.storages[path]
	return meta, ok
}

// StoragePaths returns all storage paths in discovery order, including the
// empty root path as the first element.
func (c *Container) StoragePaths() []string {
	return slices.Clone(c.storageOrder)
}

// Format returns the CFB major version used by the container.
func (c *Container) Format() Format { return c.format }

type cfbHeader struct {
	geometry
	sectorCount   uint32
	firstDir      uint32
	miniCutoff    uint32
	firstMiniFat  uint32
	numMiniFat    uint32
	firstDifat    uint32
	numDifat      uint32
	numFat        uint32
	numDirSectors uint32
	difat         [difatHeaderLen]uint32
}

var cfbSig = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func parseHeader(data []byte) (*cfbHeader, error) {
	if len(data) < headerSize {
		return nil, errors.New("cfb: file shorter than 512 bytes")
	}
	for i, b := range cfbSig {
		if data[i] != b {
			return nil, fmt.Errorf("cfb: signature mismatch at offset %d", i)
		}
	}
	format := Format(binary.LittleEndian.Uint16(data[26:28]))
	g, err := geometryFor(format)
	if err != nil {
		return nil, err
	}
	if len(data) < g.headerSpan || (len(data)-g.headerSpan)%g.sectorSize != 0 {
		return nil, fmt.Errorf("cfb: file length %d is not aligned to the v%d %d-byte sector geometry", len(data), format, g.sectorSize)
	}
	sectorCount64 := uint64((len(data) - g.headerSpan) / g.sectorSize)
	if sectorCount64 > uint64(^uint32(0)) {
		return nil, fmt.Errorf("cfb: sector count %d exceeds the format limit", sectorCount64)
	}
	if binary.LittleEndian.Uint16(data[24:26]) != minorVersion {
		return nil, fmt.Errorf("cfb: unsupported minor version 0x%04X", binary.LittleEndian.Uint16(data[24:26]))
	}
	if binary.LittleEndian.Uint16(data[28:30]) != byteOrderMark {
		return nil, errors.New("cfb: invalid byte order marker")
	}
	if shift := binary.LittleEndian.Uint16(data[30:32]); shift != g.sectorShift {
		return nil, fmt.Errorf("cfb: sector shift %d does not match major version %d", shift, format)
	}
	if shift := binary.LittleEndian.Uint16(data[32:34]); shift != miniSectorShift {
		return nil, fmt.Errorf("cfb: unsupported mini sector shift %d", shift)
	}
	for _, b := range data[34:40] {
		if b != 0 {
			return nil, errors.New("cfb: reserved header bytes are nonzero")
		}
	}

	h := &cfbHeader{
		geometry:      g,
		sectorCount:   uint32(sectorCount64),
		numDirSectors: binary.LittleEndian.Uint32(data[40:44]),
		numFat:        binary.LittleEndian.Uint32(data[44:48]),
		firstDir:      binary.LittleEndian.Uint32(data[48:52]),
		miniCutoff:    binary.LittleEndian.Uint32(data[56:60]),
		firstMiniFat:  binary.LittleEndian.Uint32(data[60:64]),
		numMiniFat:    binary.LittleEndian.Uint32(data[64:68]),
		firstDifat:    binary.LittleEndian.Uint32(data[68:72]),
		numDifat:      binary.LittleEndian.Uint32(data[72:76]),
	}
	for i := range difatHeaderLen {
		h.difat[i] = binary.LittleEndian.Uint32(data[76+i*4:])
	}
	if h.miniCutoff != cutoff {
		return nil, fmt.Errorf("cfb: mini stream cutoff is %d (want %d)", h.miniCutoff, cutoff)
	}
	if h.sectorCount == 0 || h.numFat == 0 {
		return nil, errors.New("cfb: container has no sectors or FAT")
	}
	if h.numFat > h.sectorCount {
		return nil, fmt.Errorf("cfb: FAT sector count %d exceeds physical sector count %d", h.numFat, h.sectorCount)
	}
	if h.format == FormatV3 && h.numDirSectors != 0 {
		return nil, fmt.Errorf("cfb: v3 directory sector count must be zero (got %d)", h.numDirSectors)
	}
	if h.format == FormatV4 && h.numDirSectors == 0 {
		return nil, errors.New("cfb: v4 directory sector count is zero")
	}
	wantDifat := uint32(0)
	if h.numFat > difatHeaderLen {
		wantDifat = uint32(ceilDiv64(uint64(h.numFat-difatHeaderLen), uint64(h.entriesPerDifatSector)))
	}
	if h.numDifat != wantDifat {
		return nil, fmt.Errorf("cfb: DIFAT sector count is %d (want %d for %d FAT sectors)", h.numDifat, wantDifat, h.numFat)
	}
	if h.numDifat == 0 && h.firstDifat != endOfChain {
		return nil, fmt.Errorf("cfb: first DIFAT sector is 0x%08X with no DIFAT sectors", h.firstDifat)
	}
	return h, nil
}

func ceilDiv64(a, b uint64) uint64 {
	if a == 0 {
		return 0
	}
	return 1 + (a-1)/b
}

func checkedInt(v uint64, label string) (int, error) {
	if uint64(int(v)) != v {
		return 0, fmt.Errorf("cfb: %s %d exceeds platform limits", label, v)
	}
	return int(v), nil
}

func sectorSlice(data []byte, h *cfbHeader, sector uint32) ([]byte, error) {
	if sector >= h.sectorCount {
		return nil, fmt.Errorf("cfb: sector %d is outside physical sector count %d", sector, h.sectorCount)
	}
	off64 := uint64(h.headerSpan) + uint64(sector)*uint64(h.sectorSize)
	off, err := checkedInt(off64, "sector offset")
	if err != nil {
		return nil, err
	}
	return data[off : off+h.sectorSize], nil
}

func buildFAT(data []byte, h *cfbHeader) ([]uint32, []uint32, []uint32, error) {
	fatSectorNums := make([]uint32, 0, h.numFat)
	headerCount := min(int(h.numFat), difatHeaderLen)
	for i := range headerCount {
		s := h.difat[i]
		if s == freeSect || s == endOfChain || s >= h.sectorCount {
			return nil, nil, nil, fmt.Errorf("cfb: invalid header DIFAT entry 0x%08X at index %d", s, i)
		}
		fatSectorNums = append(fatSectorNums, s)
	}
	for i := headerCount; i < difatHeaderLen; i++ {
		if h.difat[i] != freeSect {
			return nil, nil, nil, fmt.Errorf("cfb: unused header DIFAT entry %d is 0x%08X (want FREESECT)", i, h.difat[i])
		}
	}

	difatSectorNums := make([]uint32, 0, h.numDifat)
	seenDifat := make([]bool, h.sectorCount)
	next := h.firstDifat
	for i := range int(h.numDifat) {
		if next >= h.sectorCount || seenDifat[next] {
			return nil, nil, nil, fmt.Errorf("cfb: invalid or cyclic DIFAT sector %d at chain index %d", next, i)
		}
		seenDifat[next] = true
		difatSectorNums = append(difatSectorNums, next)
		sec, err := sectorSlice(data, h, next)
		if err != nil {
			return nil, nil, nil, err
		}
		remaining := int(h.numFat) - len(fatSectorNums)
		used := min(remaining, h.entriesPerDifatSector)
		for j := range used {
			s := binary.LittleEndian.Uint32(sec[j*4:])
			if s == freeSect || s == endOfChain || s >= h.sectorCount {
				return nil, nil, nil, fmt.Errorf("cfb: invalid DIFAT FAT-sector entry 0x%08X at sector %d index %d", s, next, j)
			}
			fatSectorNums = append(fatSectorNums, s)
		}
		for j := used; j < h.entriesPerDifatSector; j++ {
			if v := binary.LittleEndian.Uint32(sec[j*4:]); v != freeSect {
				return nil, nil, nil, fmt.Errorf("cfb: unused DIFAT entry at sector %d index %d is 0x%08X", next, j, v)
			}
		}
		next = binary.LittleEndian.Uint32(sec[h.entriesPerDifatSector*4:])
	}
	if next != endOfChain {
		return nil, nil, nil, fmt.Errorf("cfb: DIFAT chain ends with 0x%08X instead of ENDOFCHAIN", next)
	}
	if len(fatSectorNums) != int(h.numFat) {
		return nil, nil, nil, fmt.Errorf("cfb: collected %d FAT sectors (want %d)", len(fatSectorNums), h.numFat)
	}

	seenFat := make([]bool, h.sectorCount)
	for _, s := range fatSectorNums {
		if seenFat[s] || seenDifat[s] {
			return nil, nil, nil, fmt.Errorf("cfb: FAT/DIFAT sector %d is referenced more than once", s)
		}
		seenFat[s] = true
	}
	fatLen, err := checkedInt(uint64(h.numFat)*uint64(h.entriesPerFatSector), "FAT entry count")
	if err != nil {
		return nil, nil, nil, err
	}
	if uint64(fatLen) < uint64(h.sectorCount) {
		return nil, nil, nil, fmt.Errorf("cfb: FAT covers %d sectors, fewer than the %d physical sectors", fatLen, h.sectorCount)
	}
	fat := make([]uint32, 0, fatLen)
	for _, s := range fatSectorNums {
		sec, err := sectorSlice(data, h, s)
		if err != nil {
			return nil, nil, nil, err
		}
		for i := range h.entriesPerFatSector {
			fat = append(fat, binary.LittleEndian.Uint32(sec[i*4:]))
		}
	}
	for sector := range h.sectorCount {
		v := fat[sector]
		if v < h.sectorCount || v == freeSect || v == endOfChain || v == fatSect || v == difSect {
			continue
		}
		return nil, nil, nil, fmt.Errorf("cfb: FAT entry for sector %d has invalid value 0x%08X", sector, v)
	}
	for _, s := range fatSectorNums {
		if fat[s] != fatSect {
			return nil, nil, nil, fmt.Errorf("cfb: FAT sector %d is marked 0x%08X instead of FATSECT", s, fat[s])
		}
	}
	for _, s := range difatSectorNums {
		if fat[s] != difSect {
			return nil, nil, nil, fmt.Errorf("cfb: DIFAT sector %d is marked 0x%08X instead of DIFSECT", s, fat[s])
		}
	}
	return fat, fatSectorNums, difatSectorNums, nil
}

func readRegularChain(data []byte, h *cfbHeader, fat []uint32, owners []string, start uint32, expected int, owner string) ([]byte, error) {
	if expected == 0 {
		if start != endOfChain && start != freeSect {
			return nil, fmt.Errorf("cfb: empty %s starts at sector %d", owner, start)
		}
		return nil, nil
	}
	if start == endOfChain || start == freeSect {
		return nil, fmt.Errorf("cfb: non-empty %s has no starting sector", owner)
	}
	capacity := 0
	if expected > 0 {
		if uint64(expected) > uint64(h.sectorCount) {
			return nil, fmt.Errorf("cfb: %s declares %d sectors, exceeding the %d physical sectors", owner, expected, h.sectorCount)
		}
		var err error
		capacity, err = checkedInt(uint64(expected)*uint64(h.sectorSize), owner+" chain size")
		if err != nil {
			return nil, err
		}
	}
	out := make([]byte, 0, capacity)
	count := 0
	for sector := start; ; {
		if sector >= h.sectorCount {
			return nil, fmt.Errorf("cfb: %s references out-of-range sector %d", owner, sector)
		}
		if owners[sector] != "" {
			return nil, fmt.Errorf("cfb: sector %d is shared by %s and %s", sector, owners[sector], owner)
		}
		owners[sector] = owner
		sec, err := sectorSlice(data, h, sector)
		if err != nil {
			return nil, err
		}
		out = append(out, sec...)
		count++
		next := fat[sector]
		if expected >= 0 && count == expected {
			if next != endOfChain {
				return nil, fmt.Errorf("cfb: %s chain exceeds declared %d sectors", owner, expected)
			}
			break
		}
		if next == endOfChain {
			if expected >= 0 {
				return nil, fmt.Errorf("cfb: %s chain has %d sectors (want %d)", owner, count, expected)
			}
			break
		}
		if next == freeSect || next == fatSect || next == difSect {
			return nil, fmt.Errorf("cfb: %s chain contains invalid marker 0x%08X", owner, next)
		}
		sector = next
	}
	return out, nil
}

func readMiniChain(miniData []byte, miniFAT []uint32, owners []string, start uint32, size uint64, owner string) ([]byte, error) {
	expected, err := checkedInt(ceilDiv64(size, miniSectorSize), owner+" mini-sector count")
	if err != nil {
		return nil, err
	}
	if start == endOfChain || start == freeSect {
		return nil, fmt.Errorf("cfb: non-empty %s has no mini-sector", owner)
	}
	out := make([]byte, 0, expected*miniSectorSize)
	sector := start
	for count := range expected {
		if sector >= uint32(len(miniFAT)) {
			return nil, fmt.Errorf("cfb: %s references out-of-range mini-sector %d", owner, sector)
		}
		if owners[sector] != "" {
			return nil, fmt.Errorf("cfb: mini-sector %d is shared by %s and %s", sector, owners[sector], owner)
		}
		owners[sector] = owner
		off := int(sector) * miniSectorSize
		if off+miniSectorSize > len(miniData) {
			return nil, fmt.Errorf("cfb: %s mini-sector %d exceeds the root mini stream", owner, sector)
		}
		out = append(out, miniData[off:off+miniSectorSize]...)
		next := miniFAT[sector]
		if count == expected-1 {
			if next != endOfChain {
				return nil, fmt.Errorf("cfb: %s mini-chain exceeds declared size", owner)
			}
			break
		}
		if next == endOfChain || next == freeSect || next == fatSect || next == difSect {
			return nil, fmt.Errorf("cfb: %s mini-chain ended after %d of %d sectors", owner, count+1, expected)
		}
		sector = next
	}
	contentSize, err := checkedInt(size, owner+" size")
	if err != nil {
		return nil, err
	}
	return out[:contentSize], nil
}

type dirEntry struct {
	name               string
	objType            byte
	left, right, child uint32
	start              uint32
	size               uint64
	meta               StorageMeta
}

func parseDirEntry(raw []byte, format Format) (dirEntry, error) {
	nameLen := int(binary.LittleEndian.Uint16(raw[64:66]))
	if nameLen < 2 || nameLen > dirNameMaxBytes || nameLen%2 != 0 {
		if raw[66] == objUnknown && nameLen == 0 {
			return dirEntry{objType: objUnknown, left: noStream, right: noStream, child: noStream}, nil
		}
		return dirEntry{}, fmt.Errorf("cfb: invalid directory name length %d", nameLen)
	}
	if binary.LittleEndian.Uint16(raw[nameLen-2:nameLen]) != 0 {
		return dirEntry{}, errors.New("cfb: directory name is not NUL-terminated")
	}
	units := make([]uint16, (nameLen-2)/2)
	for i := range units {
		units[i] = binary.LittleEndian.Uint16(raw[i*2:])
		if forbiddenDirectoryNameUnit(units[i]) {
			return dirEntry{}, fmt.Errorf("cfb: directory name contains forbidden character %q", rune(units[i]))
		}
	}
	size := binary.LittleEndian.Uint64(raw[120:128])
	if format == FormatV3 {
		size &= uint64(^uint32(0))
	}
	entry := dirEntry{
		name:    string(utf16.Decode(units)),
		objType: raw[66],
		left:    binary.LittleEndian.Uint32(raw[68:72]),
		right:   binary.LittleEndian.Uint32(raw[72:76]),
		child:   binary.LittleEndian.Uint32(raw[76:80]),
		start:   binary.LittleEndian.Uint32(raw[116:120]),
		size:    size,
		meta: StorageMeta{
			StateBits: binary.LittleEndian.Uint32(raw[96:100]),
			Created:   binary.LittleEndian.Uint64(raw[100:108]),
			Modified:  binary.LittleEndian.Uint64(raw[108:116]),
		},
	}
	copy(entry.meta.CLSID[:], raw[80:96])
	if entry.objType == objStream {
		if entry.meta.CLSID != ([16]byte{}) {
			return dirEntry{}, errors.New("cfb: stream directory entry has a nonzero CLSID")
		}
		if entry.meta.Created != 0 || entry.meta.Modified != 0 {
			return dirEntry{}, errors.New("cfb: stream directory entry has a nonzero FILETIME")
		}
	}
	if entry.objType == objRoot && entry.meta.Created != 0 {
		return dirEntry{}, errors.New("cfb: root directory entry has a nonzero creation FILETIME")
	}
	return entry, nil
}

// Open parses a CFB container and reconstructs all streams.
func Open(data []byte) (*Container, error) {
	h, err := parseHeader(data)
	if err != nil {
		return nil, err
	}
	fat, fatSectors, difatSectors, err := buildFAT(data, h)
	if err != nil {
		return nil, err
	}
	owners := make([]string, h.sectorCount)
	for _, sector := range fatSectors {
		owners[sector] = "FAT"
	}
	for _, sector := range difatSectors {
		owners[sector] = "DIFAT"
	}
	dirExpected := -1
	if h.format == FormatV4 {
		dirExpected, err = checkedInt(uint64(h.numDirSectors), "directory sector count")
		if err != nil {
			return nil, err
		}
	}
	dirData, err := readRegularChain(data, h, fat, owners, h.firstDir, dirExpected, "directory")
	if err != nil {
		return nil, err
	}
	if len(dirData) == 0 || len(dirData)%dirEntrySize != 0 {
		return nil, errors.New("cfb: invalid or empty directory stream")
	}
	entries := make([]dirEntry, len(dirData)/dirEntrySize)
	for i := range entries {
		entries[i], err = parseDirEntry(dirData[i*dirEntrySize:(i+1)*dirEntrySize], h.format)
		if err != nil {
			return nil, fmt.Errorf("cfb: directory entry %d: %w", i, err)
		}
	}
	root := entries[0]
	if root.objType != objRoot {
		return nil, fmt.Errorf("cfb: entry 0 is not the Root Entry (objType=%d)", root.objType)
	}

	var miniData []byte
	if root.size > 0 {
		expected, err := checkedInt(ceilDiv64(root.size, uint64(h.sectorSize)), "root mini stream sector count")
		if err != nil {
			return nil, err
		}
		raw, err := readRegularChain(data, h, fat, owners, root.start, expected, "root mini stream")
		if err != nil {
			return nil, err
		}
		size, err := checkedInt(root.size, "root mini stream size")
		if err != nil {
			return nil, err
		}
		miniData = raw[:size]
	}

	var miniFAT []uint32
	if h.numMiniFat == 0 {
		if h.firstMiniFat != endOfChain && h.firstMiniFat != freeSect {
			return nil, errors.New("cfb: mini-FAT start is set while its sector count is zero")
		}
	} else {
		miniFATCount, countErr := checkedInt(uint64(h.numMiniFat), "mini-FAT sector count")
		if countErr != nil {
			return nil, countErr
		}
		raw, err := readRegularChain(data, h, fat, owners, h.firstMiniFat, miniFATCount, "mini-FAT")
		if err != nil {
			return nil, err
		}
		miniFAT = make([]uint32, len(raw)/4)
		for i := range miniFAT {
			miniFAT[i] = binary.LittleEndian.Uint32(raw[i*4:])
		}
	}
	miniOwners := make([]string, len(miniFAT))
	c := &Container{
		streams:      map[string][]byte{},
		storages:     map[string]StorageMeta{"": root.meta},
		storageOrder: []string{""},
		format:       h.format,
	}

	type walkFrame struct {
		index  uint32
		prefix string
		phase  byte
	}
	states := make([]byte, len(entries))
	states[0] = 2
	stack := []walkFrame{{index: root.child}}
	type childIdentity struct {
		name    string
		objType byte
	}
	childrenByParent := map[string]map[string]childIdentity{}
	pathBytes := 0
	pathBudget := max(len(data)*8, 1<<20)
	for len(stack) > 0 {
		frame := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if frame.index == noStream {
			continue
		}
		if frame.index >= uint32(len(entries)) {
			return nil, fmt.Errorf("cfb: directory reference %d is out of range", frame.index)
		}
		entry := entries[frame.index]
		if frame.phase == 0 {
			if states[frame.index] != 0 {
				return nil, fmt.Errorf("cfb: cyclic or duplicate directory reference to entry %d", frame.index)
			}
			states[frame.index] = 1
			stack = append(stack, walkFrame{index: frame.index, prefix: frame.prefix, phase: 1})
			stack = append(stack, walkFrame{index: entry.left, prefix: frame.prefix})
			continue
		}
		states[frame.index] = 2
		stack = append(stack, walkFrame{index: entry.right, prefix: frame.prefix})
		if entry.name == "" {
			return nil, fmt.Errorf("cfb: referenced directory entry %d has an empty name", frame.index)
		}
		children := childrenByParent[frame.prefix]
		if children == nil {
			children = map[string]childIdentity{}
			childrenByParent[frame.prefix] = children
		}
		identity := DirectoryNameKey(entry.name)
		if previous, exists := children[identity]; exists {
			return nil, fmt.Errorf(
				"cfb: duplicate directory identity %q (type %d) and %q (type %d) under %q",
				previous.name, previous.objType, entry.name, entry.objType, frame.prefix,
			)
		}
		children[identity] = childIdentity{name: entry.name, objType: entry.objType}
		switch entry.objType {
		case objStorage:
			path := frame.prefix + entry.name
			pathBytes += len(path)
			if pathBytes > pathBudget {
				return nil, fmt.Errorf("cfb: aggregate directory paths exceed the %d-byte safety budget", pathBudget)
			}
			if _, exists := c.storages[path]; exists {
				return nil, fmt.Errorf("cfb: duplicate storage path %q", path)
			}
			c.storages[path] = entry.meta
			c.storageOrder = append(c.storageOrder, path)
			stack = append(stack, walkFrame{index: entry.child, prefix: path + "/"})
		case objStream:
			path := frame.prefix + entry.name
			pathBytes += len(path)
			if pathBytes > pathBudget {
				return nil, fmt.Errorf("cfb: aggregate directory paths exceed the %d-byte safety budget", pathBudget)
			}
			if _, exists := c.streams[path]; exists {
				return nil, fmt.Errorf("cfb: duplicate stream path %q", path)
			}
			var content []byte
			switch {
			case entry.size == 0:
				content = []byte{}
			case entry.size < uint64(h.miniCutoff):
				content, err = readMiniChain(miniData, miniFAT, miniOwners, entry.start, entry.size, path)
			default:
				expected, countErr := checkedInt(ceilDiv64(entry.size, uint64(h.sectorSize)), path+" sector count")
				if countErr != nil {
					return nil, countErr
				}
				var raw []byte
				raw, err = readRegularChain(data, h, fat, owners, entry.start, expected, path)
				if err == nil {
					var size int
					size, err = checkedInt(entry.size, path+" size")
					if err == nil {
						content = raw[:size]
					}
				}
			}
			if err != nil {
				return nil, err
			}
			c.streams[path] = content
			c.order = append(c.order, path)
		default:
			return nil, fmt.Errorf("cfb: referenced directory entry %d has invalid object type %d", frame.index, entry.objType)
		}
	}
	return c, nil
}
