// Package cfb is a minimal-profile reader and writer for [MS-CFB] v3. It is a
// generic CFB implementation and knows nothing about VBA.
//
// Sectors are 512 bytes (sectorShift=9) and mini sectors are 64 bytes
// (miniSectorShift=6). Streams smaller than 4096 bytes (miniStreamCutoffSize)
// are placed in the mini stream; larger streams use the regular FAT path
// (stored directly in 512-byte sectors). The output is not byte-exact; instead
// it aims for "semantic equivalence": when read back with richardlehane/mscfb,
// the tree structure and the contents of every stream match.
package cfb

// Sector and mini-sector sizes.
const (
	sectorSize     = 512
	miniSectorSize = 64
	dirEntrySize   = 128  // byte length of one DirectoryEntry
	cutoff         = 4096 // miniStreamCutoffSize
)

// Special FAT / DIFAT sector values ([MS-CFB] §2.2).
const (
	freeSect   = 0xFFFFFFFF // free (unused)
	endOfChain = 0xFFFFFFFE // end of chain
	fatSect    = 0xFFFFFFFD // sector occupied by the FAT itself
	difSect    = 0xFFFFFFFC // sector occupied by the DIFAT itself
	noStream   = 0xFFFFFFFF // no sibling/child in the directory
)

// DirectoryEntry objectType values ([MS-CFB] §2.6.1).
const (
	objUnknown = 0
	objStorage = 1
	objStream  = 2
	objRoot    = 5
)

// CFB v3 header constants.
const (
	headerSize      = 512
	minorVersion    = 0x003E
	majorVersion    = 0x0003
	byteOrderMark   = 0xFFFE
	sectorShift     = 0x0009
	miniSectorShift = 0x0006
	difatHeaderLen  = 109 // number of DIFAT array entries at the end of the header
)

// Constants for counting 512B sectors on 4096B boundaries, to avoid magic numbers.
const (
	entriesPerFatSector   = sectorSize / 4            // 128
	entriesPerDirSector   = sectorSize / dirEntrySize // 4
	entriesPerDifatSector = entriesPerFatSector - 1   // 127: a DIFAT sector uses its last 4 bytes for the next DIFAT sector number
	dirNameMaxBytes       = 64                        // byte length of the DirectoryEntry name field (UTF-16, incl. NUL)
)
