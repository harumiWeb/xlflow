// Package cfb is a reader and writer for [MS-CFB] v3 and v4. It is a
// generic CFB implementation and knows nothing about VBA.
//
// V3 sectors are 512 bytes and v4 sectors are 4096 bytes. Mini sectors are
// always 64 bytes. Streams smaller than 4096 bytes (miniStreamCutoffSize) are
// placed in the mini stream; larger streams use the regular FAT path. The
// output is not byte-exact; instead it aims for semantic equivalence: the tree
// structure, storage directory metadata, and contents of every stream match
// when read back. The header CLSID and physical sector layout are not preserved.
package cfb

import "fmt"

// Format is the on-disk CFB major version. Writers preserve the format read
// from the template instead of silently converting between sector geometries.
type Format uint16

const (
	FormatV3 Format = 3
	FormatV4 Format = 4
)

type geometry struct {
	format                Format
	sectorSize            int
	headerSpan            int
	sectorShift           uint16
	entriesPerFatSector   int
	entriesPerDirSector   int
	entriesPerDifatSector int
}

func geometryFor(format Format) (geometry, error) {
	var sectorSize int
	var sectorShift uint16
	switch format {
	case FormatV3:
		sectorSize, sectorShift = 512, 9
	case FormatV4:
		sectorSize, sectorShift = 4096, 12
	default:
		return geometry{}, fmt.Errorf("cfb: unsupported major version %d (supported: 3 and 4)", format)
	}
	return geometry{
		format:                format,
		sectorSize:            sectorSize,
		headerSpan:            sectorSize,
		sectorShift:           sectorShift,
		entriesPerFatSector:   sectorSize / 4,
		entriesPerDirSector:   sectorSize / dirEntrySize,
		entriesPerDifatSector: sectorSize/4 - 1,
	}, nil
}

// Sector and mini-sector sizes.
const (
	miniSectorSize = 64
	dirEntrySize   = 128  // byte length of one DirectoryEntry
	cutoff         = 4096 // miniStreamCutoffSize
	headerSize     = 512  // fixed header fields; v4 pads this to one 4096-byte sector
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

// CFB header constants shared by v3 and v4.
const (
	minorVersion    = 0x003E
	byteOrderMark   = 0xFFFE
	miniSectorShift = 0x0006
	difatHeaderLen  = 109 // number of DIFAT array entries at the end of the header
)

const dirNameMaxBytes = 64 // byte length of the DirectoryEntry name field (UTF-16, incl. NUL)

func forbiddenDirectoryNameUnit(unit uint16) bool {
	switch unit {
	case 0, '/', '\\', ':', '!':
		return true
	default:
		return false
	}
}
