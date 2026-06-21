package vbaproject

import (
	"encoding/binary"
	"encoding/hex"
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

// Read parses vbaProject.bin and decomposes it into a *Project model.
func Read(data []byte) (*Project, error) {
	c, err := cfb.Open(data)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: cfb open: %w", err)
	}
	projRaw, ok := c.Stream("PROJECT")
	if !ok {
		return nil, fmt.Errorf("vbaproject: missing PROJECT stream")
	}
	pt := ovba.ParseProjectText(projRaw)

	dirComp, ok := c.Stream("VBA/dir")
	if !ok {
		return nil, fmt.Errorf("vbaproject: missing VBA/dir stream")
	}
	dirPlain, err := ovba.Decompress(dirComp)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: dir decompress: %w", err)
	}
	di := ovba.ParseDir(dirPlain)

	protected, err := isProtected(pt.CMG)
	if err != nil {
		return nil, err
	}
	p := &Project{
		Props: ProjectProps{
			ProjectID: pt.ID, Name: pt.Name,
			SysKind: di.SysKind, LCID: di.LCID, CodePage: di.CodePage,
		},
		Protection: Protection{
			CMG: pt.CMG, DPB: pt.DPB, GC: pt.GC,
			IsProtected: protected,
		},
	}
	for _, name := range di.RefNames {
		p.References = append(p.References, Reference{Name: name})
	}
	p.ReferencesRaw = di.RefsRaw
	p.ProjectInfoRaw = di.ProjectInfoRaw
	p.ProjectStreamRaw = projRaw
	for _, dm := range di.Modules {
		stream := dm.StreamName
		if stream == "" {
			stream = dm.Name
		}
		raw, ok := c.Stream("VBA/" + stream)
		if !ok {
			return nil, fmt.Errorf("vbaproject: missing module stream VBA/%s", stream)
		}
		if int(dm.Offset) > len(raw) {
			return nil, fmt.Errorf("vbaproject: %s offset %d > stream %d", dm.Name, dm.Offset, len(raw))
		}
		plain, err := ovba.Decompress(raw[dm.Offset:])
		if err != nil {
			return nil, fmt.Errorf("vbaproject: %s decompress: %w", dm.Name, err)
		}
		src, err := decodeMBCS(plain, di.CodePage)
		if err != nil {
			return nil, fmt.Errorf("vbaproject: %s decode: %w", dm.Name, err)
		}
		p.Modules = append(p.Modules, Module{
			Name:       dm.Name,
			StreamName: stream,
			Type:       classify(dm.Name, pt.Kinds[dm.Name], c),
			Source:     src,
		})
	}

	// Capture every stream the writer does not own (see Project.RawStreams) so
	// Write can re-emit it verbatim. VBA/* is regenerated and PROJECT is written
	// back from ProjectStreamRaw, so both are excluded to avoid a write collision.
	p.RawStreams = make(map[string][]byte)
	for _, path := range c.Paths() {
		switch firstSegment(path) {
		case "VBA", "PROJECT":
			// owned by Write
		default:
			if s, ok := c.Stream(path); ok {
				p.RawStreams[path] = s
			}
		}
	}
	return p, nil
}

// classify decides the module type primarily from the PROJECT key, secondarily from the presence of a CFB sub-storage.
func classify(name, projectKey string, c *cfb.Container) ModuleType {
	switch projectKey {
	case "Module":
		return ModuleStd
	case "Class":
		return ModuleClass
	case "Document":
		return ModuleDocument
	case "BaseClass":
		return ModuleForm
	}
	// Fallback when absent from PROJECT: if a same-named sub-storage exists, it is a form.
	if _, ok := c.Stream(name + "/\x03VBFrame"); ok {
		return ModuleForm
	}
	return ModuleClass
}

// ProjectProtectionState bits (MS-OVBA §2.3.1.15). The decrypted CMG Data is a
// 4-byte little-endian value; any of these bits means the project's VBA is locked
// for viewing, which pack cannot regenerate from source, so such projects are
// rejected (see write.go / pack.go).
const (
	fUserProtected uint32 = 1 << 0
	fHostProtected uint32 = 1 << 1
	fVBAProtected  uint32 = 1 << 2
)

// isProtected reports whether the project's CMG (ProjectProtectionState, §2.3.1.15)
// marks it as protected. cmg is the unquoted hex string from the PROJECT stream.
// An earlier heuristic measured DPB (password) length; this reads the actual
// protection bits by decrypting CMG (reversible obfuscation, no password needed).
//
// An empty CMG means the project has no protection record and is unprotected. A
// CMG that cannot be hex-decoded or decrypted, or whose state is not 4 bytes, is a
// fail-loud error rather than a silent misclassification in either direction.
func isProtected(cmg string) (bool, error) {
	if cmg == "" {
		return false, nil
	}
	raw, err := hex.DecodeString(cmg)
	if err != nil {
		return false, fmt.Errorf("vbaproject: CMG hex decode: %w", err)
	}
	// CMG always decrypts to a 4-byte ProjectProtectionState (§2.3.1.15), so the
	// encrypted structure has a fixed size: Seed+VersionEnc+ProjKeyEnc (3) +
	// IgnoredEnc + LengthEnc (4) + DataEnc (4). Validate that up front so a large
	// but internally consistent CMG cannot force O(size) decrypt work before the
	// 4-byte state check rejects it.
	if len(raw) < 11 {
		return false, fmt.Errorf("vbaproject: CMG too short (%d bytes)", len(raw))
	}
	ignoredLen := int((raw[0] & 0x06) >> 1)
	if want := 3 + ignoredLen + 4 + 4; len(raw) != want {
		return false, fmt.Errorf("vbaproject: CMG length is %d bytes (want %d)", len(raw), want)
	}
	state, err := ovba.DecryptData(raw)
	if err != nil {
		return false, fmt.Errorf("vbaproject: CMG decrypt: %w", err)
	}
	if len(state) != 4 {
		return false, fmt.Errorf("vbaproject: CMG protection state is %d bytes (want 4)", len(state))
	}
	bits := binary.LittleEndian.Uint32(state)
	return bits&(fUserProtected|fHostProtected|fVBAProtected) != 0, nil
}

// firstSegment returns the first "/"-separated component of a CFB stream path
// (e.g. "VBA" for "VBA/dir", "UserForm1" for "UserForm1/\x03VBFrame", and
// "PROJECTwm" for a root-level stream).
func firstSegment(path string) string {
	if i := strings.IndexByte(path, '/'); i >= 0 {
		return path[:i]
	}
	return path
}
