package vbaproject

import (
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

	p := &Project{
		Props: ProjectProps{
			ProjectID: pt.ID, Name: pt.Name,
			SysKind: di.SysKind, LCID: di.LCID, CodePage: di.CodePage,
		},
		Protection: Protection{
			CMG: pt.CMG, DPB: pt.DPB, GC: pt.GC,
			IsProtected: isProtected(pt.DPB),
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

// isProtected determines whether a project is protected. The spec-compliant check decrypts
// CMG/ProjectProtectionState ([MS-OVBA] §2.4.3), but this library does not decrypt the DPB, so it uses a
// heuristic: an unprotected DPB is a short obfuscated value (~16-20 chars measured in the corpus), while a
// protected one contains a key hash and is longer (over 60 chars in the protected sample). The threshold 28
// is the valley between them. Protected projects are treated as an edge case in v1 (see DESIGN.md).
func isProtected(dpb string) bool { return len(dpb) > 28 }

// firstSegment returns the first "/"-separated component of a CFB stream path
// (e.g. "VBA" for "VBA/dir", "UserForm1" for "UserForm1/\x03VBFrame", and
// "PROJECTwm" for a root-level stream).
func firstSegment(path string) string {
	if i := strings.IndexByte(path, '/'); i >= 0 {
		return path[:i]
	}
	return path
}
