package vbaproject

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

// Write assembles a *Project (whose Source fields hold the in-bin form) into a
// source-only vbaProject.bin. ProjectInfoRaw and ReferencesRaw are written
// verbatim. The PROJECT stream preserves unrelated text while rebuilding its
// component declarations; PROJECTMODULES and module streams are rebuilt from
// the model.
func Write(p *Project) ([]byte, error) {
	if p.Protection.IsProtected {
		return nil, fmt.Errorf("vbaproject: protected projects are not supported in v1")
	}
	specs := make([]ovba.ModuleSpec, 0, len(p.Modules))
	projectSpecs := make([]ovba.ProjectComponentSpec, 0, len(p.Modules))
	names := make(map[string]string, len(p.Modules))
	streams := make(map[string]string, len(p.Modules))
	for _, m := range p.Modules {
		// A UserForm's code-behind module is written like any other module (its
		// in-bin source is an Attribute header plus code, the same shape as a class
		// or document module). The form's designer storage (UserForm1/...) is not
		// modeled here; it is carried verbatim via RawStreams (see Read).
		// Module/stream names are assumed to be within ASCII. Only ASCII can be written losslessly to both
		// the dir MBCS field (CODEPAGE written directly) and the UNICODE field (utf16le); for non-ASCII,
		// utf16le would truncate non-BMP characters and MBCS would ignore the CODEPAGE, so it is rejected.
		if err := ValidateWritableComponentIdentity(m.Name, m.StreamName); err != nil {
			return nil, err
		}
		nameKey := strings.ToLower(m.Name)
		if prior, ok := names[nameKey]; ok {
			return nil, fmt.Errorf("vbaproject: duplicate module names %q and %q", prior, m.Name)
		}
		names[nameKey] = m.Name
		streamKey := strings.ToLower(m.StreamName)
		if prior, ok := streams[streamKey]; ok {
			return nil, fmt.Errorf("vbaproject: duplicate module stream names %q and %q", prior, m.StreamName)
		}
		streams[streamKey] = m.StreamName
		kind := projectKind(m.Type)
		if kind == "" {
			return nil, fmt.Errorf("vbaproject: module %q has an unknown ModuleType %d", m.Name, m.Type)
		}
		specs = append(specs, ovba.ModuleSpec{
			Name: m.Name, StreamName: m.StreamName, TypeID: moduleTypeID(m.Type),
		})
		projectSpecs = append(projectSpecs, ovba.ProjectComponentSpec{Kind: kind, Name: m.Name})
	}
	projectStream, err := ovba.RebuildProjectText(p.ProjectStreamRaw, projectSpecs)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: rebuild PROJECT stream: %w", err)
	}

	// dir.plain = PROJECTINFORMATION(span) ++ PROJECTREFERENCES(span) ++ PROJECTMODULES(built).
	dirPlain := make([]byte, 0, len(p.ProjectInfoRaw)+len(p.ReferencesRaw)+256)
	dirPlain = append(dirPlain, p.ProjectInfoRaw...)
	dirPlain = append(dirPlain, p.ReferencesRaw...)
	dirPlain = append(dirPlain, ovba.BuildProjectModules(specs)...)

	w := cfb.NewWriter()
	// Pass through every stream the writer does not own (root-level designer
	// storages, PROJECTwm, ...) verbatim before adding the regenerated VBA/* and
	// PROJECT. RawStreams excludes the VBA/ and PROJECT namespaces, so there is no
	// collision with the AddStream calls below.
	for path, data := range p.RawStreams {
		w.AddStream(strings.Split(path, "/"), data)
	}
	w.AddStream([]string{"PROJECT"}, projectStream)
	w.AddStream([]string{"VBA", "_VBA_PROJECT"}, ovba.VBAProjectStub())
	dirComp, err := ovba.Compress(dirPlain)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: compress dir: %w", err)
	}
	w.AddStream([]string{"VBA", "dir"}, dirComp)
	for _, m := range p.Modules {
		enc, err := encodeMBCS(m.Source, p.Props.CodePage)
		if err != nil {
			return nil, fmt.Errorf("vbaproject: module %q: %w", m.Name, err)
		}
		comp, err := ovba.Compress(enc)
		if err != nil {
			return nil, fmt.Errorf("vbaproject: module %q: %w", m.Name, err)
		}
		w.AddStream([]string{"VBA", m.StreamName}, comp)
	}
	return w.Bytes()
}

// ValidateWritableComponentIdentity checks the name/stream-name restrictions
// imposed by the current vbaProject.bin writer without mutating a project.
func ValidateWritableComponentIdentity(name, streamName string) error {
	if !isASCII(name) || !isASCII(streamName) {
		return fmt.Errorf("vbaproject: non-ASCII module names are not supported in v1 (Name=%q StreamName=%q)", name, streamName)
	}
	if name == "" || streamName == "" {
		return fmt.Errorf("vbaproject: module and stream names must not be empty")
	}
	return nil
}

// moduleTypeID maps the model's ModuleType to the dir MODULETYPE value.
// Std=0x0021 (procedural), Class/Document=0x0022 (non-procedural).
func moduleTypeID(t ModuleType) uint16 {
	if t == ModuleStd {
		return 0x0021
	}
	return 0x0022
}

// isASCII reports whether s is entirely ASCII (0x00-0x7F).
func isASCII(s string) bool {
	for i := 0; i < len(s); i++ {
		if s[i] >= 0x80 {
			return false
		}
	}
	return true
}

// projectKind maps a ModuleType to the kind keyword PROJECT uses
// (Module/Class/BaseClass/Document, per ovba.ParseProjectText).
func projectKind(t ModuleType) string {
	switch t {
	case ModuleStd:
		return "Module"
	case ModuleClass:
		return "Class"
	case ModuleDocument:
		return "Document"
	case ModuleForm:
		return "BaseClass"
	default:
		return ""
	}
}
