package vbaproject

import (
	"fmt"
	"strings"
	"unicode/utf16"

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
		if err := ValidateWritableComponentIdentity(m.Name, m.StreamName, p.Props.CodePage); err != nil {
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
			Name:        m.Name,
			StreamName:  m.StreamName,
			TypeID:      moduleTypeID(m.Type),
			DocString:   m.DocString,
			HelpContext: m.HelpContext,
			ReadOnly:    m.ReadOnly,
			Private:     m.Private,
			Extra:       m.Extra,
		})
		projectSpecs = append(projectSpecs, ovba.ProjectComponentSpec{Kind: kind, Name: m.Name})
	}
	projectStream, err := ovba.RebuildProjectText(p.ProjectStreamRaw, projectSpecs, p.Props.CodePage)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: rebuild PROJECT stream: %w", err)
	}

	// dir.plain = PROJECTINFORMATION(span) ++ PROJECTREFERENCES(span) ++ PROJECTMODULES(built).
	dirPlain := make([]byte, 0, len(p.ProjectInfoRaw)+len(p.ReferencesRaw)+256)
	dirPlain = append(dirPlain, p.ProjectInfoRaw...)
	dirPlain = append(dirPlain, p.ReferencesRaw...)
	projectModules, err := ovba.BuildProjectModules(specs, p.Props.CodePage)
	if err != nil {
		return nil, fmt.Errorf("vbaproject: %w", err)
	}
	dirPlain = append(dirPlain, projectModules...)

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
		enc, err := ovba.EncodeMBCS(m.Source, p.Props.CodePage)
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
// imposed by the vbaProject.bin writer without mutating a project. Both names
// must be representable in the project code page (the dir stream stores them
// twice: MBCS and UTF-16), must not contain characters that corrupt the
// textual PROJECT stream, and the stream name must fit the CFB directory
// entry limit of 31 UTF-16 code units.
func ValidateWritableComponentIdentity(name, streamName string, codepage uint16) error {
	if name == "" || streamName == "" {
		return fmt.Errorf("vbaproject: module and stream names must not be empty")
	}
	for _, field := range []struct{ label, value string }{{"module", name}, {"stream", streamName}} {
		if strings.ContainsAny(field.value, "\x00\r\n") {
			return fmt.Errorf("vbaproject: %s name %q contains a NUL or newline", field.label, field.value)
		}
		if _, err := ovba.EncodeMBCS(field.value, codepage); err != nil {
			return fmt.Errorf("vbaproject: %s name %q cannot be represented in project codepage %d: %w", field.label, field.value, codepage, err)
		}
	}
	if n := len(utf16.Encode([]rune(streamName))); n > 31 {
		return fmt.Errorf("vbaproject: stream name %q needs %d UTF-16 code units (max 31 for CFB)", streamName, n)
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
