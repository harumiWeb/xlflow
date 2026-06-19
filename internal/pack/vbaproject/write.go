package vbaproject

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

// Write assembles a *Project (whose Source fields hold the in-bin form) into a
// source-only vbaProject.bin. The preserved set (ProjectInfoRaw, ReferencesRaw,
// ProjectStreamRaw) is written verbatim; only PROJECTMODULES and the module
// streams are built from the model.
func Write(p *Project) ([]byte, error) {
	if p.Protection.IsProtected {
		return nil, fmt.Errorf("vbaproject: protected projects are not supported in v1")
	}
	specs := make([]ovba.ModuleSpec, 0, len(p.Modules))
	for _, m := range p.Modules {
		// A UserForm's code-behind module is written like any other module (its
		// in-bin source is an Attribute header plus code, the same shape as a class
		// or document module). The form's designer storage (UserForm1/...) is not
		// modeled here; it is carried verbatim via RawStreams (see Read).
		// Module/stream names are assumed to be within ASCII. Only ASCII can be written losslessly to both
		// the dir MBCS field (CODEPAGE written directly) and the UNICODE field (utf16le); for non-ASCII,
		// utf16le would truncate non-BMP characters and MBCS would ignore the CODEPAGE, so it is rejected.
		if !isASCII(m.Name) || !isASCII(m.StreamName) {
			return nil, fmt.Errorf("vbaproject: non-ASCII module names are not supported in v1 (Name=%q StreamName=%q)", m.Name, m.StreamName)
		}
		specs = append(specs, ovba.ModuleSpec{
			Name: m.Name, StreamName: m.StreamName, TypeID: moduleTypeID(m.Type),
		})
	}
	// v1 is for source editing only. The PROJECT stream is preserved verbatim, but PROJECTMODULES is
	// rebuilt from p.Modules, so if the module set disagrees with PROJECT the output bin becomes internally
	// inconsistent. Adding/removing/renaming a module is an unsupported operation and is rejected fail-loud.
	if err := checkModuleSet(p); err != nil {
		return nil, err
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
	w.AddStream([]string{"PROJECT"}, p.ProjectStreamRaw)
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

// checkModuleSet verifies that the set of module names listed by the PROJECT stream being preserved
// matches p.Modules. PROJECT enumerates modules via Module=/Class=/BaseClass=/Document= lines, and any
// mismatch with p.Modules (the source for rebuilding PROJECTMODULES) would be an internal inconsistency.
func checkModuleSet(p *Project) error {
	listed := ovba.ParseProjectText(p.ProjectStreamRaw).Kinds
	have := make(map[string]bool, len(p.Modules))
	for _, m := range p.Modules {
		have[m.Name] = true
		kind, ok := listed[m.Name]
		if !ok {
			return fmt.Errorf("vbaproject: module %q is not listed in the PROJECT stream (changing the module set is not supported in v1)", m.Name)
		}
		want := projectKind(m.Type)
		if want == "" {
			return fmt.Errorf("vbaproject: module %q has an unknown ModuleType %d", m.Name, m.Type)
		}
		if kind != want {
			return fmt.Errorf("vbaproject: module %q kind mismatch: PROJECT declares %s but the model is %s (changing a module's kind is not supported in v1)", m.Name, kind, want)
		}
	}
	for name := range listed {
		if !have[name] {
			return fmt.Errorf("vbaproject: module %q from the PROJECT stream is missing from p.Modules (changing the module set is not supported in v1)", name)
		}
	}
	return nil
}

// projectKind maps a ModuleType to the kind keyword PROJECT uses
// (Module/Class/BaseClass/Document, per ovba.ParseProjectText). It returns ""
// for a ModuleType with no PROJECT keyword; checkModuleSet treats that as an
// error rather than silently skipping the kind check.
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
