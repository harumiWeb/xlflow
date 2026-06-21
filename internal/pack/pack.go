// Package pack implements xlflow's pure-Go vbaProject.bin generation engine.
//
// Reference implementation: https://github.com/kay-ws/ovba-writer
package pack

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

// ModuleType identifies the source-tree module kind supplied to GenerateVBAProject.
type ModuleType string

const (
	// ModuleTypeStandard is a standard VBA module backed by a .bas source file.
	ModuleTypeStandard ModuleType = "standard"

	// ModuleTypeClass is a class VBA module backed by a .cls source file.
	ModuleTypeClass ModuleType = "class"

	// ModuleTypeDocument is an Excel host document module, such as ThisWorkbook or a sheet module.
	ModuleTypeDocument ModuleType = "document"

	// ModuleTypeForm is a UserForm module. Its code-behind is updated from a .frm when the form
	// already exists in the template; the form's designer storage is carried verbatim and never
	// authored, so creating a new form is unsupported.
	ModuleTypeForm ModuleType = "form"
)

// SourceModule is one source-tree module to apply to a template vbaProject.bin.
//
// Source is the disk-form text read from a .bas/.cls/.frm file. Standard and class modules are
// replaced when Name and Type match an existing template module. Document modules are replaced by
// exact module name match against an existing template document module. A UserForm's code-behind is
// replaced when the form already exists in the template; its designer layout is preserved, not
// authored, and creating a new form is unsupported.
type SourceModule struct {
	Name   string
	Type   ModuleType
	Source string
}

// PackMeta summarizes the modules and opaque streams handled during pack.
type PackMeta struct {
	Standard       int
	Class          int
	Document       int
	Form           int
	CarriedStreams int
}

// GenerateVBAProject returns a regenerated vbaProject.bin based on template.
//
// The engine replaces only supplied standard, class, and unambiguous document
// module source. Project records, references, codepage, and opaque streams such
// as existing UserForm designer storages are carried through from the template.
// Unsupported content returns one of the exported sentinel errors so the CLI can
// map it to the pack error contract.
func GenerateVBAProject(template []byte, sources []SourceModule) ([]byte, error) {
	out, _, err := generateVBAProject(template, sources)
	return out, err
}

// BuildWorkbook returns a new .xlsm zip with xl/vbaProject.bin regenerated from sources.
//
// All template zip entries are copied in their original order, with their names
// and compression methods preserved. The only replaced entry is
// xl/vbaProject.bin. If the template has no vbaProject.bin entry, BuildWorkbook
// returns ErrAmbiguousLayout.
func BuildWorkbook(templateXlsm []byte, sources []SourceModule) ([]byte, PackMeta, error) {
	reader, err := zip.NewReader(bytes.NewReader(templateXlsm), int64(len(templateXlsm)))
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}

	var vbaProject []byte
	for _, entry := range reader.File {
		if entry.Name != "xl/vbaProject.bin" {
			continue
		}
		vbaProject, err = readZipEntry(entry)
		if err != nil {
			return nil, PackMeta{}, err
		}
		break
	}
	if vbaProject == nil {
		return nil, PackMeta{}, fmt.Errorf("%w: template is missing xl/vbaProject.bin", ErrAmbiguousLayout)
	}

	regenerated, meta, err := generateVBAProject(vbaProject, sources)
	if err != nil {
		return nil, PackMeta{}, err
	}

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	for _, entry := range reader.File {
		header := entry.FileHeader
		target, err := writer.CreateHeader(&header)
		if err != nil {
			_ = writer.Close()
			return nil, PackMeta{}, err
		}
		if entry.Name == "xl/vbaProject.bin" {
			if _, err := target.Write(regenerated); err != nil {
				_ = writer.Close()
				return nil, PackMeta{}, err
			}
			continue
		}
		if err := copyZipEntry(target, entry); err != nil {
			_ = writer.Close()
			return nil, PackMeta{}, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, PackMeta{}, err
	}
	return buf.Bytes(), meta, nil
}

func generateVBAProject(template []byte, sources []SourceModule) ([]byte, PackMeta, error) {
	if hasSignatureStream(template) {
		return nil, PackMeta{}, ErrSignedProject
	}
	project, err := vbaproject.Read(template)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	if project.Protection.IsProtected {
		return nil, PackMeta{}, ErrProtectedProject
	}
	meta, err := applySources(project, sources)
	if err != nil {
		return nil, PackMeta{}, err
	}
	meta.CarriedStreams = len(project.RawStreams)
	out, err := vbaproject.Write(project)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	return out, meta, nil
}

func applySources(project *vbaproject.Project, sources []SourceModule) (PackMeta, error) {
	byKey := make(map[string]int, len(project.Modules))
	for i, m := range project.Modules {
		key := moduleKey(m.Name, fromProjectModuleType(m.Type))
		if _, exists := byKey[key]; exists {
			return PackMeta{}, fmt.Errorf("%w: duplicate template module %s", ErrAmbiguousLayout, m.Name)
		}
		byKey[key] = i
	}

	var meta PackMeta
	seen := make(map[string]bool, len(sources))
	for _, source := range sources {
		if source.Name == "" {
			return PackMeta{}, fmt.Errorf("%w: source module name is empty", ErrAmbiguousLayout)
		}
		targetType, err := toProjectModuleType(source.Type)
		if err != nil {
			return PackMeta{}, err
		}
		key := moduleKey(source.Name, source.Type)
		if seen[key] {
			return PackMeta{}, fmt.Errorf("%w: duplicate source module %s", ErrAmbiguousLayout, source.Name)
		}
		seen[key] = true
		idx, ok := byKey[key]
		if !ok {
			if source.Type == ModuleTypeForm {
				// A form's designer storage cannot be authored from source, so creating a
				// form that is not already in the template is UserForm generation (Stage 3),
				// not the generic "cannot add modules" limitation.
				return PackMeta{}, fmt.Errorf("%w: form %q is not in the template; pack updates the code-behind of existing forms only and cannot create a new UserForm", ErrUserFormGenerationUnsupported, source.Name)
			}
			return PackMeta{}, fmt.Errorf("%w: source module %q is not in the template; pack updates existing modules only and cannot add new modules in the experimental MVP", ErrAmbiguousLayout, source.Name)
		}
		normalized, err := vbaproject.NormalizeModuleSource(targetType, source.Source, &project.Modules[idx])
		if err != nil {
			return PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		project.Modules[idx].Source = normalized
		switch source.Type {
		case ModuleTypeStandard:
			meta.Standard++
		case ModuleTypeClass:
			meta.Class++
		case ModuleTypeDocument:
			meta.Document++
		case ModuleTypeForm:
			meta.Form++
		}
	}
	return meta, nil
}

func toProjectModuleType(t ModuleType) (vbaproject.ModuleType, error) {
	switch t {
	case ModuleTypeStandard:
		return vbaproject.ModuleStd, nil
	case ModuleTypeClass:
		return vbaproject.ModuleClass, nil
	case ModuleTypeDocument:
		return vbaproject.ModuleDocument, nil
	case ModuleTypeForm:
		return vbaproject.ModuleForm, nil
	default:
		return 0, fmt.Errorf("%w: unsupported source module type %q", ErrAmbiguousLayout, t)
	}
}

func fromProjectModuleType(t vbaproject.ModuleType) ModuleType {
	switch t {
	case vbaproject.ModuleStd:
		return ModuleTypeStandard
	case vbaproject.ModuleClass:
		return ModuleTypeClass
	case vbaproject.ModuleDocument:
		return ModuleTypeDocument
	case vbaproject.ModuleForm:
		return ModuleTypeForm
	default:
		return ModuleType("")
	}
}

func moduleKey(name string, typ ModuleType) string {
	return string(typ) + "\x00" + name
}

func hasSignatureStream(template []byte) bool {
	container, err := cfb.Open(template)
	if err != nil {
		return false
	}
	for _, path := range container.Paths() {
		normalized := strings.ToLower(path)
		if strings.Contains(normalized, "vbaprojectsignature") ||
			strings.Contains(normalized, "_vba_project_cur") ||
			strings.Contains(normalized, "digitalsignature") {
			return true
		}
	}
	return false
}

func readZipEntry(entry *zip.File) ([]byte, error) {
	reader, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer func() { _ = reader.Close() }()
	return io.ReadAll(reader)
}

func copyZipEntry(dst io.Writer, entry *zip.File) error {
	reader, err := entry.Open()
	if err != nil {
		return err
	}
	defer func() { _ = reader.Close() }()
	_, err = io.Copy(dst, reader)
	return err
}
