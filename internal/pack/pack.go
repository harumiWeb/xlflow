// Package pack implements xlflow's pure-Go vbaProject.bin generation engine.
//
// Reference implementation: https://github.com/kay-ws/ovba-writer
package pack

import (
	"fmt"
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

	// ModuleTypeForm is a UserForm module. It is detected to fail loudly because MVP pack never generates forms.
	ModuleTypeForm ModuleType = "form"
)

// SourceModule is one source-tree module to apply to a template vbaProject.bin.
//
// Source must be the disk-form text read from a .bas or .cls file. Standard and
// class modules are replaced when Name and Type match an existing template
// module. Document modules are replaced only by exact module name match against
// an existing template document module. UserForm modules are unsupported.
type SourceModule struct {
	Name   string
	Type   ModuleType
	Source string
}

// GenerateVBAProject returns a regenerated vbaProject.bin based on template.
//
// The engine replaces only supplied standard, class, and unambiguous document
// module source. Project records, references, codepage, and opaque streams such
// as existing UserForm designer storages are carried through from the template.
// Unsupported content returns one of the exported sentinel errors so the CLI can
// map it to the pack error contract.
func GenerateVBAProject(template []byte, sources []SourceModule) ([]byte, error) {
	if hasSignatureStream(template) {
		return nil, ErrSignedProject
	}
	project, err := vbaproject.Read(template)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	if project.Protection.IsProtected {
		return nil, ErrProtectedProject
	}
	if err := applySources(project, sources); err != nil {
		return nil, err
	}
	out, err := vbaproject.Write(project)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	return out, nil
}

func applySources(project *vbaproject.Project, sources []SourceModule) error {
	byKey := make(map[string]int, len(project.Modules))
	for i, m := range project.Modules {
		key := moduleKey(m.Name, fromProjectModuleType(m.Type))
		if _, exists := byKey[key]; exists {
			return fmt.Errorf("%w: duplicate template module %s", ErrAmbiguousLayout, m.Name)
		}
		byKey[key] = i
	}

	seen := make(map[string]bool, len(sources))
	for _, source := range sources {
		if source.Type == ModuleTypeForm {
			return ErrUserFormGenerationUnsupported
		}
		if source.Name == "" {
			return fmt.Errorf("%w: source module name is empty", ErrAmbiguousLayout)
		}
		targetType, err := toProjectModuleType(source.Type)
		if err != nil {
			return err
		}
		key := moduleKey(source.Name, source.Type)
		if seen[key] {
			return fmt.Errorf("%w: duplicate source module %s", ErrAmbiguousLayout, source.Name)
		}
		seen[key] = true
		idx, ok := byKey[key]
		if !ok {
			return fmt.Errorf("%w: source module %q is not in the template; pack updates existing modules only and cannot add new modules in the experimental MVP", ErrAmbiguousLayout, source.Name)
		}
		normalized, err := vbaproject.NormalizeModuleSource(targetType, source.Source, &project.Modules[idx])
		if err != nil {
			return fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		project.Modules[idx].Source = normalized
	}
	return nil
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
		return vbaproject.ModuleForm, ErrUserFormGenerationUnsupported
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
