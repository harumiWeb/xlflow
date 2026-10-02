// Package pack implements xlflow's pure-Go vbaProject.bin generation engine.
//
// Reference implementation: https://github.com/kay-ws/ovba-writer
package pack

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"github.com/xuri/excelize/v2"
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

	// ModuleTypeForm updates existing template code or creates a blank-mode
	// Designer from canonical FormSpec.
	ModuleTypeForm ModuleType = "form"
)

// SourceModule is one source-tree module to apply to a template vbaProject.bin.
//
// Source is the disk-form text read from a .bas/.cls/.frm file. The complete set of supplied standard
// and class modules is authoritative: missing template modules are removed and new modules are added.
// Document modules are replaced by exact module name match against an existing template document
// module. A UserForm's code-behind is replaced when the form already exists in the template; its
// designer layout is preserved. Blank mode creates new Designers from FormSpec
// and interprets Source as code-only text for those forms.
type SourceModule struct {
	SourcePath   string
	RelatedPaths []string
	Name         string
	Type         ModuleType
	Source       string
	// FormSpec is authoring intent used only for new blank-mode Designers.
	FormSpec *spec.FormSpec
}

// PackMeta summarizes the modules and opaque streams handled during pack.
type PackMeta struct {
	Standard       int
	Class          int
	Document       int
	Form           int
	CarriedStreams int
}

// BlankOptions configures text encoding in a fresh VBA project.
type BlankOptions struct {
	CodePage uint16
}

const DefaultBlankCodePage uint16 = 1252

// BuildBlankWorkbook creates a one-sheet macro-enabled workbook without using
// an existing workbook as a template. Its document topology is fixed to
// ThisWorkbook and Sheet1; all other components are source-authoritative.
func BuildBlankWorkbook(sources []SourceModule, opts BlankOptions) ([]byte, PackMeta, error) {
	if opts.CodePage == 0 {
		opts.CodePage = DefaultBlankCodePage
	}
	seenDocuments := map[string]bool{}
	var codeSources, formSources []SourceModule
	seenNames := map[string]bool{}
	for _, source := range sources {
		key := strings.ToLower(source.Name)
		if seenNames[key] {
			return nil, PackMeta{}, fmt.Errorf("%w: duplicate source component %s", ErrAmbiguousLayout, source.Name)
		}
		seenNames[key] = true
		switch source.Type {
		case ModuleTypeForm:
			if source.FormSpec == nil {
				return nil, PackMeta{}, fmt.Errorf("%w: canonical spec required for %s", ErrBlankUserFormUnsupported, sourceLabel(source))
			}
			if source.FormSpec.Form.Name != source.Name {
				return nil, PackMeta{}, fmt.Errorf("%w: spec/module identity mismatch for %s", ErrAmbiguousLayout, sourceLabel(source))
			}
			formSources = append(formSources, source)
		case ModuleTypeDocument:
			if source.Name != "ThisWorkbook" && source.Name != "Sheet1" {
				return nil, PackMeta{}, fmt.Errorf("%w: blank mode only supports document modules ThisWorkbook and Sheet1, got %q", ErrAmbiguousLayout, source.Name)
			}
			seenDocuments[source.Name] = true
		}
		if source.Type != ModuleTypeForm {
			codeSources = append(codeSources, source)
		}
	}
	for _, name := range []string{"ThisWorkbook", "Sheet1"} {
		if !seenDocuments[name] {
			return nil, PackMeta{}, fmt.Errorf("%w: blank mode requires document module %s", ErrAmbiguousLayout, name)
		}
	}
	workbookModule, err := vbaproject.NewDocumentModule("ThisWorkbook", "", false)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	sheetModule, err := vbaproject.NewDocumentModule("Sheet1", "", true)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	project, err := vbaproject.NewProject(vbaproject.NewProjectSpec{
		Name: "VBAProject", CodePage: opts.CodePage,
		Modules: []vbaproject.Module{workbookModule, sheetModule},
	})
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	meta, err := applySources(project, codeSources)
	if err != nil {
		return nil, PackMeta{}, err
	}
	slices.SortFunc(formSources, func(a, b SourceModule) int { return strings.Compare(a.Name, b.Name) })
	for _, source := range formSources {
		form, err := compiler.CompileNew(*source.FormSpec, project.Props.CodePage)
		if err != nil {
			category := ErrAmbiguousLayout
			if detail, ok := errors.AsType[*compiler.Error](err); ok && detail.Code == compiler.GenerationUnsupported {
				category = ErrUserFormGenerationUnsupported
			}
			return nil, PackMeta{}, fmt.Errorf("%w: %s: %v", category, sourceLabel(source), err)
		}
		module, err := vbaproject.NewUserFormModule(form, source.Source, project.Props.CodePage)
		if err != nil {
			return nil, PackMeta{}, fmt.Errorf("%w: %s: %v", ErrAmbiguousLayout, sourceLabel(source), err)
		}
		project, err = vbaproject.WithNewUserForm(project, form, module)
		if err != nil {
			return nil, PackMeta{}, fmt.Errorf("%w: %s: %v", ErrAmbiguousLayout, sourceLabel(source), err)
		}
		meta.Form++
	}
	vbaProject, err := vbaproject.Write(project)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	if _, err := vbaproject.Read(vbaProject); err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: blank project readback: %v", ErrAmbiguousLayout, err)
	}

	workbook := excelize.NewFile()
	defer func() { _ = workbook.Close() }()
	// Excelize selects the macro-enabled OOXML content type from File.Path when
	// writing. BuildBlankWorkbook returns bytes rather than calling SaveAs, so
	// set a synthetic .xlsm path before writing to the buffer.
	workbook.Path = "Book.xlsm"
	workbookCodeName := "ThisWorkbook"
	sheetCodeName := "Sheet1"
	if err := workbook.SetWorkbookProps(&excelize.WorkbookPropsOptions{CodeName: &workbookCodeName}); err != nil {
		return nil, PackMeta{}, fmt.Errorf("pack: set workbook code name: %w", err)
	}
	if err := workbook.SetSheetProps("Sheet1", &excelize.SheetPropsOptions{CodeName: &sheetCodeName}); err != nil {
		return nil, PackMeta{}, fmt.Errorf("pack: set sheet code name: %w", err)
	}
	if err := workbook.AddVBAProject(vbaProject); err != nil {
		return nil, PackMeta{}, fmt.Errorf("pack: add VBA project: %w", err)
	}
	var buf bytes.Buffer
	if err := workbook.Write(&buf); err != nil {
		return nil, PackMeta{}, fmt.Errorf("pack: write blank workbook: %w", err)
	}
	return buf.Bytes(), meta, nil
}

// GenerateVBAProject returns a regenerated vbaProject.bin based on template.
//
// The engine reconstructs the standard and class component set from sources and
// replaces supplied unambiguous document and UserForm code. Project records,
// references, codepage, document/UserForm topology, and opaque streams such as
// existing UserForm designer storages are carried through from the template.
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
	if signed, err := hasPackageVBASignature(reader); err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	} else if signed {
		return nil, PackMeta{}, ErrSignedProject
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
	meta.CarriedStreams = project.CarriedStreamCount()
	out, err := vbaproject.Write(project)
	if err != nil {
		return nil, PackMeta{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	return out, meta, nil
}

func applySources(project *vbaproject.Project, sources []SourceModule) (PackMeta, error) {
	plan, err := PlanProject(project, sources)
	if err != nil {
		return PackMeta{}, err
	}
	project.Modules = plan.modules
	return plan.meta, nil
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
