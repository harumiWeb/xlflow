package excel

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
)

const vbaProjectEntry = "xl/vbaProject.bin"

func discoverPushCompilerInputs(root string, cfg config.Config) ([]sourceinventory.Component, []string, error) {
	components, err := sourceinventory.Discover(sourceinventory.Options{
		Root:                                  root,
		Config:                                cfg,
		ValidateFormArtifacts:                 true,
		CanonicalFormSpecs:                    true,
		AllowLegacyFormArtifacts:              true,
		IgnoreCanonicalCompatibilityArtifacts: true,
		PreserveMissingFormCode:               true,
		AllowLooseFormModules:                 true,
		AllowMissingRoots:                     true,
	})
	if err != nil {
		return nil, nil, err
	}
	var names []string
	for _, component := range components {
		if component.Type == sourceinventory.ComponentForm && component.FormSpec != nil {
			names = append(names, component.Name)
		}
	}
	slices.Sort(names)
	return components, names, nil
}

func pushPackSources(components []sourceinventory.Component, codeSource string) []pack.SourceModule {
	sources := make([]pack.SourceModule, 0, len(components))
	for _, component := range components {
		source := pack.SourceModule{
			SourcePath:   component.SourcePath,
			RelatedPaths: component.RelatedPaths(),
			Name:         component.Name,
			Source:       string(component.Source),
		}
		switch component.Type {
		case sourceinventory.ComponentStandard:
			source.Type = pack.ModuleTypeStandard
		case sourceinventory.ComponentClass:
			source.Type = pack.ModuleTypeClass
		case sourceinventory.ComponentDocument:
			source.Type = pack.ModuleTypeDocument
		case sourceinventory.ComponentForm:
			if component.FormSpec == nil {
				continue
			}
			source.Type = pack.ModuleTypeForm
			source.FormSpec = component.FormSpec
			source.FormSpecPath = component.SourcePath
			source.PreserveExistingCode = component.PreserveExistingCode
			for _, artifact := range component.Related {
				if artifact.Path == component.SourcePath {
					source.FormSpecSource = artifact.Source
					continue
				}
				if artifact.Role != sourceinventory.ArtifactRolePicture &&
					((strings.EqualFold(codeSource, "frm") && strings.EqualFold(filepath.Ext(artifact.Path), ".frm")) ||
						(!strings.EqualFold(codeSource, "frm") && strings.EqualFold(filepath.Ext(artifact.Path), ".bas") && strings.EqualFold(filepath.Base(artifact.Path), component.Name+".bas"))) {
					source.SourcePath = artifact.Path
					if !strings.EqualFold(codeSource, "frm") {
						source.Source = string(artifact.Source)
					}
				}
			}
		default:
			continue
		}
		sources = append(sources, source)
	}
	return sources
}

func compileTemplateWorkbook(template []byte, sources []pack.SourceModule) ([]byte, error) {
	compiled, _, err := pack.BuildWorkbook(template, sources)
	if err != nil {
		return nil, err
	}
	return compiled, nil
}

func compileFormWriteWorkbook(root string, cfg config.Config, template []byte, opts FormWriteOptions) ([]byte, error) {
	component, err := discoverFormWriteComponent(root, cfg, opts.SpecPath, opts.Spec.Form.Name)
	if err != nil {
		return nil, err
	}
	formSpec := component.FormSpec
	if formSpec == nil {
		return nil, fmt.Errorf("canonical UserForm spec %q was not found in source inventory", opts.Spec.Form.Name)
	}

	projectBytes, err := readWorkbookVBAProject(template)
	if err != nil {
		return nil, err
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		return nil, fmt.Errorf("parse staged workbook VBA project: %w", err)
	}

	formName := formSpec.Form.Name
	var existing *vbaproject.Module
	for i := range project.Modules {
		module := &project.Modules[i]
		if module.Type == vbaproject.ModuleForm && strings.EqualFold(module.Name, formName) {
			existing = module
			break
		}
	}
	if opts.Action == "apply" && existing == nil {
		return nil, fmt.Errorf("UserForm %q does not exist; form apply requires an existing form", formName)
	}
	if opts.Action == "build" && existing != nil && !opts.Overwrite {
		return nil, fmt.Errorf("UserForm %q already exists; pass --overwrite to replace it", formName)
	}

	sources := make([]pack.SourceModule, 0, len(project.Modules)+1)
	for _, module := range project.Modules {
		switch module.Type {
		case vbaproject.ModuleStd:
			sources = append(sources, pack.SourceModule{Name: module.Name, Type: pack.ModuleTypeStandard, Source: module.Source})
		case vbaproject.ModuleClass:
			sources = append(sources, pack.SourceModule{Name: module.Name, Type: pack.ModuleTypeClass, Source: module.Source})
		}
	}

	formSource := pushPackSources([]sourceinventory.Component{*component}, "sidecar")[0]
	formSource.PreserveExistingCode = existing != nil && component.PreserveExistingCode
	sources = append(sources, formSource)

	return compileTemplateWorkbook(template, sources)
}

func discoverFormWriteComponent(root string, cfg config.Config, specPath, formName string) (*sourceinventory.Component, error) {
	// FormWrite always treats a matching sidecar as authoritative code and
	// otherwise preserves the code already in the workbook. The inventory
	// resolver also loads picture bytes and records their roles for the compiler.
	cfg.UserForm.CodeSource = "sidecar"
	components, err := sourceinventory.Discover(sourceinventory.Options{
		Root:                                  root,
		Config:                                cfg,
		CanonicalFormSpecs:                    true,
		AllowLegacyFormArtifacts:              true,
		IgnoreCanonicalCompatibilityArtifacts: true,
		PreserveMissingFormCode:               true,
		ValidateFormArtifacts:                 true,
		AllowMissingRoots:                     true,
	})
	if err != nil {
		return nil, err
	}
	wantPath, err := resolveProjectSourcePath(root, specPath)
	if err != nil {
		return nil, fmt.Errorf("resolve UserForm spec path: %w", err)
	}
	for i := range components {
		component := &components[i]
		if component.Type != sourceinventory.ComponentForm || component.FormSpec == nil || !strings.EqualFold(component.Name, formName) {
			continue
		}
		gotPath, err := resolveProjectSourcePath(root, component.SourcePath)
		if err != nil {
			return nil, fmt.Errorf("resolve discovered UserForm spec path: %w", err)
		}
		if strings.EqualFold(filepath.Clean(gotPath), filepath.Clean(wantPath)) {
			return component, nil
		}
	}
	return nil, fmt.Errorf("canonical UserForm spec %q at %q was not found in source inventory", formName, specPath)
}

func resolveProjectSourcePath(root, path string) (string, error) {
	resolved := filepath.FromSlash(path)
	if !filepath.IsAbs(resolved) {
		resolved = filepath.Join(root, resolved)
	}
	return filepath.Abs(resolved)
}

func readWorkbookVBAProject(workbook []byte) ([]byte, error) {
	reader, err := zip.NewReader(bytes.NewReader(workbook), int64(len(workbook)))
	if err != nil {
		return nil, fmt.Errorf("read staged workbook package: %w", err)
	}
	for _, entry := range reader.File {
		if !strings.EqualFold(entry.Name, vbaProjectEntry) {
			continue
		}
		body, err := readZipEntry(entry)
		if err != nil {
			return nil, fmt.Errorf("read staged workbook VBA project: %w", err)
		}
		return body, nil
	}
	return nil, fmt.Errorf("staged workbook does not contain %s", vbaProjectEntry)
}

func readZipEntry(entry *zip.File) ([]byte, error) {
	reader, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer func() { _ = reader.Close() }()
	return io.ReadAll(reader)
}
