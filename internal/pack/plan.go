package pack

import (
	"fmt"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

type PlanAction string

const (
	PlanPreserve PlanAction = "preserve"
	PlanUpdate   PlanAction = "update"
	PlanRemove   PlanAction = "remove"
	PlanAdd      PlanAction = "add"
)

type Authority string

const (
	AuthoritySource   Authority = "source"
	AuthorityTemplate Authority = "template"
)

// PlannedComponent records one deterministic component decision. SourcePath
// and RelatedPaths identify the exact source artifacts used to make it.
type PlannedComponent struct {
	SourcePath        string
	RelatedPaths      []string
	Name              string
	Type              ModuleType
	TopologyAuthority Authority
	CodeAuthority     Authority
	Action            PlanAction
}

// PackPlan is a read-only reconciliation of source inventory and template
// topology. The private module slice contains fully normalized writer inputs,
// so applying a successful plan performs no further validation or decisions.
type PackPlan struct {
	Components []PlannedComponent
	modules    []vbaproject.Module
	meta       PackMeta
	project    *vbaproject.Project
}

// TemplateOptions controls whether omitted template forms survive packing.
// Empty UserFormTopology retains the backwards-compatible template authority.
type TemplateOptions struct {
	UserFormTopology string
}

// PlanProject validates and reconciles sources against a parsed template
// without modifying either input.
func PlanProject(project *vbaproject.Project, sources []SourceModule, options ...TemplateOptions) (PackPlan, error) {
	if project == nil {
		return PackPlan{}, fmt.Errorf("%w: template project is nil", ErrAmbiguousLayout)
	}
	if len(options) > 1 {
		return PackPlan{}, fmt.Errorf("%w: multiple template options", ErrAmbiguousLayout)
	}
	formAuthority := AuthorityTemplate
	if len(options) == 1 {
		switch options[0].UserFormTopology {
		case "", "template":
		case "source":
			formAuthority = AuthoritySource
		default:
			return PackPlan{}, fmt.Errorf("%w: invalid UserForm topology %q", ErrAmbiguousLayout, options[0].UserFormTopology)
		}
	}
	// Mutations below operate on an independent project, and the caller never
	// sees a partially reconciled Designer/reference/component state.
	working, err := vbaproject.Clone(project)
	if err != nil {
		return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
	}
	templateByName := make(map[string]vbaproject.Module, len(project.Modules))
	for _, module := range project.Modules {
		if err := vbaproject.ValidateWritableComponentIdentity(module.Name, module.StreamName, project.Props.CodePage); err != nil {
			return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		key := strings.ToLower(module.Name)
		if prior, exists := templateByName[key]; exists {
			return PackPlan{}, fmt.Errorf("%w: duplicate template modules %s and %s", ErrAmbiguousLayout, prior.Name, module.Name)
		}
		templateByName[key] = module
	}

	sourceByName := make(map[string]SourceModule, len(sources))
	sourceByKey := make(map[string]SourceModule, len(sources))
	for _, source := range sources {
		if source.FormSpec != nil && (source.Type != ModuleTypeForm || source.FormSpec.Form.Name != source.Name) {
			return PackPlan{}, &compiler.Error{Code: compiler.GenerationInvalid, Form: source.Name, Reason: "canonical spec identity differs from source component"}
		}
		if source.Type == ModuleTypeForm && formAuthority == AuthoritySource && source.FormSpec == nil {
			return PackPlan{}, &compiler.Error{Code: compiler.GenerationUnsupported, Form: source.Name, Reason: "source UserForm topology requires a canonical spec"}
		}
		if !sourceinventory.ValidComponentName(source.Name) {
			return PackPlan{}, fmt.Errorf("%w: invalid VBA component name %q", ErrAmbiguousLayout, source.Name)
		}
		if err := vbaproject.ValidateWritableComponentIdentity(source.Name, source.Name, project.Props.CodePage); err != nil {
			return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		if _, err := toProjectModuleType(source.Type); err != nil {
			return PackPlan{}, err
		}
		key := strings.ToLower(source.Name)
		if prior, exists := sourceByName[key]; exists {
			return PackPlan{}, fmt.Errorf("%w: duplicate source modules %s and %s", ErrAmbiguousLayout, prior.Name, source.Name)
		}
		sourceByName[key] = source
		sourceByKey[moduleKey(source.Name, source.Type)] = source
		if template, exists := templateByName[key]; exists {
			sourceOwned := source.Type == ModuleTypeStandard || source.Type == ModuleTypeClass
			templateOwned := template.Type == vbaproject.ModuleStd || template.Type == vbaproject.ModuleClass
			if !sourceOwned || !templateOwned {
				// Template-owned document/UserForm components must match the
				// template component by exact name and type; a case-only rename or a
				// type change is ambiguous, not a source-owned replacement.
				if source.Name != template.Name {
					return PackPlan{}, fmt.Errorf("%w: source component name %q differs in case from template component %q", ErrAmbiguousLayout, source.Name, template.Name)
				}
				if source.Type != fromProjectModuleType(template.Type) {
					return PackPlan{}, fmt.Errorf("%w: source component %q has type %q but template type is %q", ErrAmbiguousLayout, source.Name, source.Type, fromProjectModuleType(template.Type))
				}
			}
		}
	}

	plan := PackPlan{
		Components: make([]PlannedComponent, 0, len(project.Modules)+len(sources)),
		modules:    make([]vbaproject.Module, 0, len(project.Modules)+len(sources)),
	}
	consumed := make(map[string]bool, len(sources))
	for _, module := range project.Modules {
		typ := fromProjectModuleType(module.Type)
		key := moduleKey(module.Name, typ)
		source, supplied := sourceByKey[key]
		planned := PlannedComponent{Name: module.Name, Type: typ}
		setAuthorities(&planned)
		if typ == ModuleTypeForm {
			planned.TopologyAuthority = formAuthority
		}
		if !supplied {
			if typ == ModuleTypeStandard || typ == ModuleTypeClass || typ == ModuleTypeForm && formAuthority == AuthoritySource {
				if typ == ModuleTypeForm {
					working, err = vbaproject.WithoutUserForm(working, module.Name)
					if err != nil {
						return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
					}
				}
				planned.Action = PlanRemove
				plan.Components = append(plan.Components, planned)
				continue
			}
			planned.Action = PlanPreserve
			plan.Components = append(plan.Components, planned)
			plan.modules = append(plan.modules, module)
			continue
		}
		if typ == ModuleTypeForm && source.FormSpec != nil {
			index := slices.IndexFunc(working.Forms, func(f *oforms.Form) bool { return f.Name == module.Name })
			if index < 0 {
				return PackPlan{}, fmt.Errorf("%w: form %q has no Designer", ErrAmbiguousLayout, module.Name)
			}
			updated, err := compiler.CompileTemplate(working.Forms[index], *source.FormSpec, project.Props.CodePage)
			if err != nil {
				return PackPlan{}, fmt.Errorf("%s: %w", sourceLabel(source), err)
			}
			working.Forms[index] = updated
		}

		normalized, err := normalizePlannedSource(module, source)
		if err != nil {
			return PackPlan{}, err
		}
		module.Source = normalized
		planned.Action = PlanUpdate
		planned.SourcePath = source.SourcePath
		planned.RelatedPaths = slices.Clone(source.RelatedPaths)
		plan.Components = append(plan.Components, planned)
		plan.modules = append(plan.modules, module)
		consumed[key] = true
		incrementMeta(&plan.meta, source.Type)
	}

	additions := make([]SourceModule, 0)
	for _, source := range sources {
		key := moduleKey(source.Name, source.Type)
		if consumed[key] {
			continue
		}
		switch source.Type {
		case ModuleTypeForm:
			if source.FormSpec == nil {
				return PackPlan{}, &compiler.Error{Code: compiler.GenerationUnsupported, Form: source.Name, Reason: "new UserForms require a canonical spec"}
			}
			additions = append(additions, source)
		case ModuleTypeDocument:
			return PackPlan{}, fmt.Errorf("%w: document module %q is not in the template; document topology is template-owned", ErrAmbiguousLayout, source.Name)
		case ModuleTypeStandard, ModuleTypeClass:
			additions = append(additions, source)
		}
	}
	slices.SortFunc(additions, func(a, b SourceModule) int {
		if cmp := strings.Compare(string(a.Type), string(b.Type)); cmp != 0 {
			return cmp
		}
		if cmp := strings.Compare(strings.ToLower(a.Name), strings.ToLower(b.Name)); cmp != 0 {
			return cmp
		}
		return strings.Compare(a.SourcePath, b.SourcePath)
	})
	for _, source := range additions {
		if source.Type == ModuleTypeForm {
			form, err := compiler.CompileNew(*source.FormSpec, project.Props.CodePage)
			if err != nil {
				return PackPlan{}, fmt.Errorf("%s: %w", sourceLabel(source), err)
			}
			module, err := vbaproject.NewUserFormModule(form, source.Source, project.Props.CodePage)
			if err != nil {
				return PackPlan{}, fmt.Errorf("%s: %w", sourceLabel(source), err)
			}
			working, err = vbaproject.WithNewUserForm(working, form, module)
			if err != nil {
				return PackPlan{}, fmt.Errorf("%s: %w", sourceLabel(source), err)
			}
			plan.modules = append(plan.modules, module)
			plan.Components = append(plan.Components, PlannedComponent{SourcePath: source.SourcePath, RelatedPaths: slices.Clone(source.RelatedPaths), Name: source.Name, Type: source.Type, Action: PlanAdd, TopologyAuthority: formAuthority, CodeAuthority: AuthoritySource})
			incrementMeta(&plan.meta, source.Type)
			continue
		}
		targetType, _ := toProjectModuleType(source.Type)
		normalized, err := normalizePlannedSource(vbaproject.Module{Name: source.Name, StreamName: source.Name, Type: targetType}, source)
		if err != nil {
			return PackPlan{}, err
		}
		plan.modules = append(plan.modules, vbaproject.Module{Name: source.Name, StreamName: source.Name, Type: targetType, Source: normalized})
		planned := PlannedComponent{
			SourcePath: source.SourcePath, RelatedPaths: slices.Clone(source.RelatedPaths),
			Name: source.Name, Type: source.Type, Action: PlanAdd,
		}
		setAuthorities(&planned)
		plan.Components = append(plan.Components, planned)
		incrementMeta(&plan.meta, source.Type)
	}
	// Stream names are CFB directory entries under the VBA storage, so their
	// uniqueness is governed by the [MS-CFB] case-insensitive comparison, not
	// by Unicode lowercase. Reject collisions (e.g. dotted/dotless I in a
	// CP1254 project) before an ambiguous container could be assembled.
	streams := make(map[string]string, len(plan.modules))
	for _, module := range plan.modules {
		key := cfb.DirectoryNameKey(module.StreamName)
		if prior, exists := streams[key]; exists {
			return PackPlan{}, fmt.Errorf("%w: module stream names %s and %s collide in the CFB directory", ErrAmbiguousLayout, prior, module.StreamName)
		}
		streams[key] = module.StreamName
	}
	working.Modules = slices.Clone(plan.modules)
	plan.project = working
	return plan, nil
}

func moduleKey(name string, typ ModuleType) string {
	return string(typ) + "\x00" + name
}

func normalizePlannedSource(module vbaproject.Module, source SourceModule) (string, error) {
	if source.Type == ModuleTypeForm && source.FormSpec != nil {
		if hasAttribute(source.Source) {
			return "", &compiler.Error{Code: compiler.GenerationInvalid, Form: source.Name, Reason: "canonical form code must not contain Attribute headers"}
		}
		// NormalizeModuleSource keeps all of the existing component attributes.
		// The synthetic disk header is only an in-memory code boundary.
		source.Source = "Attribute VB_Name = \"" + source.Name + "\"\r\n" + source.Source
	}
	if source.Type == ModuleTypeDocument && hasAttribute(source.Source) {
		return "", fmt.Errorf("%w: %s: document source must not contain Attribute headers", ErrAmbiguousLayout, sourceLabel(source))
	}
	if source.Type == ModuleTypeForm || hasVBName(source.Source) {
		if err := vbaproject.ValidateModuleIdentity(source.Name, source.Source); err != nil {
			return "", fmt.Errorf("%w: %s: %v", ErrAmbiguousLayout, sourceLabel(source), err)
		}
	}
	normalized, err := vbaproject.NormalizeModuleSource(module.Type, source.Source, &module)
	if err != nil {
		return "", fmt.Errorf("%w: %s: %v", ErrAmbiguousLayout, sourceLabel(source), err)
	}
	if source.Type == ModuleTypeStandard || source.Type == ModuleTypeClass {
		if err := vbaproject.ValidateModuleIdentity(source.Name, normalized); err != nil {
			return "", fmt.Errorf("%w: %s: %v", ErrAmbiguousLayout, sourceLabel(source), err)
		}
	}
	return normalized, nil
}

func hasAttribute(source string) bool {
	for line := range strings.SplitSeq(strings.ReplaceAll(source, "\r\n", "\n"), "\n") {
		if strings.HasPrefix(strings.TrimSpace(line), "Attribute ") {
			return true
		}
	}
	return false
}

func hasVBName(source string) bool {
	for line := range strings.SplitSeq(strings.ReplaceAll(source, "\r\n", "\n"), "\n") {
		if strings.HasPrefix(line, "Attribute VB_Name ") {
			return true
		}
	}
	return false
}

func sourceLabel(source SourceModule) string {
	if source.SourcePath != "" {
		return source.SourcePath
	}
	return source.Name
}

func setAuthorities(component *PlannedComponent) {
	switch component.Type {
	case ModuleTypeStandard, ModuleTypeClass:
		component.TopologyAuthority = AuthoritySource
		component.CodeAuthority = AuthoritySource
	case ModuleTypeDocument:
		component.TopologyAuthority = AuthorityTemplate
		component.CodeAuthority = AuthoritySource
	case ModuleTypeForm:
		component.TopologyAuthority = AuthorityTemplate
		component.CodeAuthority = AuthoritySource
	}
}

func incrementMeta(meta *PackMeta, typ ModuleType) {
	switch typ {
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
