package pack

import (
	"fmt"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
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
}

// PlanProject validates and reconciles sources against a parsed template
// without modifying either input.
func PlanProject(project *vbaproject.Project, sources []SourceModule) (PackPlan, error) {
	if project == nil {
		return PackPlan{}, fmt.Errorf("%w: template project is nil", ErrAmbiguousLayout)
	}
	templateByName := make(map[string]vbaproject.Module, len(project.Modules))
	for _, module := range project.Modules {
		if err := vbaproject.ValidateWritableComponentIdentity(module.Name, module.StreamName); err != nil {
			return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		key := strings.ToLower(module.Name)
		if prior, exists := templateByName[key]; exists {
			return PackPlan{}, fmt.Errorf("%w: duplicate template modules %s and %s", ErrAmbiguousLayout, prior.Name, module.Name)
		}
		templateByName[key] = module
	}

	sourceByName := make(map[string]SourceModule, len(sources))
	for _, source := range sources {
		if !sourceinventory.ValidComponentName(source.Name) {
			return PackPlan{}, fmt.Errorf("%w: invalid VBA component name %q", ErrAmbiguousLayout, source.Name)
		}
		if err := vbaproject.ValidateWritableComponentIdentity(source.Name, source.Name); err != nil {
			return PackPlan{}, fmt.Errorf("%w: %v", ErrAmbiguousLayout, err)
		}
		if _, err := toProjectModuleType(source.Type); err != nil {
			return PackPlan{}, err
		}
		key := strings.ToLower(source.Name)
		if prior, exists := sourceByName[key]; exists {
			return PackPlan{}, fmt.Errorf("%w: duplicate source modules %s and %s", ErrAmbiguousLayout, prior.Name, source.Name)
		}
		if template, exists := templateByName[key]; exists {
			if source.Name != template.Name {
				return PackPlan{}, fmt.Errorf("%w: source component name %q differs in case from template component %q", ErrAmbiguousLayout, source.Name, template.Name)
			}
			if source.Type != fromProjectModuleType(template.Type) {
				return PackPlan{}, fmt.Errorf("%w: source component %q has type %q but template type is %q", ErrAmbiguousLayout, source.Name, source.Type, fromProjectModuleType(template.Type))
			}
		}
		sourceByName[key] = source
	}

	plan := PackPlan{
		Components: make([]PlannedComponent, 0, len(project.Modules)+len(sources)),
		modules:    make([]vbaproject.Module, 0, len(project.Modules)+len(sources)),
	}
	consumed := make(map[string]bool, len(sources))
	for _, module := range project.Modules {
		typ := fromProjectModuleType(module.Type)
		key := strings.ToLower(module.Name)
		source, supplied := sourceByName[key]
		planned := PlannedComponent{Name: module.Name, Type: typ}
		setAuthorities(&planned)
		if !supplied {
			if typ == ModuleTypeStandard || typ == ModuleTypeClass {
				planned.Action = PlanRemove
				plan.Components = append(plan.Components, planned)
				continue
			}
			planned.Action = PlanPreserve
			plan.Components = append(plan.Components, planned)
			plan.modules = append(plan.modules, module)
			continue
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
		key := strings.ToLower(source.Name)
		if consumed[key] {
			continue
		}
		switch source.Type {
		case ModuleTypeForm:
			return PackPlan{}, fmt.Errorf("%w: form %q is not in the template; pack updates the code-behind of existing forms only and cannot create a new UserForm", ErrUserFormGenerationUnsupported, source.Name)
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
	return plan, nil
}

func normalizePlannedSource(module vbaproject.Module, source SourceModule) (string, error) {
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
