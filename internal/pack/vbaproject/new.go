package vbaproject

import (
	"fmt"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

const (
	stdoleLibID = `*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation`
	officeLibID = `*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library`
)

// NewProjectSpec describes a fresh source-only VBA project.
type NewProjectSpec struct {
	Name     string
	CodePage uint16
	Modules  []Module
}

// NewProject constructs a deterministic VBA project without template bytes.
func NewProject(spec NewProjectSpec) (*Project, error) {
	info, err := ovba.BuildProjectInformation(ovba.ProjectInformationSpec{
		SysKind: 1, CodePage: spec.CodePage, Name: spec.Name,
	})
	if err != nil {
		return nil, err
	}
	references, err := ovba.BuildProjectReferences([]ovba.RegisteredReferenceSpec{
		{Name: "stdole", LibID: stdoleLibID},
		{Name: "Office", LibID: officeLibID},
	}, spec.CodePage)
	if err != nil {
		return nil, err
	}
	components := make([]ovba.ProjectComponentSpec, 0, len(spec.Modules))
	for _, module := range spec.Modules {
		kind := projectKind(module.Type)
		if kind == "" {
			return nil, fmt.Errorf("vbaproject: module %q has an unknown ModuleType %d", module.Name, module.Type)
		}
		components = append(components, ovba.ProjectComponentSpec{Kind: kind, Name: module.Name})
	}
	projectText, err := ovba.BuildProjectText(spec.Name, components, spec.CodePage)
	if err != nil {
		return nil, err
	}
	return &Project{
		CFBFormat:               cfb.FormatV3,
		Modules:                 spec.Modules,
		References:              []Reference{{Name: "stdole"}, {Name: "Office"}},
		ReferencesRaw:           references,
		ProjectInfoRaw:          info,
		ProjectStreamRaw:        projectText,
		GenerateProjectMetadata: true,
		RawStreams:              map[string][]byte{},
		StorageMetadata:         map[string]cfb.StorageMeta{},
		Props:                   ProjectProps{Name: spec.Name, SysKind: 1, LCID: ovba.CanonicalProjectLCID, CodePage: spec.CodePage},
		Protection:              Protection{CMG: "7577CB4035B139B139B139B139", DPB: "EAE854D3C8D4C8D4C8", GC: "5F5DE16E23969997999766"},
	}, nil
}
