package vbaproject

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

const formsLibID = `*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Windows\System32\FM20.DLL#Microsoft Forms 2.0 Object Library`

// EnsureMSFormsReference returns an independent project with a Forms reference.
// Unrelated reference records are retained verbatim, in their original order.
func EnsureMSFormsReference(p *Project) (*Project, error) {
	if p == nil || p.Protection.IsProtected {
		return nil, fmt.Errorf("vbaproject: cannot add Forms reference to nil/protected project")
	}
	for path := range p.RawStreams {
		key := strings.ToLower(path)
		if strings.Contains(key, "signature") || strings.Contains(key, "_vba_project_cur") {
			return nil, fmt.Errorf("vbaproject: signed project cannot add Forms reference")
		}
	}
	refs, err := ovba.ParseProjectReferences(p.ReferencesRaw, p.Props.CodePage)
	if err != nil {
		return nil, err
	}
	result := cloneProject(p)
	for _, ref := range refs {
		if ref.Kind == 0x000D && hasMSFormsLIBID([]byte(ref.LibID)) || ref.Kind == 0x002F && ref.OriginalTypeLib == userFormMSFormsReferenceGUIDWire {
			return result, nil
		}
	}
	addition, err := ovba.BuildProjectReferences([]ovba.RegisteredReferenceSpec{{Name: "MSForms", LibID: formsLibID}}, p.Props.CodePage)
	if err != nil {
		return nil, err
	}
	result.ReferencesRaw = append(result.ReferencesRaw, addition...)
	result.References = append(result.References, Reference{Name: "MSForms"})
	if _, err := ovba.ParseProjectReferences(result.ReferencesRaw, p.Props.CodePage); err != nil {
		return nil, err
	}
	return result, nil
}
