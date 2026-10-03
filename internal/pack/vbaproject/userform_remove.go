package vbaproject

import (
	"bytes"
	"fmt"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

// Clone owns mutable project data and independently reparses each Designer.
// It preserves the original writer bookkeeping without modifying the input.
func Clone(p *Project) (*Project, error) {
	if p == nil {
		return nil, fmt.Errorf("vbaproject: cannot clone nil project")
	}
	result := cloneProject(p)
	for i, form := range p.Forms {
		serialized, err := oforms.SerializeForm(form, p.Props.CodePage)
		if err != nil {
			return nil, err
		}
		result.Forms[i], err = cloneSerializedUserForm(serialized, form.Name, p.Props.CodePage)
		if err != nil {
			return nil, err
		}
	}
	return result, nil
}

// WithoutUserForm removes a module and its owned Designer atomically. Existing
// references remain available to other code, including after the final removal.
func WithoutUserForm(p *Project, name string) (*Project, error) {
	if p == nil || p.Protection.IsProtected {
		return nil, fmt.Errorf("vbaproject: cannot remove form from nil/protected project")
	}
	moduleIndex := slices.IndexFunc(p.Modules, func(m Module) bool { return m.Name == name && m.Type == ModuleForm })
	formIndex := slices.IndexFunc(p.Forms, func(f *oforms.Form) bool { return f != nil && f.Name == name })
	if moduleIndex < 0 || formIndex < 0 {
		return nil, fmt.Errorf("vbaproject: form %q has no matching module/Designer", name)
	}
	if _, err := oforms.SerializeForm(p.Forms[formIndex], p.Props.CodePage); err != nil {
		return nil, err
	}
	parsed, err := ovba.ParseProjectText(p.ProjectStreamRaw, p.Props.CodePage)
	if err != nil {
		return nil, err
	}
	if !slices.ContainsFunc(parsed.Components, func(c ovba.ProjectComponent) bool { return c.Kind == "BaseClass" && c.Name == name }) {
		return nil, fmt.Errorf("vbaproject: form %q has no BaseClass declaration", name)
	}
	var output bytes.Buffer
	workspace := false
	for _, line := range splitProjectLines(p.ProjectStreamRaw) {
		text, err := ovba.DecodeMBCS(line.body, p.Props.CodePage)
		if err != nil {
			return nil, err
		}
		trimmed := strings.TrimSpace(text)
		if strings.HasPrefix(trimmed, "[") {
			workspace = strings.EqualFold(trimmed, "[Workspace]")
		}
		key, value, pair := strings.Cut(text, "=")
		if pair && (key == "BaseClass" && strings.Trim(strings.TrimSpace(value), "\"") == name || workspace && strings.EqualFold(strings.TrimSpace(key), name)) {
			continue
		}
		output.Write(line.body)
		output.Write(line.eol)
	}
	result := cloneProject(p)
	result.Modules = slices.Delete(result.Modules, moduleIndex, moduleIndex+1)
	result.Forms = slices.Delete(result.Forms, formIndex, formIndex+1)
	result.ProjectStreamRaw = output.Bytes()
	result.rebuildProjectWM = true
	return result, nil
}
