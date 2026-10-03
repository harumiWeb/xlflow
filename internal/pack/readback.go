package pack

import (
	"bytes"
	"fmt"
	"maps"
	"reflect"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

// validateProjectReadback binds writer output to the complete planned project.
// It checks Designer ownership as well as parseability before publication.
func validateProjectReadback(expected *vbaproject.Project, body []byte) error {
	actual, err := vbaproject.Read(body)
	if err != nil {
		return err
	}
	if len(expected.Modules) != len(actual.Modules) || len(expected.Forms) != len(actual.Forms) {
		return fmt.Errorf("component or Designer count differs from plan")
	}
	if !bytes.Equal(expected.ReferencesRaw, actual.ReferencesRaw) {
		return fmt.Errorf("reference records differ from plan")
	}
	components, err := ovba.ParseProjectText(actual.ProjectStreamRaw, actual.Props.CodePage)
	if err != nil {
		return err
	}
	if len(components.Components) != len(actual.Modules) {
		return fmt.Errorf("PROJECT declarations differ from module count")
	}
	var originalWM, actualWM []byte
	for path, value := range expected.RawStreams {
		if cfb.DirectoryNameKey(path) == cfb.DirectoryNameKey("PROJECTwm") {
			originalWM = value
		}
	}
	for path, value := range actual.RawStreams {
		if cfb.DirectoryNameKey(path) == cfb.DirectoryNameKey("PROJECTwm") {
			actualWM = value
		}
	}
	if !bytes.Equal(originalWM, actualWM) {
		specs := make([]ovba.ProjectComponentSpec, 0, len(expected.Modules))
		for _, module := range expected.Modules {
			specs = append(specs, ovba.ProjectComponentSpec{Kind: componentKind(module.Type), Name: module.Name})
		}
		mapping, err := ovba.BuildProjectWM(specs, expected.Props.CodePage)
		if err != nil {
			return err
		}
		if !bytes.Equal(mapping, actualWM) {
			return fmt.Errorf("rebuilt PROJECTwm differs from component plan")
		}
	}
	forms := make(map[string]*oforms.Form, len(actual.Forms))
	for _, form := range actual.Forms {
		if _, exists := forms[form.Name]; exists {
			return fmt.Errorf("duplicate Designer %q", form.Name)
		}
		forms[form.Name] = form
	}
	for i, module := range expected.Modules {
		got := actual.Modules[i]
		if module.Name != got.Name || module.StreamName != got.StreamName || module.Type != got.Type || module.Source != got.Source {
			return fmt.Errorf("module %q differs from plan", module.Name)
		}
		if !slices.ContainsFunc(components.Components, func(c ovba.ProjectComponent) bool { return c.Name == got.Name && componentKind(got.Type) == c.Kind }) {
			return fmt.Errorf("module %q has no matching PROJECT declaration", got.Name)
		}
		if got.Type == vbaproject.ModuleForm {
			if _, ok := forms[got.Name]; !ok {
				return fmt.Errorf("form module %q has no matching Designer", got.Name)
			}
			delete(forms, got.Name)
		}
	}
	if len(forms) != 0 {
		return fmt.Errorf("designer has no matching form module")
	}
	for _, form := range expected.Forms {
		index := slices.IndexFunc(actual.Forms, func(f *oforms.Form) bool { return f.Name == form.Name })
		if index < 0 {
			return fmt.Errorf("missing Designer %q", form.Name)
		}
		before, err := oforms.SerializeForm(form, expected.Props.CodePage)
		if err != nil {
			return err
		}
		after, err := oforms.SerializeForm(actual.Forms[index], actual.Props.CodePage)
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(before, after) {
			return fmt.Errorf("designer %q differs from plan", form.Name)
		}
	}
	// PROJECTwm is writer-owned when form topology changes; every other
	// opaque payload must be present exactly once, in both directions.
	for _, pair := range [][2]map[string][]byte{{expected.RawStreams, actual.RawStreams}, {actual.RawStreams, expected.RawStreams}} {
		for path, raw := range pair[0] {
			if cfb.DirectoryNameKey(path) == cfb.DirectoryNameKey("PROJECTwm") {
				continue
			}
			if got, ok := pair[1][path]; !ok || !bytes.Equal(raw, got) {
				return fmt.Errorf("opaque stream %q differs from plan", path)
			}
		}
	}
	metadata := maps.Clone(expected.StorageMetadata)
	if metadata == nil {
		metadata = make(map[string]cfb.StorageMeta)
	}
	// The writer creates these two structural storages even for a fresh project.
	for _, path := range []string{"", "VBA"} {
		if _, present := metadata[path]; !present {
			metadata[path] = cfb.StorageMeta{}
		}
	}
	if !reflect.DeepEqual(metadata, actual.StorageMetadata) {
		return fmt.Errorf("storage metadata differs from plan")
	}
	return nil
}

func countCarriedTemplateStreams(template, output []byte) (int, error) {
	before, err := cfb.Open(template)
	if err != nil {
		return 0, err
	}
	after, err := cfb.Open(output)
	if err != nil {
		return 0, err
	}
	count := 0
	for _, path := range before.Paths() {
		root := strings.SplitN(path, "/", 2)[0]
		if cfb.DirectoryNameKey(root) == cfb.DirectoryNameKey("VBA") || cfb.DirectoryNameKey(root) == cfb.DirectoryNameKey("PROJECT") {
			continue
		}
		old, exists := before.Stream(path)
		if !exists {
			continue
		}
		current, exists := after.Stream(path)
		if exists && bytes.Equal(old, current) {
			count++
		}
	}
	return count, nil
}

func componentKind(kind vbaproject.ModuleType) string {
	switch kind {
	case vbaproject.ModuleStd:
		return "Module"
	case vbaproject.ModuleClass:
		return "Class"
	case vbaproject.ModuleDocument:
		return "Document"
	case vbaproject.ModuleForm:
		return "BaseClass"
	default:
		return ""
	}
}
