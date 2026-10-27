package symbols

import "strings"

// IsPredeclaredClass identifies the module symbol of a class whose exported
// attributes declare a default instance. A class type name alone is not a
// project-visible assignment target.
func IsPredeclaredClass(symbol Symbol) bool {
	if symbol.Parent != "" || !strings.EqualFold(symbol.Kind, "module") || !strings.EqualFold(symbol.ModuleKind, "class") {
		return false
	}
	for _, attribute := range symbol.Attributes {
		if strings.EqualFold(attribute.Name, "VB_PredeclaredId") {
			return strings.EqualFold(strings.TrimSpace(attribute.Value), "true")
		}
	}
	return false
}
