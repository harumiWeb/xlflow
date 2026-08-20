package analyze

import "strings"

// declarationScope layers procedure-local declarations over immutable module
// declarations. Keeping the layers separate avoids copying a large module
// declaration map for every procedure while preserving ordinary map lookup
// semantics for callers that need the merged view.
type declarationScope struct {
	module     map[string]sourceDeclaration
	extra      map[string]sourceDeclaration
	local      map[string]sourceDeclaration
	parameters map[string]sourceDeclaration
}

func newDeclarationScope(file parsedFile, proc sourceProcedure) declarationScope {
	scope := declarationScope{
		module:     file.moduleDecls(),
		local:      file.procedureDeclarationsFor(proc),
		parameters: make(map[string]sourceDeclaration, len(proc.Params)),
	}
	for _, parameter := range proc.Params {
		name := strings.ToLower(strings.TrimSpace(parameter.Name))
		if name == "" {
			continue
		}
		scope.parameters[name] = sourceDeclaration{
			Name: parameter.Name, Type: parameter.Type, Line: proc.StartLine,
			Object: isObjectType(parameter.Type), Parameter: true,
		}
	}
	return scope
}

func (scope declarationScope) lookup(name string) (sourceDeclaration, bool) {
	key := strings.ToLower(strings.TrimSpace(name))
	if declaration, ok := scope.parameters[key]; ok {
		return declaration, true
	}
	if declaration, ok := scope.local[key]; ok {
		return declaration, true
	}
	if declaration, ok := scope.extra[key]; ok {
		return declaration, true
	}
	declaration, ok := scope.module[key]
	return declaration, ok
}

func (scope *declarationScope) addExtraIfMissing(name string, declaration sourceDeclaration) {
	if scope == nil {
		return
	}
	key := strings.ToLower(strings.TrimSpace(name))
	if key == "" {
		return
	}
	if _, exists := scope.lookup(key); exists {
		return
	}
	if scope.extra == nil {
		scope.extra = make(map[string]sourceDeclaration)
	}
	scope.extra[key] = declaration
}

// addProcedureIRDeclaration preserves procedure declarations recovered by the
// IR when the source scanner cannot represent them. Procedure declarations
// shadow module declarations, while parameters remain the highest-priority
// layer in lookup and iteration.
func (scope *declarationScope) addProcedureIRDeclaration(declaration sourceDeclaration) {
	if scope == nil {
		return
	}
	key := strings.ToLower(strings.TrimSpace(declaration.Name))
	if key == "" {
		return
	}
	if _, exists := scope.local[key]; exists {
		return
	}
	if scope.extra == nil {
		scope.extra = make(map[string]sourceDeclaration)
	}
	scope.extra[key] = declaration
}

// forEach visits the effective merged scope without allocating a merged map.
// Overlay order matches map construction in the former implementation:
// module, procedure declarations, IR-only extras, and parameters.
func (scope declarationScope) forEach(visit func(string, sourceDeclaration)) {
	if visit == nil {
		return
	}
	for key, declaration := range scope.module {
		if _, shadowed := scope.local[key]; shadowed {
			continue
		}
		if _, shadowed := scope.extra[key]; shadowed {
			continue
		}
		if _, shadowed := scope.parameters[key]; shadowed {
			continue
		}
		visit(key, declaration)
	}
	for key, declaration := range scope.extra {
		if _, shadowed := scope.local[key]; shadowed {
			continue
		}
		if _, shadowed := scope.parameters[key]; shadowed {
			continue
		}
		visit(key, declaration)
	}
	for key, declaration := range scope.local {
		if _, shadowed := scope.parameters[key]; shadowed {
			continue
		}
		visit(key, declaration)
	}
	for key, declaration := range scope.parameters {
		visit(key, declaration)
	}
}

func (scope declarationScope) len() int {
	count := len(scope.module) + len(scope.extra) + len(scope.local) + len(scope.parameters)
	return count
}
