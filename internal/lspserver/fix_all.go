package lspserver

import "strings"

func sourceCodeActionKind(kind string) bool {
	return kind == "source" || strings.HasPrefix(kind, "source.")
}

func withoutSourceCodeActionKinds(kinds []string) []string {
	out := make([]string, 0, len(kinds))
	for _, kind := range kinds {
		if !sourceCodeActionKind(kind) {
			out = append(out, kind)
		}
	}
	return out
}
