package output

import (
	"encoding/json"
	"fmt"
	"strings"
)

func (r renderer) renderBuild(env Envelope) string {
	build := objectMap(env.Build)
	if len(build) == 0 {
		return ""
	}
	var b strings.Builder
	if base := stringValue(build, "base"); base != "" {
		fmt.Fprintf(&b, "Base:\n  %s\n\n", base)
	}
	if outputPath := stringValue(build, "output"); outputPath != "" {
		fmt.Fprintf(&b, "Output:\n  %s\n\n", outputPath)
	}
	r.renderBuildComponents(&b, "Included", build["included"])
	r.renderBuildComponents(&b, "Excluded", build["excluded"])
	if warnings := buildComponentMaps(build["warnings"]); len(warnings) > 0 {
		b.WriteString("Warnings:\n")
		for _, warning := range warnings {
			message := stringValue(warning, "message")
			if message != "" {
				fmt.Fprintf(&b, "  %s\n", message)
			}
		}
	}
	return b.String()
}

func (r renderer) renderBuildComponents(b *strings.Builder, title string, raw any) {
	components := buildComponentMaps(raw)
	if len(components) == 0 {
		return
	}
	fmt.Fprintf(b, "%s:\n", title)
	for _, component := range components {
		name := stringValue(component, "name")
		path := stringValue(component, "source_path")
		if path == "" {
			fmt.Fprintf(b, "  %s\n", name)
			continue
		}
		fmt.Fprintf(b, "  %s (%s)\n", name, path)
	}
	b.WriteString("\n")
}

func buildComponentMaps(raw any) []map[string]any {
	if raw == nil {
		return nil
	}
	body, err := json.Marshal(raw)
	if err != nil {
		return nil
	}
	var items []map[string]any
	if json.Unmarshal(body, &items) != nil {
		return nil
	}
	return items
}
