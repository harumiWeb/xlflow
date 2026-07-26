package output

import (
	"encoding/json"
	"fmt"
	"strings"
)

// BuildTemporaryCleanupFailedCode is emitted by the Excel bridge when the
// artifact was published but its bridge-owned staging directory remains.
const BuildTemporaryCleanupFailedCode = "build_temporary_cleanup_failed"

func (r renderer) renderBuild(env Envelope) string {
	build := objectMap(env.Build)
	outputPayload := objectMap(env.Output)
	if len(build) == 0 && len(outputPayload) == 0 {
		return ""
	}
	var b strings.Builder
	if base := stringValue(build, "base"); base != "" {
		fmt.Fprintf(&b, "Base:\n  %s\n\n", base)
	}
	if outputPath := stringValue(build, "output"); outputPath != "" {
		fmt.Fprintf(&b, "Output:\n  %s\n\n", outputPath)
	}
	if outputPath := stringValue(outputPayload, "path"); outputPath != "" {
		fmt.Fprintf(&b, "Published output:\n  %s\n\n", outputPath)
	}
	if publication := stringValue(outputPayload, "publication"); publication != "" {
		fmt.Fprintf(&b, "Publication:\n  %s\n\n", publication)
	}
	if replaced, ok := boolValueOK(outputPayload, "replaced_existing"); ok {
		fmt.Fprintf(&b, "Replaced existing:\n  %t\n\n", replaced)
	}
	if cleanup := objectMap(outputPayload["temporary_cleanup"]); len(cleanup) > 0 {
		if status := stringValue(cleanup, "status"); status != "" {
			fmt.Fprintf(&b, "Temporary cleanup:\n  %s\n\n", status)
		}
		if residualPath := stringValue(cleanup, "residual_path"); residualPath != "" {
			fmt.Fprintf(&b, "Temporary residual:\n  %s\n\n", residualPath)
		}
	} else if cleanup := stringValue(outputPayload, "temporary_cleanup"); cleanup != "" {
		// Accept the compact string representation used by early bridge builds.
		fmt.Fprintf(&b, "Temporary cleanup:\n  %s\n\n", cleanup)
	}
	if manifest := objectMap(build["manifest"]); len(manifest) > 0 {
		if path := stringValue(manifest, "path"); path != "" {
			fmt.Fprintf(&b, "Manifest:\n  %s\n\n", path)
		}
		if published, ok := boolValueOK(manifest, "published"); ok {
			fmt.Fprintf(&b, "Manifest published:\n  %t\n\n", published)
		}
		if message := stringValue(manifest, "error"); message != "" {
			fmt.Fprintf(&b, "Manifest error:\n  %s\n\n", message)
		}
	}
	included := build["included_components"]
	if included == nil {
		included = build["included"]
	}
	excluded := build["excluded_components"]
	if excluded == nil {
		excluded = build["excluded"]
	}
	r.renderBuildComponents(&b, "Included", included)
	r.renderBuildComponents(&b, "Excluded", excluded)
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
