package output

import (
	"fmt"
	"strings"
)

// Pack diagnostics include source locations and copyable configuration examples.
// Keep them outside a table so multiline suggestions retain their line breaks.
func (r renderer) renderPackIssues(env Envelope) string {
	issues := listOfObjects(env.Issues)
	if len(issues) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString(r.section("Pack input issues"))
	for _, issue := range issues {
		fmt.Fprintf(&b, "%s [%s] %s\n", r.severityBadge(stringValue(issue, "severity")), stringValue(issue, "code"), lintIssueLocation(issue))
		fmt.Fprintln(&b, stringValue(issue, "message"))
		if suggestion := stringValue(issue, "suggestion"); suggestion != "" {
			fmt.Fprintf(&b, "Suggestion: %s\n", suggestion)
		}
		b.WriteByte('\n')
	}
	return b.String()
}
