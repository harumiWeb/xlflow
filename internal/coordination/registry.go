package coordination

import (
	"fmt"
	"sort"
	"strings"
)

type CLISelector struct {
	Path string `json:"path"`
}

// BridgeSelector matches a bridge command and, where required, selected
// payload fields. Argument names and values are compared case-insensitively.
type BridgeSelector struct {
	Command string            `json:"command"`
	Args    map[string]string `json:"args,omitempty"`
}

type Descriptor struct {
	ID     CommandID        `json:"id"`
	Policy Policy           `json:"policy"`
	CLI    []CLISelector    `json:"cli,omitempty"`
	Bridge []BridgeSelector `json:"bridge,omitempty"`
}

var descriptors = buildDescriptors()

// Lookup returns a defensive copy of the descriptor with id.
func Lookup(id CommandID) (Descriptor, error) {
	for _, descriptor := range descriptors {
		if descriptor.ID == id {
			return cloneDescriptor(descriptor), nil
		}
	}
	return Descriptor{}, &MissingPolicyError{Selector: "command id", Value: string(id)}
}

// LookupCLI resolves a Cobra command path. Both "xlflow push" and "push" are
// accepted so callers do not need to special-case the root command name.
func LookupCLI(path string) (Descriptor, error) {
	normalized := normalizeCLIPath(path)
	for _, descriptor := range descriptors {
		for _, selector := range descriptor.CLI {
			if normalizeCLIPath(selector.Path) == normalized {
				return cloneDescriptor(descriptor), nil
			}
		}
	}
	return Descriptor{}, &MissingPolicyError{Selector: "CLI command", Value: path}
}

// LookupBridge resolves a bridge request. Selectors with payload constraints
// take precedence over command-only selectors. Unknown actions therefore fail
// closed rather than inheriting the policy of a sibling action.
func LookupBridge(command string, args map[string]string) (Descriptor, error) {
	command = strings.TrimSpace(command)
	bestSpecificity := -1
	var best *Descriptor
	for i := range descriptors {
		descriptor := &descriptors[i]
		for _, selector := range descriptor.Bridge {
			if !strings.EqualFold(strings.TrimSpace(selector.Command), command) || !bridgeArgsMatch(selector.Args, args) {
				continue
			}
			if len(selector.Args) > bestSpecificity {
				bestSpecificity = len(selector.Args)
				best = descriptor
			}
		}
	}
	if best == nil {
		return Descriptor{}, &MissingPolicyError{Selector: "bridge command", Value: bridgeDisplayValue(command, args)}
	}
	return cloneDescriptor(*best), nil
}

// All returns descriptors in stable ID order. The returned descriptors,
// slices, and maps can be mutated without changing the registry.
func All() []Descriptor {
	result := make([]Descriptor, len(descriptors))
	for i, descriptor := range descriptors {
		result[i] = cloneDescriptor(descriptor)
	}
	sort.Slice(result, func(i, j int) bool { return result[i].ID < result[j].ID })
	return result
}

// ValidateRegistry checks the authoritative built-in registry for invalid or
// ambiguous selectors. It is exported for coverage and capability tests.
func ValidateRegistry() error { return validateDescriptors(descriptors) }

func normalizeCLIPath(path string) string {
	fields := strings.Fields(strings.TrimSpace(path))
	if len(fields) > 0 && strings.EqualFold(fields[0], "xlflow") {
		fields = fields[1:]
	}
	return strings.ToLower(strings.Join(fields, " "))
}

func bridgeArgsMatch(want, got map[string]string) bool {
	for wantKey, wantValue := range want {
		matched := false
		for gotKey, gotValue := range got {
			if strings.EqualFold(strings.TrimSpace(gotKey), strings.TrimSpace(wantKey)) &&
				strings.EqualFold(strings.TrimSpace(gotValue), strings.TrimSpace(wantValue)) {
				matched = true
				break
			}
		}
		if !matched {
			return false
		}
	}
	return true
}

func bridgeDisplayValue(command string, args map[string]string) string {
	if len(args) == 0 {
		return command
	}
	keys := make([]string, 0, len(args))
	for key := range args {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	parts := make([]string, 0, len(keys)+1)
	parts = append(parts, command)
	for _, key := range keys {
		parts = append(parts, fmt.Sprintf("%s=%s", key, args[key]))
	}
	return strings.Join(parts, " ")
}

func cloneDescriptor(source Descriptor) Descriptor {
	cloned := source
	cloned.CLI = append([]CLISelector(nil), source.CLI...)
	cloned.Bridge = make([]BridgeSelector, len(source.Bridge))
	for i, selector := range source.Bridge {
		cloned.Bridge[i] = BridgeSelector{Command: selector.Command}
		if selector.Args != nil {
			cloned.Bridge[i].Args = make(map[string]string, len(selector.Args))
			for key, value := range selector.Args {
				cloned.Bridge[i].Args[key] = value
			}
		}
	}
	return cloned
}

func validateDescriptors(values []Descriptor) error {
	ids := map[CommandID]struct{}{}
	cli := map[string]CommandID{}
	bridge := map[string]CommandID{}
	for _, descriptor := range values {
		if strings.TrimSpace(string(descriptor.ID)) == "" {
			return fmt.Errorf("coordination descriptor has an empty id")
		}
		if _, exists := ids[descriptor.ID]; exists {
			return fmt.Errorf("duplicate coordination command id %q", descriptor.ID)
		}
		ids[descriptor.ID] = struct{}{}
		if err := descriptor.Policy.Validate(); err != nil {
			return fmt.Errorf("coordination command %q: %w", descriptor.ID, err)
		}
		if len(descriptor.CLI) == 0 && len(descriptor.Bridge) == 0 {
			return fmt.Errorf("coordination command %q has no selector", descriptor.ID)
		}
		for _, selector := range descriptor.CLI {
			key := normalizeCLIPath(selector.Path)
			if key == "" {
				return fmt.Errorf("coordination command %q has an empty CLI selector", descriptor.ID)
			}
			if prior, exists := cli[key]; exists {
				return fmt.Errorf("duplicate CLI selector %q for %q and %q", key, prior, descriptor.ID)
			}
			cli[key] = descriptor.ID
		}
		for _, selector := range descriptor.Bridge {
			key := bridgeSelectorKey(selector)
			if strings.TrimSpace(selector.Command) == "" {
				return fmt.Errorf("coordination command %q has an empty bridge selector", descriptor.ID)
			}
			if prior, exists := bridge[key]; exists {
				return fmt.Errorf("duplicate bridge selector %q for %q and %q", key, prior, descriptor.ID)
			}
			bridge[key] = descriptor.ID
		}
	}
	return nil
}

func bridgeSelectorKey(selector BridgeSelector) string {
	keys := make([]string, 0, len(selector.Args))
	for key := range selector.Args {
		keys = append(keys, strings.ToLower(strings.TrimSpace(key)))
	}
	sort.Strings(keys)
	parts := []string{strings.ToLower(strings.TrimSpace(selector.Command))}
	for _, key := range keys {
		var value string
		for originalKey, originalValue := range selector.Args {
			if strings.EqualFold(strings.TrimSpace(originalKey), key) {
				value = strings.ToLower(strings.TrimSpace(originalValue))
				break
			}
		}
		parts = append(parts, key+"="+value)
	}
	return strings.Join(parts, "|")
}
