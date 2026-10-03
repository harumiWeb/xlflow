package oforms

import (
	"bytes"
	"fmt"
	"math"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

// TopologyControl describes one retained or newly authored control. Parent is
// a control name; an empty parent denotes the form. Input order is sibling order.
// Definition is required only for additions and explicit replacements.
type TopologyControl struct {
	Name       string
	Parent     string
	Definition *ControlDefinition
}

// ApplyTopology changes ownership on an independent, validated persistence
// tree. Unchanged records and opaque streams remain retained bytes.
func ApplyTopology(base *Form, desired []TopologyControl, codePage uint16) (*Form, error) {
	stored, err := SerializeForm(base, codePage)
	if err != nil {
		return nil, err
	}
	clone, err := reparseEditedForm(stored, base.Name, codePage)
	if err != nil {
		return nil, err
	}
	var current []TopologyControl
	keys := map[*Control]string{}
	levelOwners := map[*Level]string{clone.Levels[0]: ""}
	unnamed := 0
	var walk func([]*Control, string)
	walk = func(controls []*Control, parent string) {
		for _, c := range controls {
			if c.Name == "" && c.CLSIDCacheIndex == 18 {
				keys[c] = "<hidden:" + parent + ">"
				continue
			}
			name := c.Name
			if strings.TrimSpace(name) == "" {
				unnamed++
				name = fmt.Sprintf("<unnamed_%d>", unnamed)
			}
			keys[c] = name
			if c.Level != nil {
				levelOwners[c.Level] = name
			}
			current = append(current, TopologyControl{Name: name, Parent: parent})
			if c.MultiPage != nil {
				keys[c.MultiPage.Hidden] = "<hidden:" + name + ">"
				walk(c.MultiPage.Pages, name)
			} else {
				walk(c.Children, name)
			}
		}
	}
	walk(clone.Controls, "")
	if slices.Equal(current, desired) {
		return clone, nil
	}
	if len(desired) > maxSitesPerForm {
		return nil, fmt.Errorf("%w: total site limit", ErrInvalidEdit)
	}
	existing := map[string]*Control{}
	owners := map[*Control]*Level{}
	payloads := map[*Control][]byte{}
	original := map[*Level][]*Control{}
	originalTabs := map[*Control]Tab{}
	originalPageProperties := map[int32][]byte{}
	newLevels := map[*Level]bool{}
	var nextID int64 = 1
	for _, level := range clone.Levels {
		original[level] = slices.Clone(level.Controls)
		nextID = max(nextID, level.Record.Values["NextAvailableID"])
		offset := 0
		for _, control := range level.Controls {
			if control.MultiPage != nil {
				for i, page := range control.MultiPage.Pages {
					originalTabs[page] = control.MultiPage.Hidden.TabStrip.Tabs[i]
					originalPageProperties[page.ID] = bytes.Clone(control.MultiPage.PageProperties[page.ID])
				}
			}
			if existing[keys[control]] != nil {
				return nil, fmt.Errorf("%w: ambiguous control name", ErrInvalidEdit)
			}
			existing[keys[control]], owners[control] = control, level
			nextID = max(nextID, int64(control.ID)+1)
			end := offset + int(control.ObjectStreamSize)
			payloads[control] = bytes.Clone(level.ORaw[offset:end])
			offset = end
		}
	}
	selected := map[string]*Control{}
	parents := map[string]string{}
	children := map[string][]*Control{}
	caseNames := map[string]bool{}
	for _, item := range desired {
		key := strings.ToLower(item.Name)
		if item.Name == "" || caseNames[key] {
			return nil, fmt.Errorf("%w: duplicate or empty name", ErrInvalidEdit)
		}
		caseNames[key] = true
		control := existing[item.Name]
		if item.Definition != nil {
			if item.Definition.Name != item.Name || len(item.Definition.Controls) != 0 {
				return nil, fmt.Errorf("%w: topology definition identity/children", ErrInvalidEdit)
			}
			generated, err := NewForm(Definition{Name: "TopologyScratch", Size: Size{100, 100}, Controls: []ControlDefinition{*item.Definition}}, codePage)
			if err != nil {
				return nil, err
			}
			control = generated.Controls[0]
			if nextID > math.MaxInt32 {
				return nil, fmt.Errorf("%w: site ID overflow", ErrInvalidEdit)
			}
			control.ID = int32(nextID)
			control.Site.Values["ID"] = nextID
			nextID++
			control.Site.Raw, err = encodeEditedSite(control.Site)
			if err != nil {
				return nil, err
			}
			payloads[control] = bytes.Clone(generated.Levels[0].ORaw)
			if control.Level != nil {
				original[control.Level] = slices.Clone(control.Level.Controls)
				newLevels[control.Level] = true
				if control.MultiPage != nil {
					hidden := control.MultiPage.Hidden
					hidden.ID = int32(nextID)
					hidden.Site.Values["ID"] = nextID
					nextID++
					hidden.Site.Raw, err = encodeEditedSite(hidden.Site)
					if err != nil {
						return nil, err
					}
					payloads[hidden] = bytes.Clone(control.Level.ORaw)
				}
			}
		}
		if control == nil {
			return nil, fmt.Errorf("%w: missing retained control %q", ErrInvalidEdit, item.Name)
		}
		selected[item.Name], parents[item.Name] = control, item.Parent
		keys[control] = item.Name
		children[item.Parent] = append(children[item.Parent], control)
	}
	for name, control := range selected {
		if control.MultiPage != nil {
			hidden := control.MultiPage.Hidden
			keys[hidden] = "<hidden:" + name + ">"
			children[name] = append([]*Control{hidden}, children[name]...)
		}
	}
	for name, parent := range parents {
		if parent == "" {
			if selected[name].CLSIDCacheIndex == 7 {
				return nil, fmt.Errorf("%w: Page requires a MultiPage parent", ErrUnsupportedEdit)
			}
			continue
		}
		p := selected[parent]
		if p == nil || p.Level == nil || (p.CLSIDCacheIndex == 57) != (selected[name].CLSIDCacheIndex == 7) {
			return nil, fmt.Errorf("%w: invalid parent %q", ErrUnsupportedEdit, parent)
		}
		seen := map[string]bool{name: true}
		for at := parent; at != ""; at = parents[at] {
			if seen[at] {
				return nil, fmt.Errorf("%w: cyclic hierarchy", ErrInvalidEdit)
			}
			seen[at] = true
		}
	}
	result := &SerializedForm{Streams: map[string][]byte{}, Storages: map[string]cfb.StorageMeta{}}
	cookieChanges := map[*Level]int64{}
	var oldMembers func(*Level, map[*Control]bool)
	oldMembers = func(level *Level, members map[*Control]bool) {
		for _, control := range original[level] {
			members[control] = true
			if control.Level != nil {
				oldMembers(control.Level, members)
			}
		}
	}
	var newMembers func(string, map[*Control]bool, int) error
	newMembers = func(owner string, members map[*Control]bool, depth int) error {
		if depth >= maxNestingDepth {
			return fmt.Errorf("%w: nesting limit", ErrInvalidEdit)
		}
		for _, control := range children[owner] {
			members[control] = true
			if control.Level != nil {
				if err := newMembers(keys[control], members, depth+1); err != nil {
					return err
				}
			}
		}
		return nil
	}
	for level := range original {
		owner := levelOwners[level]
		if newLevels[level] {
			for name, control := range selected {
				if control.Level == level {
					owner = name
					break
				}
			}
		}
		old, next := map[*Control]bool{}, map[*Control]bool{}
		oldMembers(level, old)
		if err := newMembers(owner, next, 0); err != nil {
			return nil, err
		}
		for control := range old {
			if !next[control] {
				cookieChanges[level]++
			}
		}
		for control := range next {
			if !old[control] {
				cookieChanges[level]++
			}
		}
	}
	visited := 0
	var rebuild func(*Level, string, string, int) error
	rebuild = func(level *Level, owner, path string, depth int) error {
		if depth >= maxNestingDepth {
			return fmt.Errorf("%w: nesting limit", ErrInvalidEdit)
		}
		wanted := children[owner]
		changed := newLevels[level] || !slices.Equal(original[level], wanted)
		relocated := !newLevels[level] && level.Path != path
		multi := owner != "" && selected[owner].MultiPage != nil
		if (changed || relocated) && (level.HasXStream && !multi || len(level.ExtraStreams) > 0 || len(level.ClassTable) > 0 || len(level.TrailingRaw) > 0 && !multi || level.Record.Values["GroupCnt"] != 0) {
			return fmt.Errorf("%w: unimplemented container bookkeeping", ErrUnsupportedEdit)
		}
		if changed {
			if multi {
				if err := rebuildMultiPage(selected[owner], wanted, originalTabs, originalPageProperties, payloads); err != nil {
					return err
				}
			}
			level.Controls, level.Sites, level.ORaw, level.DepthsRaw = wanted, nil, nil, nil
			for _, control := range wanted {
				if control.Depth != 0 || control.SiteType != 1 {
					return fmt.Errorf("%w: nonstandard site depth/type", ErrUnsupportedEdit)
				}
				if previous := owners[control]; previous != nil && previous != level {
					if len(previous.ClassTable) > 0 || control.CLSIDCacheIndex < 7 {
						return fmt.Errorf("%w: class table relocation", ErrUnsupportedEdit)
					}
				}
				level.Sites = append(level.Sites, control.Site)
				level.ORaw = append(level.ORaw, payloads[control]...)
				level.DepthsRaw = append(level.DepthsRaw, control.Depth, control.SiteType)
			}
			level.DepthsRaw = appendEditPadding(level.DepthsRaw, 4, nil)
			if len(level.ORaw) > maxDesignerStreamSize {
				return fmt.Errorf("%w: object stream limit", ErrInvalidEdit)
			}
		}
		counterChanged := cookieChanges[level] != 0 || owner == "" && nextID > level.Record.Values["NextAvailableID"]
		if changed || counterChanged {
			if err := setGenerationProperty(level.Record, &formSpec, "NextAvailableID", max(nextID, level.Record.Values["NextAvailableID"])); err != nil {
				return err
			}
			if cookieChanges[level] != 0 {
				if err := setGenerationProperty(level.Record, &formSpec, "ShapeCookie", level.Record.Values["ShapeCookie"]+cookieChanges[level]); err != nil {
					return err
				}
			}
			level.Record.Raw, err = encodeEditedRecord(level.Record, &formSpec)
			if err != nil {
				return err
			}
			level.FRaw, err = encodeEditedLevel(level)
			if err != nil {
				return err
			}
		}
		level.Path = path
		if owner == "" && cookieChanges[level] != 0 && level.HasVBFrame {
			level.VBFrameRaw, err = rewriteVBFrameTypeInfo(level.VBFrameRaw, level.Record.Values["ShapeCookie"], codePage)
			if err != nil {
				return err
			}
		}
		if err := appendLevel(result, base.Name, level); err != nil {
			return err
		}
		for _, control := range wanted {
			if control.CLSIDCacheIndex != 18 || control.Name != "" {
				visited++
			}
			if control.Level != nil {
				childPath := fmt.Sprintf("%s/i%02d", path, control.ID)
				if owners[control] == level {
					childPath = path + "/" + control.Level.Path[strings.LastIndexByte(control.Level.Path, '/')+1:]
				}
				if err := rebuild(control.Level, keys[control], childPath, depth+1); err != nil {
					return err
				}
			}
		}
		return nil
	}
	if err := rebuild(clone.Levels[0], "", base.Name, 0); err != nil {
		return nil, err
	}
	if visited != len(desired) {
		return nil, fmt.Errorf("%w: unreachable controls", ErrInvalidEdit)
	}
	accepted, err := reparseEditedForm(result, base.Name, codePage)
	if err != nil {
		return nil, err
	}
	var actual []TopologyControl
	unnamed = 0
	var verify func([]*Control, string)
	verify = func(controls []*Control, parent string) {
		for _, c := range controls {
			if c.Name == "" && c.CLSIDCacheIndex == 18 {
				continue
			}
			name := c.Name
			if strings.TrimSpace(name) == "" {
				unnamed++
				name = fmt.Sprintf("<unnamed_%d>", unnamed)
			}
			actual = append(actual, TopologyControl{Name: name, Parent: parent})
			if c.MultiPage != nil {
				verify(c.MultiPage.Pages, name)
			} else {
				verify(c.Children, name)
			}
		}
	}
	verify(accepted.Controls, "")
	// Compare parent/sibling relationships independently of whether callers used
	// preorder or another valid flat authoring order.
	actualChildren := map[string][]string{}
	for _, c := range actual {
		actualChildren[c.Parent] = append(actualChildren[c.Parent], c.Name)
	}
	for parent, cs := range children {
		names := make([]string, len(cs))
		for i, c := range cs {
			names[i] = keys[c]
		}
		if len(cs) > 0 && cs[0].Name == "" && cs[0].CLSIDCacheIndex == 18 {
			names = names[1:]
		}
		if !slices.Equal(names, actualChildren[parent]) {
			return nil, fmt.Errorf("%w: hierarchy read-back differs for %q", ErrInvalidEdit, parent)
		}
	}
	return accepted, nil
}
