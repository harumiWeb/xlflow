package oforms

import (
	"bytes"
	"fmt"
)

// TabEdit applies a complete, reconciled state after structural compilation.
type TabEdit struct {
	Control string
	State   *TabStrip
	Enabled *bool
}

func ApplyTabEdits(base *Form, edits []TabEdit, codePage uint16) (*Form, error) {
	stored, err := SerializeForm(base, codePage)
	if err != nil {
		return nil, err
	}
	clone, err := reparseEditedForm(stored, base.Name, codePage)
	if err != nil {
		return nil, err
	}
	changedRecords := map[*Record]bool{}
	changedLevels := map[*Level]bool{}
	for _, edit := range edits {
		var control *Control
		for _, level := range clone.Levels {
			for _, c := range level.Controls {
				if c.Name == edit.Control {
					if control != nil {
						return nil, ErrInvalidEdit
					}
					control = c
				}
			}
		}
		if control == nil {
			return nil, fmt.Errorf("%w: tab owner missing", ErrInvalidEdit)
		}
		target := control
		if control.MultiPage != nil {
			target = control.MultiPage.Hidden
		}
		if target.TabStrip == nil {
			return nil, ErrUnsupportedEdit
		}
		if edit.State != nil {
			if control.MultiPage != nil && len(control.MultiPage.Pages) == 0 && !equalTabStripState(target.TabStrip, edit.State) {
				return nil, fmt.Errorf("%w: empty MultiPage cached tab state must remain unchanged", ErrUnsupportedEdit)
			}
			if control.MultiPage != nil && len(edit.State.Tabs) != len(control.MultiPage.Pages) {
				if len(control.MultiPage.Pages) != 0 || !equalTabStripState(target.TabStrip, edit.State) {
					return nil, fmt.Errorf("%w: tabs/pages mismatch", ErrInvalidEdit)
				}
			}
			previous, err := ParseTabStrip(target.Record)
			if err != nil {
				return nil, err
			}
			before := bytes.Clone(target.Record.Raw)
			if err := SetTabStrip(target.Record, edit.State); err != nil {
				return nil, err
			}
			raw := before
			if !equalTabStripState(previous, edit.State) {
				raw, err = encodeEditedRecord(target.Record, &tabStripSpec)
				if err != nil {
					return nil, err
				}
			}
			if !bytes.Equal(before, raw) {
				target.Record.Raw = raw
				changedRecords[target.Record] = true
			}
			if control.MultiPage != nil && previous.SelectedIndex != edit.State.SelectedIndex {
				for i, page := range control.MultiPage.Pages {
					flags := page.Site.Values["BitFlags"]
					next := editedBit(flags, int32(i) == edit.State.SelectedIndex)
					if flags != next {
						page.Site.Mask |= 1 << 4
						page.Site.Values["BitFlags"] = next
						page.Site.Raw, err = encodeEditedSite(page.Site)
						if err != nil {
							return nil, err
						}
						changedLevels[control.Level] = true
					}
				}
			}
		}
		if edit.Enabled != nil && control.MultiPage != nil {
			props := control.MultiPage.Properties
			mask := props.Mask
			if *edit.Enabled {
				props.Mask &^= 1 << 3
			} else {
				props.Mask |= 1 << 3
			}
			if mask != props.Mask {
				control.Level.XRaw, err = encodeMultiPageX(control.MultiPage)
				if err != nil {
					return nil, err
				}
				stored.Streams[control.Level.Path+"/"+control.Level.XStreamName] = control.Level.XRaw
			}
		}
	}
	for _, level := range clone.Levels {
		offset := 0
		object := []byte(nil)
		objectChanged := false
		for _, control := range level.Controls {
			end := offset + int(control.ObjectStreamSize)
			raw := level.ORaw[offset:end]
			offset = end
			if changedRecords[control.Record] && control.Level == nil {
				raw = control.Record.Raw
				objectChanged = true
				control.Site.Values["ObjectStreamSize"] = int64(len(raw))
				control.Site.Mask |= 1 << 5
				control.Site.Raw, err = encodeEditedSite(control.Site)
				if err != nil {
					return nil, err
				}
				changedLevels[level] = true
			}
			object = append(object, raw...)
		}
		if objectChanged {
			if len(object) > maxDesignerStreamSize {
				return nil, ErrInvalidEdit
			}
			stored.Streams[level.Path+"/"+level.OStreamName] = object
		}
		if changedLevels[level] {
			level.FRaw, err = encodeEditedLevel(level)
			if err != nil {
				return nil, err
			}
			stored.Streams[level.Path+"/"+level.FStreamName] = level.FRaw
		}
	}
	return reparseEditedForm(stored, base.Name, codePage)
}
