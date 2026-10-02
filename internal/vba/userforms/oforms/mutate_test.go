package oforms

import (
	"bytes"
	"encoding/binary"
	"errors"
	"maps"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestApplyEditsNoOpAndOwnership(t *testing.T) {
	for _, fixture := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		t.Run(fixture, func(t *testing.T) {
			original := openFixture(t, fixture)
			base, err := ReadForm(original, "UserForm1", 932)
			if err != nil {
				t.Fatal(err)
			}
			got, err := ApplyEdits(base, nil, 932)
			if err != nil {
				t.Fatal(err)
			}
			stored, err := SerializeForm(got, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertSerializedSubtree(t, original, "UserForm1", stored)
			before, err := SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(before, stored) {
				for path, raw := range before.Streams {
					if !reflect.DeepEqual(raw, stored.Streams[path]) {
						t.Errorf("stream %s differs: nil %v -> %v, bytes %d -> %d", path, raw == nil, stored.Streams[path] == nil, len(raw), len(stored.Streams[path]))
					}
				}
				if !maps.Equal(before.Storages, stored.Storages) {
					t.Errorf("metadata differs: %v -> %v", before.Storages, stored.Storages)
				}
				t.Fatal("no-op differs")
			}
			got.Levels[0].FRaw[0] ^= 1
			if _, err := SerializeForm(base, 932); err != nil {
				t.Fatalf("base aliased: %v", err)
			}
		})
	}
}

func TestApplyEditsFixtureRoundTrip(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	edits := []Edit{
		{"", "Caption", "日本語😀フォーム"},
		{"CommandButton1", "Caption", "実行😀"},
		{"CommandButton1", "Left", int32(-17)},
		{"CommandButton1", "Top", int32(42)},
		{"CommandButton1", "Width", int32(12345)},
		{"CommandButton1", "Height", int32(6789)},
		{"CommandButton1", "TabIndex", int16(7)},
		{"CommandButton1", "Tag", "日本語タグ"},
		{"CommandButton1", "ControlTipText", "😀 hint"},
		{"CommandButton1", "BackColor", uint32(0x80000005)},
		{"CommandButton1", "ForeColor", uint32(0xffffff)},
	}
	got, err := ApplyEdits(base, edits, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.Levels[0].Record.Strings["Caption"].Text != "日本語😀フォーム" {
		t.Fatal("root caption lost")
	}
	c := got.Controls[0]
	if c.Record.Strings["Caption"].Text != "実行😀" || c.Record.Strings["Caption"].Compressed {
		t.Fatal("caption fallback lost")
	}
	if *c.Site.Position != (Position{-17, 42}) || c.Record.Sizes["Size"] != (Size{12345, 6789}) {
		t.Fatal("geometry lost")
	}
	if c.TabIndex == nil || *c.TabIndex != 7 {
		t.Fatal("TabIndex lost")
	}
	if c.Site.Strings["Tag"].Text != "日本語タグ" || c.Site.Strings["Tag"].Compressed {
		t.Fatal("UTF-16 Tag lost")
	}
	if c.Site.Strings["ControlTipText"].Text != "😀 hint" {
		t.Fatal("tooltip lost")
	}
	if c.Record.Values["BackColor"] != 0x80000005 || c.Record.Values["ForeColor"] != 0xffffff {
		t.Fatal("colors lost")
	}
	if !bytes.Equal(c.Record.TextProps.Raw, base.Controls[0].Record.TextProps.Raw) {
		t.Fatal("TextProps changed")
	}
	for i := 1; i < len(base.Controls); i++ {
		if !reflect.DeepEqual(recordSignature(base.Controls[i].Record), recordSignature(got.Controls[i].Record)) || !bytes.Equal(base.Controls[i].Site.Raw, got.Controls[i].Site.Raw) {
			t.Fatalf("unedited control %d changed", i)
		}
	}
	after, err := SerializeForm(got, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !maps.Equal(before.Storages, after.Storages) {
		t.Fatal("metadata changed")
	}
	for path, raw := range before.Streams {
		if path != "UserForm1/f" && path != "UserForm1/o" && !bytes.Equal(raw, after.Streams[path]) {
			t.Fatalf("stream %s changed", path)
		}
	}
	again, err := ApplyEdits(base, edits, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.sourceSignature != again.sourceSignature {
		t.Fatal("non-deterministic edits")
	}
	repeated, err := ApplyEdits(got, edits, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.sourceSignature != repeated.sourceSignature {
		t.Fatal("idempotent edits changed persistence")
	}
	unchanged, err := SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(before, unchanged) {
		t.Fatal("base mutated")
	}
	got.Controls[0].Record.Values["BackColor"] = 123
	if _, err := SerializeForm(got, 932); !errors.Is(err, ErrUnsupportedMutation) {
		t.Fatalf("unsigned mutation accepted: %v", err)
	}
}

func TestApplyEditsInvalidAtomicAndContext(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		edit     Edit
		sentinel error
	}{
		{Edit{"", "Width", int32(5)}, ErrUnsupportedEdit},
		{Edit{"missing", "Caption", "x"}, ErrInvalidEdit},
		{Edit{"CommandButton1", "Width", -1}, ErrInvalidEdit},
		{Edit{"CommandButton1", "Width", int32(-1)}, ErrInvalidEdit},
		{Edit{"CommandButton1", "TabIndex", int16(-1)}, ErrInvalidEdit},
		{Edit{"CommandButton1", "TabIndex", 5}, ErrInvalidEdit},
		{Edit{"CommandButton1", "BackColor", uint64(1 << 32)}, ErrInvalidEdit},
		{Edit{"CommandButton1", "ForeColor", -1}, ErrInvalidEdit},
		{Edit{"CommandButton1", "Caption", 7}, ErrInvalidEdit},
		{Edit{"CommandButton1", "Caption", string([]byte{0xff})}, ErrInvalidEdit},
		{Edit{"CommandButton1", "Value", "x"}, ErrUnsupportedEdit},
		{Edit{"CommandButton1", "List", []string{"x"}}, ErrUnsupportedEdit},
	} {
		t.Run(tc.edit.Control+tc.edit.Property+reflect.TypeOf(tc.edit.Value).String(), func(t *testing.T) {
			got, err := ApplyEdits(base, []Edit{{"", "Caption", "partial"}, tc.edit}, 932)
			if got != nil || !errors.Is(err, tc.sentinel) {
				t.Fatalf("got=%v err=%v", got, err)
			}
			context, _ := errors.AsType[*EditError](err)
			if context == nil || context.Control != tc.edit.Control || context.Property != tc.edit.Property {
				t.Fatalf("context=%v", context)
			}
			if _, err := SerializeForm(base, 932); err != nil {
				t.Fatalf("base modified: %v", err)
			}
		})
	}
	base.Controls[0].Name = "changed"
	if _, err := ApplyEdits(base, nil, 932); !errors.Is(err, ErrUnsupportedMutation) {
		t.Fatalf("invalid base accepted: %v", err)
	}
}

func TestEditRecordEncoderSharedTables(t *testing.T) {
	for index, spec := range specsByCacheIndex {
		t.Run(controlKinds[index], func(t *testing.T) {
			r := &Record{Type: spec.typeName, Major: spec.major, Values: map[string]int64{}, Strings: map[string]StoredString{}, Sizes: map[string]Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{}, Padding: map[string][]byte{}}
			for _, f := range spec.data {
				r.Mask |= uint64(1) << f.bit
				r.Values[f.name] = 1
				if f.kind == fieldMarker {
					r.Values[f.name] = 0xffff
				}
			}
			for _, extra := range spec.extra {
				r.Mask |= uint64(1) << extra.bit
				switch extra.kind {
				case extraSize:
					r.Sizes[extra.name] = Size{123, 456}
				case extraString:
					r.Strings[extra.name] = StoredString{Text: "abc", Compressed: true, Raw: []byte("abc")}
					r.Values[extra.name] = packedStringLength(r.Strings[extra.name])
				case extraArray:
					r.Arrays[extra.name] = []byte{1, 2, 3, 4}
					r.Values[extra.sizeFrom] = 4
				}
			}
			for _, stream := range spec.stream {
				picture := make([]byte, 24)
				picture[0] = 0x7f
				r.Pictures[stream.name] = picture
			}
			if spec.textProps {
				r.TextProps = &Record{Raw: []byte{0, 2, 4, 0, 0, 0, 0, 0}}
			}
			if spec.rawTail {
				r.TailRaw = []byte{7, 8, 9}
			}
			for bit := range spec.flags {
				r.Mask |= uint64(1) << bit
			}
			body, err := encodeEditedRecord(r, spec)
			if err != nil {
				t.Fatal(err)
			}
			parsed, _, err := parseRecord(body, spec, 932, len(body))
			if err != nil {
				t.Fatal(err)
			}
			if !maps.Equal(parsed.Values, r.Values) && len(spec.flags) == 0 {
				t.Fatal("scalar values lost")
			}
			if !reflect.DeepEqual(parsed.Strings, r.Strings) || !reflect.DeepEqual(parsed.Sizes, r.Sizes) || !reflect.DeepEqual(parsed.Pictures, r.Pictures) || !bytes.Equal(parsed.TailRaw, r.TailRaw) {
				t.Fatal("extra/stream payload lost")
			}
			reencoded, err := encodeEditedRecord(parsed, spec)
			if err != nil || !bytes.Equal(body, reencoded) {
				t.Fatalf("reencode differs: %v", err)
			}
			r.Mask |= uint64(1) << 63
			if _, err := encodeEditedRecord(r, spec); !errors.Is(err, ErrUnsupportedEdit) {
				t.Fatalf("unknown bit accepted: %v", err)
			}
		})
	}
}

func TestEditSiteEncoderStringsAndPadding(t *testing.T) {
	s := &Site{Values: map[string]int64{}, Strings: map[string]StoredString{}, Padding: map[string][]byte{}, Position: &Position{-100, 200}}
	s.Mask = 1 << 8
	for _, f := range siteFields {
		s.Mask |= uint32(1) << f.bit
		s.Values[f.name] = 1
	}
	for _, items := range [][][2]string{siteStringsBeforePosition, siteStringsAfterPosition} {
		for _, item := range items {
			s.Strings[item[1]] = StoredString{Text: "xyz", Compressed: true, Raw: []byte("xyz")}
			s.Values[item[0]] = packedStringLength(s.Strings[item[1]])
			s.Padding["str:"+item[1]] = []byte{0xee}
		}
	}
	body, err := encodeEditedSite(s)
	if err != nil {
		t.Fatal(err)
	}
	parsed, err := parseSite(newByteReader(body), 932)
	if err != nil {
		t.Fatal(err)
	}
	if !maps.Equal(parsed.Values, s.Values) || !reflect.DeepEqual(parsed.Strings, s.Strings) || *parsed.Position != *s.Position {
		t.Fatal("site payload lost")
	}
	if !bytes.Equal(parsed.Padding["str:Tag"], []byte{0xee}) {
		t.Fatal("padding lost")
	}
	s.Mask |= 1 << 31
	if _, err := encodeEditedSite(s); !errors.Is(err, ErrUnsupportedEdit) {
		t.Fatalf("unknown mask accepted: %v", err)
	}
}

func TestApplyEditsOmittedDefaultsAndBits(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	// Seed persisted bits, not assumptions about omitted MSForms defaults.
	r := base.Controls[0].Record
	r.Mask |= 1 << 2
	r.Values["VariousPropertyBits"] = 0x1234567a
	raw, err := encodeEditedRecord(r, &commandButtonSpec)
	if err != nil {
		t.Fatal(err)
	}
	s := base.Controls[0].Site
	s.Mask |= 1 << 4
	s.Values["BitFlags"] = 0x1234567a
	s.Values["ObjectStreamSize"] = int64(len(raw))
	s.Raw, err = encodeEditedSite(s)
	if err != nil {
		t.Fatal(err)
	}
	base.Levels[0].ORaw = raw
	base.Levels[0].FRaw, err = encodeEditedLevel(base.Levels[0])
	if err != nil {
		t.Fatal(err)
	}
	stored := &SerializedForm{Streams: map[string][]byte{}, Storages: map[string]cfb.StorageMeta{}}
	if err := appendLevel(stored, base.Name, base.Levels[0]); err != nil {
		t.Fatal(err)
	}
	seed, err := reparseEditedForm(stored, base.Name, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := ApplyEdits(seed, []Edit{{"CommandButton1", "Enabled", false}, {"CommandButton1", "Visible", false}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Record.Values["VariousPropertyBits"] != 0x12345678 || got.Controls[0].Site.Values["BitFlags"] != 0x12345678 {
		t.Fatal("unrelated bits changed")
	}
	// The specified CommandButton omission default is 0x1b.
	delete(r.Values, "VariousPropertyBits")
	r.Mask &^= 1 << 2
	raw, err = encodeEditedRecord(r, &commandButtonSpec)
	if err != nil {
		t.Fatal(err)
	}
	stored.Streams["UserForm1/o"] = raw
	s.Values["ObjectStreamSize"] = int64(len(raw))
	s.Raw, err = encodeEditedSite(s)
	if err != nil {
		t.Fatal(err)
	}
	stored.Streams["UserForm1/f"], err = encodeEditedLevel(base.Levels[0])
	if err != nil {
		t.Fatal(err)
	}
	seed, err = reparseEditedForm(stored, base.Name, 932)
	if err != nil {
		t.Fatal(err)
	}
	enabled, err := ApplyEdits(seed, []Edit{{"CommandButton1", "Enabled", false}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if enabled.Controls[0].Record.Values["VariousPropertyBits"] != 0x19 {
		t.Fatal("incorrect omission default")
	}
}

func TestEditedLevelPreservesClassTableDistinction(t *testing.T) {
	for _, class := range [][]byte{nil, {0, 0}} {
		level := &Level{Record: &Record{Raw: []byte{0, 4, 4, 0, 0, 0, 0, 0}}, ClassTableRaw: class}
		raw, err := encodeEditedLevel(level)
		if err != nil {
			t.Fatal(err)
		}
		parsed, err := parseFormStream(raw, "Test", 932)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(parsed.ClassTableRaw, class) {
			t.Fatalf("class table changed: %x -> %x", class, parsed.ClassTableRaw)
		}
		if binary.LittleEndian.Uint32(raw[len(raw)-4:]) != 0 {
			t.Fatal("CountOfBytes wrong")
		}
	}
}

func TestApplyEditsNestedFrameAndPreservedLevelData(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	var frame *Control
	for _, level := range base.Levels {
		for _, c := range level.Controls {
			if c.CLSIDCacheIndex == 14 {
				frame = c
				break
			}
		}
	}
	if frame == nil {
		t.Fatal("fixture has no Frame")
		return
	}
	got, err := ApplyEdits(base, []Edit{{frame.Name, "Caption", "nested 😀"}, {frame.Name, "Width", int32(7777)}, {frame.Name, "Height", int32(5555)}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	for i, level := range got.Levels {
		old := base.Levels[i]
		if level.Path == frame.Level.Path {
			if level.Record.Strings["Caption"].Text != "nested 😀" || level.Record.Sizes["DisplayedSize"] != (Size{7777, 5555}) {
				t.Fatal("Frame edits lost")
			}
		} else if !bytes.Equal(level.FRaw, old.FRaw) {
			t.Fatalf("other level changed: %s", level.Path)
		}
		if !bytes.Equal(level.ORaw, old.ORaw) || !bytes.Equal(level.ClassTableRaw, old.ClassTableRaw) || !bytes.Equal(level.DepthsRaw, old.DepthsRaw) || !bytes.Equal(level.FontRaw, old.FontRaw) || !bytes.Equal(level.PictureRaw, old.PictureRaw) || !bytes.Equal(level.TrailingRaw, old.TrailingRaw) {
			t.Fatalf("retained level data changed: %s", level.Path)
		}
	}
	if _, err := SerializeForm(got, 932); err != nil {
		t.Fatal(err)
	}
	disabled, err := ApplyEdits(base, []Edit{{frame.Name, "Enabled", false}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, control := range disabled.Controls {
		if control.Name == frame.Name && control.Record.Values["BooleanProperties"] != frame.Record.Values["BooleanProperties"]&^4 {
			t.Fatal("Frame Enabled changed other bits")
		}
	}
}

func TestFormsStringCompressionIsIndependentOfProjectCodePage(t *testing.T) {
	for _, codePage := range []uint16{932, 1252} {
		for _, text := range []string{"café ÿ", "日本語", "shell 🐚"} {
			stored, err := editedString(StoredString{Compressed: true}, text, codePage)
			if err != nil {
				t.Fatal(err)
			}
			wantCompressed := text == "café ÿ"
			if stored.Compressed != wantCompressed {
				t.Fatalf("%q compressed=%v", text, stored.Compressed)
			}
			if wantCompressed && !bytes.Equal(stored.Raw, []byte{'c', 'a', 'f', 0xe9, ' ', 0xff}) {
				t.Fatal("compressed bytes are not low UTF-16 bytes")
			}
			decoded, err := decodeStoredText(stored.Raw, stored.Compressed, codePage)
			if err != nil || decoded != text {
				t.Fatalf("decoded=%q err=%v", decoded, err)
			}
		}
	}
}

func TestApplyEditsUnknownRecordMaskAndLengthErrors(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	object, _ := original.Stream("UserForm1/o")
	object = bytes.Clone(object)
	binary.LittleEndian.PutUint32(object[4:8], binary.LittleEndian.Uint32(object[4:8])|1<<31)
	base, err := ReadForm(rewriteContainer(t, original, map[string][]byte{"UserForm1/o": object}, nil), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ApplyEdits(base, nil, 932); err != nil {
		t.Fatal("unknown bits blocked no-op", err)
	}
	_, err = ApplyEdits(base, []Edit{{"CommandButton1", "Caption", "x"}}, 932)
	context, _ := errors.AsType[*EditError](err)
	if !errors.Is(err, ErrUnsupportedEdit) || context == nil || context.Property != "Caption" {
		t.Fatalf("unknown mask error=%v", err)
	}
	base, err = ReadForm(original, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	_, err = ApplyEdits(base, []Edit{{"CommandButton1", "Caption", strings.Repeat("a", 65535)}}, 932)
	context, _ = errors.AsType[*EditError](err)
	if !errors.Is(err, ErrInvalidEdit) || context == nil || context.Property != "Caption" {
		t.Fatalf("length error=%v", err)
	}
	got, err := ApplyEdits(base, []Edit{{"CommandButton1", "Caption", ""}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Record.Strings["Caption"].Text != "" {
		t.Fatal("empty string lost")
	}
}

func TestApplyEditsPreservesCaseNamesAndOpaqueStreams(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	w := cfb.NewWriter()
	for _, path := range original.StoragePaths() {
		meta, _ := original.Storage(path)
		w.AddStorage(splitPath(path), meta)
	}
	for _, path := range original.Paths() {
		parts := splitPath(path)
		if strings.HasPrefix(path, "UserForm1/") {
			parts[len(parts)-1] = strings.ToUpper(parts[len(parts)-1])
		}
		raw, _ := original.Stream(path)
		w.AddStream(parts, raw)
	}
	w.AddStream([]string{"UserForm1", "vendor"}, []byte{9, 8, 7})
	body, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	c, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	base, err := ReadForm(c, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := ApplyEdits(base, []Edit{{"", "Caption", "edited"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := SerializeForm(got, 932)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := stored.Streams["UserForm1/F"]; !ok {
		t.Fatal("case changed")
	}
	if _, ok := stored.Streams["UserForm1/O"]; !ok {
		t.Fatal("case changed")
	}
	if !bytes.Equal(stored.Streams["UserForm1/vendor"], []byte{9, 8, 7}) {
		t.Fatal("opaque stream changed")
	}
}

func TestApplyEditsVisibleOmittedDefaults(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	if _, found := base.Controls[0].Site.Values["BitFlags"]; found {
		t.Fatal("expected omitted fixture BitFlags")
	}
	got, err := ApplyEdits(base, []Edit{{"CommandButton1", "Visible", false}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Site.Values["BitFlags"] != 0x31 {
		t.Fatal("commandbutton omitted default lost")
	}
	for _, tc := range []struct {
		index uint16
		want  int64
	}{{21, 0x30}, {14, 0x40021}, {7, 0x40021}, {57, 0x40021}, {17, 0x31}} {
		s := &Site{Values: map[string]int64{}}
		form := &Form{Levels: []*Level{{Controls: []*Control{{Name: "test", CLSIDCacheIndex: tc.index, Site: s}}}}}
		if err := applyPersistenceEdit(form, Edit{"test", "Visible", false}, 932, map[*Record]Edit{}, map[*Site]Edit{}); err != nil {
			t.Fatal(err)
		}
		if s.Values["BitFlags"] != tc.want || s.Mask&(1<<4) == 0 {
			t.Fatalf("index %d flags %#x", tc.index, s.Values["BitFlags"])
		}
	}
}

func TestApplyEditsMorphReservedBitPreserved(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	var original *Control
	for _, level := range base.Levels {
		for _, c := range level.Controls {
			if c.Name == "InnerText" {
				original = c
			}
		}
	}
	if original == nil {
		t.Fatal("missing InnerText")
		return
	}
	if original.Record.Mask&(1<<31) == 0 {
		t.Fatal("fixture missing reserved bit")
	}
	got, err := ApplyEdits(base, []Edit{{"InnerText", "Value", "nested 😀 text"}, {"InnerText", "GroupName", "group"}, {"InnerText", "MaxLength", uint32(120)}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, level := range got.Levels {
		for _, c := range level.Controls {
			if c.Name == "InnerText" {
				if c.Record.Mask&(1<<31) == 0 || c.Record.Strings["Value"].Text != "nested 😀 text" || c.Record.Strings["GroupName"].Text != "group" || c.Record.Values["MaxLength"] != 120 {
					t.Fatal("morph edits or reserved bit lost")
				}
			}
		}
	}
	if _, err := SerializeForm(got, 932); err != nil {
		t.Fatal(err)
	}
}
