package oforms

import (
	"bytes"
	"errors"
	"maps"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

func TestSerializeFormPreservesDesignerSubtree(t *testing.T) {
	for _, fixture := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		t.Run(fixture, func(t *testing.T) {
			original := openFixture(t, fixture)
			form, err := ReadForm(original, "UserForm1", 932)
			if err != nil {
				t.Fatal(err)
			}
			serialized, err := SerializeForm(form, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertSerializedSubtree(t, original, "UserForm1", serialized)
		})
	}
}

func TestSerializeFormPreservesCP932DesignerText(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	encoded, err := ovba.EncodeMBCS("日本語フォーム\r\n", 932)
	if err != nil {
		t.Fatal(err)
	}
	container := rewriteContainer(t, original, map[string][]byte{"UserForm1/\x03VBFrame": encoded}, nil)
	form, err := ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	if form.DesignerSource.Text != "日本語フォーム\r\n" {
		t.Fatalf("DesignerSource = %q", form.DesignerSource.Text)
	}
	serialized, err := SerializeForm(form, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(serialized.Streams["UserForm1/\x03VBFrame"], encoded) {
		t.Fatal("CP932 VBFrame bytes changed")
	}
}

func TestSerializeFormPreservesEmptyOptionalStream(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	if _, exists := original.Stream("UserForm1/x"); exists {
		t.Fatal("fixture unexpectedly already has an x stream")
	}
	writer, err := cfb.NewWriterForFormat(original.Format())
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range original.StoragePaths() {
		meta, _ := original.Storage(path)
		writer.AddStorage(splitPath(path), meta)
	}
	for _, path := range original.Paths() {
		body, _ := original.Stream(path)
		writer.AddStream(splitPath(path), body)
	}
	writer.AddStream([]string{"UserForm1", "x"}, []byte{})
	body, err := writer.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	form, err := ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	serialized, err := SerializeForm(form, 932)
	if err != nil {
		t.Fatal(err)
	}
	x, exists := serialized.Streams["UserForm1/x"]
	if !exists || len(x) != 0 {
		t.Fatalf("empty x stream = %x, present=%v", x, exists)
	}
}

func TestSerializeFormRejectsSemanticMutation(t *testing.T) {
	form, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	form.Controls[0].Name = "RenamedControl"
	_, err = SerializeForm(form, 932)
	if !errors.Is(err, ErrUnsupportedMutation) {
		t.Fatalf("SerializeForm error = %v, want ErrUnsupportedMutation", err)
	}
}

func TestSerializeFormRejectsMalformedRetainedStream(t *testing.T) {
	form, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	form.Levels[0].ORaw = append(form.Levels[0].ORaw, 0xff)
	_, err = SerializeForm(form, 932)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("SerializeForm error = %v, want ErrMalformed", err)
	}
}

func assertSerializedSubtree(t *testing.T, original *cfb.Container, root string, serialized *SerializedForm) {
	t.Helper()
	wantStreams := map[string][]byte{}
	for _, path := range original.Paths() {
		if strings.HasPrefix(path, root+"/") {
			body, _ := original.Stream(path)
			wantStreams[path] = body
		}
	}
	if !slices.Equal(slices.Sorted(maps.Keys(serialized.Streams)), slices.Sorted(maps.Keys(wantStreams))) {
		t.Fatalf("stream paths = %q, want %q", slices.Sorted(maps.Keys(serialized.Streams)), slices.Sorted(maps.Keys(wantStreams)))
	}
	for path, want := range wantStreams {
		if !bytes.Equal(serialized.Streams[path], want) {
			t.Errorf("stream %q changed", path)
		}
	}
	wantStorages := map[string]cfb.StorageMeta{}
	for _, path := range original.StoragePaths() {
		if path == root || strings.HasPrefix(path, root+"/") {
			wantStorages[path], _ = original.Storage(path)
		}
	}
	if !maps.Equal(serialized.Storages, wantStorages) {
		t.Errorf("storage metadata changed: got=%v want=%v", serialized.Storages, wantStorages)
	}
}
