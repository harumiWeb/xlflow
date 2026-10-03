package vbaproject

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

func TestWithoutUserFormRemovesOwnedStateAndPreservesReferences(t *testing.T) {
	for _, fixture := range []string{"p4_form", "p6_nested_form"} {
		t.Run(fixture, func(t *testing.T) {
			p, err := Read(loadCorpus(t, fixture))
			if err != nil {
				t.Fatal(err)
			}
			name := p.Forms[0].Name
			p.ProjectStreamRaw = append(p.ProjectStreamRaw, []byte("\r\n[Workspace]\r\n"+name+"=0, 0, 0, 0, C\r\nKeep=0, 0, 0, 0, C\r\n")...)
			before, err := Clone(p)
			if err != nil {
				t.Fatal(err)
			}
			result, err := WithoutUserForm(p, name)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(before, p) {
				t.Fatal("input changed")
			}
			body, err := Write(result)
			if err != nil {
				t.Fatal(err)
			}
			readback, err := Read(body)
			if err != nil {
				t.Fatal(err)
			}
			if len(readback.Forms) != 0 {
				t.Fatal("Designer remained")
			}
			for _, m := range readback.Modules {
				if m.Name == name {
					t.Fatal("module remained")
				}
			}
			if strings.Contains(string(readback.ProjectStreamRaw), name+"=") || strings.Contains(string(readback.ProjectStreamRaw), "BaseClass="+name) {
				t.Fatal("PROJECT form state remained")
			}
			if !strings.Contains(string(readback.ProjectStreamRaw), "Keep=") || !bytes.Equal(p.ReferencesRaw, readback.ReferencesRaw) {
				t.Fatal("unrelated project state changed")
			}
			for path := range readback.StorageMetadata {
				if firstSegment(path) == name {
					t.Fatal("orphan storage remained")
				}
			}
			if result, err := WithoutUserForm(p, "Missing"); err == nil || result != nil {
				t.Fatal("unknown form accepted")
			}
			if !reflect.DeepEqual(before, p) {
				t.Fatal("failed removal changed input")
			}
		})
	}
}
