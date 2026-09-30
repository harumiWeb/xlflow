package pack

import (
	"archive/zip"
	"bytes"
	"errors"
	"testing"
)

func TestBuildWorkbookRejectsPackageSignatureRelationships(t *testing.T) {
	tests := []struct {
		name             string
		relationshipType string
	}{
		{name: "legacy", relationshipType: "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"},
		{name: "agile compatibility", relationshipType: "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile"},
		{name: "agile", relationshipType: "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile"},
		{name: "v3", relationshipType: "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			template := packageFixture(t, []byte("not a CFB container"), map[string][]byte{
				vbaProjectRelationshipsPath: []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="` + test.relationshipType + `" Target="signature.bin"/></Relationships>`),
			})
			_, _, err := BuildWorkbook(template, nil)
			if !errors.Is(err, ErrSignedProject) {
				t.Fatalf("error = %v, want ErrSignedProject", err)
			}
		})
	}
}

func TestBuildWorkbookRejectsPackageSignatureContentTypes(t *testing.T) {
	contentTypes := []string{
		"application/vnd.ms-office.vbaProjectSignature",
		"application/vnd.ms-office.vbaProjectSignatureAgile",
		"application/vnd.ms-office.vbaProjectSignatureV3",
	}
	for _, contentType := range contentTypes {
		t.Run(contentType, func(t *testing.T) {
			template := packageFixture(t, []byte("not a CFB container"), map[string][]byte{
				contentTypesPath:   []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/signature.bin" ContentType="` + contentType + `"/></Types>`),
				"xl/signature.bin": []byte("package signature sentinel"),
			})
			_, _, err := BuildWorkbook(template, nil)
			if !errors.Is(err, ErrSignedProject) {
				t.Fatalf("error = %v, want ErrSignedProject", err)
			}
		})
	}
}

func TestBuildWorkbookDoesNotInferSignatureFromArbitraryPartName(t *testing.T) {
	template := packageFixture(t, readTestFile(t, "corpus", "p1_compiled.bin"), map[string][]byte{
		contentTypesPath:             []byte(`<Types><Metadata ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>`),
		"xl/vbaProjectSignature.bin": []byte("unreferenced filename sentinel"),
	})
	if _, _, err := BuildWorkbook(template, p1SourceModules(t)); err != nil {
		t.Fatalf("BuildWorkbook() error = %v", err)
	}
}

func TestBuildWorkbookRejectsMalformedSignatureMetadata(t *testing.T) {
	template := packageFixture(t, readTestFile(t, "corpus", "p1_compiled.bin"), map[string][]byte{
		vbaProjectRelationshipsPath: []byte(`<Relationships><Relationship`),
	})
	_, _, err := BuildWorkbook(template, nil)
	if !errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("error = %v, want ErrAmbiguousLayout", err)
	}
}

func packageFixture(t *testing.T, vbaProject []byte, entries map[string][]byte) []byte {
	t.Helper()
	fixtureEntries := make(map[string][]byte, len(entries)+2)
	for name, body := range entries {
		fixtureEntries[name] = body
	}
	if _, ok := fixtureEntries[contentTypesPath]; !ok {
		fixtureEntries[contentTypesPath] = []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	}
	fixtureEntries["xl/vbaProject.bin"] = vbaProject
	fixtureEntries["xl/workbook.xml"] = []byte(`<workbook/>`)

	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for name, body := range fixtureEntries {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buffer.Bytes()
}
