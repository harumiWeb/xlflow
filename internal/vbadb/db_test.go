package vbadb

import "testing"

func TestLoadBuiltinResolvesCoreExcelAndCommonCOMTypes(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	if _, ok := db.ResolveType("Excel.Application"); !ok {
		t.Fatal("Excel.Application was not loaded")
	}
	if _, ok := db.ResolveType("Workbook"); !ok {
		t.Fatal("Workbook alias did not resolve")
	}
	if typ, ok := db.ResolveProgID("Scripting.Dictionary"); !ok || typ.Name != "Scripting.Dictionary" {
		t.Fatalf("ResolveProgID(Scripting.Dictionary) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveGlobal("Worksheets"); !ok || typ.Name != "Excel.Worksheets" {
		t.Fatalf("ResolveGlobal(Worksheets) = %+v, %v", typ, ok)
	}
}

func TestResolveMemberHandlesCollectionDefaultMembersAndFactories(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	if got, ok := db.ResolveMember("Excel.Worksheets", "Item"); !ok || got.ReturnType != "Excel.Worksheet" {
		t.Fatalf("Worksheets.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Workbooks", "Open"); !ok || got.ReturnType != "Excel.Workbook" {
		t.Fatalf("Workbooks.Open = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Worksheet", "Range"); !ok || got.ReturnType != "Excel.Range" {
		t.Fatalf("Worksheet.Range = %+v, %v", got, ok)
	}
}

func TestResolveConstant(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	constant, ok := db.ResolveConstant("xlUp")
	if !ok {
		t.Fatal("xlUp was not loaded")
	}
	if constant.Library != "Excel" || constant.EnumGroup != "XlDirection" {
		t.Fatalf("unexpected xlUp metadata: %+v", constant)
	}
}
