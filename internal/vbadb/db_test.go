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
	if typ, ok := db.ResolveType("Collection"); !ok || typ.Name != "VBA.Collection" {
		t.Fatalf("ResolveType(Collection) = %+v, %v", typ, ok)
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
	if got, ok := db.ResolveMember("Excel.Range", "Font"); !ok || got.ReturnType != "Excel.Font" {
		t.Fatalf("Range.Font = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Borders", "Item"); !ok || got.ReturnType != "Excel.Border" {
		t.Fatalf("Borders.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Application", "WorksheetFunction"); !ok || got.ReturnType != "Excel.WorksheetFunction" {
		t.Fatalf("Application.WorksheetFunction = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Collection", "Count"); !ok || got.ReturnType != "Long" {
		t.Fatalf("Collection.Count = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Collection", "Item"); !ok || got.ReturnType != "Variant" {
		t.Fatalf("Collection.Item = %+v, %v", got, ok)
	}
}

func TestProgIDsListPreservesDisplayNames(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	progIDs := db.ProgIDsList()
	if !hasString(progIDs, "Scripting.Dictionary") {
		t.Fatalf("ProgIDsList should include canonical Scripting.Dictionary, got %+v", progIDs)
	}
	if hasString(progIDs, "scripting.dictionary") {
		t.Fatalf("ProgIDsList should not expose folded ProgID names: %+v", progIDs)
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
	constant, ok = db.ResolveConstant("xlLandscape")
	if !ok {
		t.Fatal("xlLandscape was not loaded")
	}
	if constant.EnumGroup != "XlPageOrientation" {
		t.Fatalf("unexpected xlLandscape metadata: %+v", constant)
	}
}

func TestCompletionListsExposeGlobalsConstantsAndMembers(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	if !hasGlobal(db.GlobalsList(), "ThisWorkbook", "Excel.Workbook") {
		t.Fatalf("ThisWorkbook global missing: %+v", db.GlobalsList())
	}
	if !hasConstant(db.ConstantsList(), "xlUp") {
		t.Fatal("xlUp constant missing")
	}
	if !hasConstant(db.ConstantsList(), "xlThin") {
		t.Fatal("xlThin constant missing")
	}
	if !hasMember(db.Members("Excel.Range"), "Value") {
		t.Fatalf("Range.Value member missing: %+v", db.Members("Excel.Range"))
	}
	if !hasMember(db.Members("Excel.Range"), "Font") {
		t.Fatalf("Range.Font member missing: %+v", db.Members("Excel.Range"))
	}
	if !hasMember(db.Members("Excel.Font"), "Color") {
		t.Fatalf("Font.Color member missing: %+v", db.Members("Excel.Font"))
	}
}

func hasGlobal(items []MemberInfo, name, typ string) bool {
	for _, item := range items {
		if item.Name == name && item.ReturnType == typ {
			return true
		}
	}
	return false
}

func hasConstant(items []ConstantInfo, name string) bool {
	for _, item := range items {
		if item.Name == name {
			return true
		}
	}
	return false
}

func hasMember(items []MemberInfo, name string) bool {
	for _, item := range items {
		if item.Name == name {
			return true
		}
	}
	return false
}

func hasString(items []string, name string) bool {
	for _, item := range items {
		if item == name {
			return true
		}
	}
	return false
}
