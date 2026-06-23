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
	if typ, ok := db.ResolveProgID("ADODB.Connection"); !ok || typ.Name != "ADODB.Connection" {
		t.Fatalf("ResolveProgID(ADODB.Connection) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("ADODB.Recordset"); !ok || typ.Name != "ADODB.Recordset" {
		t.Fatalf("ResolveProgID(ADODB.Recordset) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("Excel.Application"); !ok || typ.Name != "Excel.Application" {
		t.Fatalf("ResolveProgID(Excel.Application) = %+v, %v", typ, ok)
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

	if got, ok := db.ResolveMember("Excel.Worksheets", "Item"); !ok || got.ReturnType != "Excel.Worksheet" || len(got.Parameters) != 1 {
		t.Fatalf("Worksheets.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Workbooks", "Item"); !ok || got.ReturnType != "Excel.Workbook" || len(got.Parameters) != 1 {
		t.Fatalf("Workbooks.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Workbooks", "Open"); !ok || got.ReturnType != "Excel.Workbook" {
		t.Fatalf("Workbooks.Open = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Workbooks", "Open"); !ok || len(got.Parameters) != 15 || got.Parameters[0].Name != "Filename" || !got.Parameters[14].Optional {
		t.Fatalf("Workbooks.Open parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Range", "Find"); !ok || len(got.Parameters) != 9 || got.Parameters[0].Name != "What" || got.ReturnType != "Excel.Range" {
		t.Fatalf("Range.Find parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Range", "Cells"); !ok || len(got.Parameters) != 2 || got.ReturnType != "Excel.Range" {
		t.Fatalf("Range.Cells parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Scripting.FileSystemObject", "OpenTextFile"); !ok || len(got.Parameters) != 4 || got.Parameters[0].Name != "Filename" || got.ReturnType != "Scripting.TextStream" {
		t.Fatalf("FileSystemObject.OpenTextFile parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("ADODB.Recordset", "Fields"); !ok || got.ReturnType != "ADODB.Fields" {
		t.Fatalf("Recordset.Fields = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Global", "MsgBox"); !ok || len(got.Parameters) != 5 || got.ReturnType != "VbMsgBoxResult" {
		t.Fatalf("VBA.Global.MsgBox parameters = %+v, %v", got, ok)
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
	if got, ok := db.ResolveMember("Excel.ListObjects", "Item"); !ok || got.ReturnType != "Excel.ListObject" || len(got.Parameters) != 1 {
		t.Fatalf("ListObjects.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.ListObject", "DataBodyRange"); !ok || got.ReturnType != "Excel.Range" {
		t.Fatalf("ListObject.DataBodyRange = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.ListColumns", "Item"); !ok || got.ReturnType != "Excel.ListColumn" || len(got.Parameters) != 1 {
		t.Fatalf("ListColumns.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.ListColumn", "DataBodyRange"); !ok || got.ReturnType != "Excel.Range" {
		t.Fatalf("ListColumn.DataBodyRange = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.ListRows", "Item"); !ok || got.ReturnType != "Excel.ListRow" {
		t.Fatalf("ListRows.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Worksheet", "PivotTables"); !ok || got.ReturnType != "Excel.PivotTables" {
		t.Fatalf("Worksheet.PivotTables = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.PivotTables", "Item"); !ok || got.ReturnType != "Excel.PivotTable" || len(got.Parameters) != 1 {
		t.Fatalf("PivotTables.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.PivotFields", "Item"); !ok || got.ReturnType != "Excel.PivotField" || len(got.Parameters) != 1 {
		t.Fatalf("PivotFields.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Worksheet", "Shapes"); !ok || got.ReturnType != "Excel.Shapes" {
		t.Fatalf("Worksheet.Shapes = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Shapes", "Item"); !ok || got.ReturnType != "Excel.Shape" || len(got.Parameters) != 1 {
		t.Fatalf("Shapes.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Shape", "TextFrame"); !ok || got.ReturnType != "Excel.TextFrame" {
		t.Fatalf("Shape.TextFrame = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Collection", "Count"); !ok || got.ReturnType != "Long" {
		t.Fatalf("Collection.Count = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Collection", "Item"); !ok || got.ReturnType != "Variant" {
		t.Fatalf("Collection.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Scripting.Dictionary", "Item"); !ok || got.ReturnType != "Variant" || !got.Default || len(got.Parameters) != 1 {
		t.Fatalf("Dictionary.Item = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("MSForms.Controls", "Item"); !ok || got.ReturnType != "MSForms.Control" || len(got.Parameters) != 1 {
		t.Fatalf("Controls.Item = %+v, %v", got, ok)
	}
	for _, name := range []string{"Sum", "Average", "CountA"} {
		if got, ok := db.ResolveMember("Excel.WorksheetFunction", name); !ok || len(got.Parameters) != 30 || !got.Parameters[29].Optional {
			t.Fatalf("WorksheetFunction.%s parameters = %+v, %v", name, got, ok)
		}
	}
}

func TestResolveMemberIncludesEvents(t *testing.T) {
	db := &DB{Types: map[string]TypeInfo{}, Aliases: map[string]string{}}
	db.addType(TypeInfo{
		Name:   "Test.Widget",
		Events: []MemberInfo{{Name: "Changed", ReturnType: "Void"}},
	})

	if got, ok := db.ResolveMember("Test.Widget", "Changed"); !ok || got.Name != "Changed" {
		t.Fatalf("ResolveMember event = %+v, %v", got, ok)
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
	for _, want := range []string{"ADODB.Connection", "ADODB.Recordset", "Excel.Application"} {
		if !hasString(progIDs, want) {
			t.Fatalf("ProgIDsList should include %s, got %+v", want, progIDs)
		}
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
