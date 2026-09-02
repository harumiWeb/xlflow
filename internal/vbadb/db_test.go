package vbadb

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

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
	for _, enumName := range []string{"VbAppWinStyle", "VbVarType", "XlCalculation"} {
		if typ, ok := db.ResolveType(enumName); !ok || typ.Kind != "enum" {
			t.Fatalf("ResolveType(%s) = %+v, %v; want embedded enum type", enumName, typ, ok)
		}
	}
	for name, want := range map[string]string{
		"IUnknown":        "stdole.IUnknown",
		"stdole.IUnknown": "stdole.IUnknown",
		"IEnumVARIANT":    "stdole.IEnumVARIANT",
	} {
		if typ, ok := db.ResolveType(name); !ok || typ.Name != want {
			t.Fatalf("ResolveType(%s) = %+v, %v; want %s", name, typ, ok, want)
		}
	}
	if typ, ok := db.ResolveProgID("Scripting.Dictionary"); !ok || typ.Name != "Scripting.Dictionary" {
		t.Fatalf("ResolveProgID(Scripting.Dictionary) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("WScript.Shell"); !ok || typ.Name != "WScript.Shell" {
		t.Fatalf("ResolveProgID(WScript.Shell) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("VBScript.RegExp"); !ok || typ.Name != "VBScript.RegExp" {
		t.Fatalf("ResolveProgID(VBScript.RegExp) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("ADODB.Connection"); !ok || typ.Name != "ADODB.Connection" {
		t.Fatalf("ResolveProgID(ADODB.Connection) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("ADODB.Recordset"); !ok || typ.Name != "ADODB.Recordset" {
		t.Fatalf("ResolveProgID(ADODB.Recordset) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("ADODB.Command"); !ok || typ.Name != "ADODB.Command" {
		t.Fatalf("ResolveProgID(ADODB.Command) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveProgID("Excel.Application"); !ok || typ.Name != "Excel.Application" {
		t.Fatalf("ResolveProgID(Excel.Application) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveType("Collection"); !ok || typ.Name != "VBA.Collection" {
		t.Fatalf("ResolveType(Collection) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveType("FileDialog"); !ok || typ.Name != "Office.FileDialog" {
		t.Fatalf("ResolveType(FileDialog) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveType("Office.FileDialogSelectedItems"); !ok || typ.Name != "Office.FileDialogSelectedItems" {
		t.Fatalf("ResolveType(Office.FileDialogSelectedItems) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveGlobal("Worksheets"); !ok || typ.Name != "Excel.Worksheets" {
		t.Fatalf("ResolveGlobal(Worksheets) = %+v, %v", typ, ok)
	}
	if typ, ok := db.ResolveGlobal("Err"); !ok || typ.Name != "VBA.ErrObject" {
		t.Fatalf("ResolveGlobal(Err) = %+v, %v", typ, ok)
	}
}

func TestMergeJSONPreservesParamArrayParameters(t *testing.T) {
	db := New()
	if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Test.Logger",
    "methods": [{
      "name": "Log",
      "parameters": [
        { "name": "Level", "type": "String" },
        { "name": "Parts", "type": "Variant", "optional": true, "param_array": true }
      ]
    }]
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	member, ok := db.ResolveMember("Test.Logger", "Log")
	if !ok || len(member.Parameters) != 2 {
		t.Fatalf("Log member = %+v, %v", member, ok)
	}
	if param := member.Parameters[1]; !param.Optional || !param.ParamArray {
		t.Fatalf("ParamArray metadata not preserved: %+v", param)
	}
}

func TestCuratedOverlayPreservesGeneratedTypeLibProvenance(t *testing.T) {
	db := New()
	if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.Worksheet",
    "library": "Excel",
    "kind": "interface",
    "confidence": "generated",
    "source": "typelib",
    "properties": [{ "name": "GeneratedMember", "return_type": "String" }]
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.Worksheet",
    "kind": "class",
    "confidence": "curated",
    "source": "xlflow",
    "summary": "Curated worksheet summary.",
    "properties": [{ "name": "CuratedMember", "return_type": "Long" }]
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	typ, ok := db.ResolveType("Excel.Worksheet")
	if !ok {
		t.Fatal("Excel.Worksheet missing after merge")
	}
	if typ.Source != "typelib" || typ.Confidence != "generated" {
		t.Fatalf("generated provenance was downgraded: %+v", typ)
	}
	if typ.Kind != "class" || typ.Summary != "Curated worksheet summary." {
		t.Fatalf("curated overlay metadata was not retained: %+v", typ)
	}
	for _, name := range []string{"GeneratedMember", "CuratedMember"} {
		if _, ok := db.ResolveMember("Excel.Worksheet", name); !ok {
			t.Fatalf("merged member %s missing", name)
		}
	}
}

func TestCuratedOverlayPreservesGeneratedMemberSignature(t *testing.T) {
	db := New()
	if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.Application",
    "confidence": "generated",
    "source": "typelib",
    "methods": [{
      "name": "Run",
      "parameters": [
        { "name": "Macro", "type": "Variant" },
        { "name": "Arg1", "type": "Variant", "optional": true },
        { "name": "Arg2", "type": "Variant", "optional": true },
        { "name": "Arg3", "type": "Variant", "optional": true },
        { "name": "Arg4", "type": "Variant", "optional": true },
        { "name": "Arg5", "type": "Variant", "optional": true },
        { "name": "Arg6", "type": "Variant", "optional": true },
        { "name": "Arg7", "type": "Variant", "optional": true }
      ]
    }]
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.Application",
    "confidence": "curated",
    "source": "xlflow",
    "methods": [{
      "name": "Run",
      "parameters": [
        { "name": "Macro", "type": "Variant" },
        { "name": "Arg1", "type": "Variant", "optional": true }
      ]
    }]
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	member, ok := db.ResolveMember("Excel.Application", "Run")
	if !ok || len(member.Parameters) != 8 || member.Parameters[7].Name != "Arg7" {
		t.Fatalf("generated Application.Run signature was shortened by curated overlay: %+v, %v", member, ok)
	}
}

func TestBuiltinApplicationRunSupportsThirtyArguments(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	member, ok := db.ResolveMember("Excel.Application", "Run")
	if !ok || len(member.Parameters) != 31 {
		t.Fatalf("Application.Run parameters = %+v, %v; want Macro plus Arg1..Arg30", member.Parameters, ok)
	}
	if !strings.EqualFold(member.Parameters[0].Name, "Macro") || member.Parameters[0].Optional {
		t.Fatalf("Application.Run Macro parameter = %+v; want required Macro", member.Parameters[0])
	}
	for i, parameter := range member.Parameters[1:] {
		if parameter.Name != fmt.Sprintf("Arg%d", i+1) || !parameter.Optional {
			t.Fatalf("Application.Run parameter %d = %+v; want optional Arg%d", i+1, parameter, i+1)
		}
	}
}

func TestIsAssignableUsesExplicitRelationshipsOnly(t *testing.T) {
	db := New()
	if err := db.MergeJSON([]byte(`{
  "types": [
    { "name": "Test.Control" },
    { "name": "Test.TextBox", "assignable_to": ["Test.Control"] },
    { "name": "Test.Worksheet" },
    { "name": "Test.Workbook" }
  ]
}`)); err != nil {
		t.Fatal(err)
	}
	if assignable, known := db.IsAssignable("Test.Control", "Test.TextBox"); !known || !assignable {
		t.Fatalf("TextBox should be assignable to Control: assignable=%v known=%v", assignable, known)
	}
	if assignable, known := db.IsAssignable("Test.Worksheet", "Test.TextBox"); !known || assignable {
		t.Fatalf("TextBox to Worksheet should be known incompatible: assignable=%v known=%v", assignable, known)
	}
	if assignable, known := db.IsAssignable("Test.Worksheet", "Test.Workbook"); known || assignable {
		t.Fatalf("Workbook to Worksheet should be unknown without relationship metadata: assignable=%v known=%v", assignable, known)
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
	if got, ok := db.ResolveMember("Excel.Range", "SpecialCells"); !ok || got.ReturnType != "Excel.Range" || len(got.Parameters) != 2 || got.Parameters[0].Name != "Type" || got.Parameters[0].Optional || !got.Parameters[1].Optional {
		t.Fatalf("Range.SpecialCells parameters = %+v, %v", got, ok)
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
	if got, ok := db.ResolveMember("ADODB.Command", "CommandText"); !ok || got.ReturnType != "String" {
		t.Fatalf("Command.CommandText = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("ADODB.Command", "CreateParameter"); !ok || got.ReturnType != "ADODB.Parameter" || len(got.Parameters) != 5 {
		t.Fatalf("Command.CreateParameter = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Global", "MsgBox"); !ok || len(got.Parameters) != 5 || got.ReturnType != "VbMsgBoxResult" {
		t.Fatalf("VBA.Global.MsgBox parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Global", "IsObject"); !ok || len(got.Parameters) != 1 || got.ReturnType != "Boolean" {
		t.Fatalf("VBA.Global.IsObject parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Global", "CVErr"); !ok || got.ReturnType != "Variant" || len(got.Parameters) != 1 || got.Parameters[0].Name != "ErrorNumber" || got.Parameters[0].Type != "Long" {
		t.Fatalf("VBA.Global.CVErr = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.Global", "GetObject"); !ok || got.ReturnType != "Object" {
		t.Fatalf("VBA.Global.GetObject = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.ErrObject", "Raise"); !ok || len(got.Parameters) != 5 || got.Parameters[0].Name != "Number" {
		t.Fatalf("Err.Raise parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("VBA.ErrObject", "Description"); !ok || got.ReturnType != "String" {
		t.Fatalf("Err.Description = %+v, %v", got, ok)
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
	for _, name := range []string{"Match", "VLookup", "XLookup"} {
		got, ok := db.ResolveMember("Excel.Application", name)
		if !ok || got.ReturnType != "Variant" {
			t.Fatalf("Application.%s = %+v, %v", name, got, ok)
		}
	}
	for _, name := range []string{"Min", "Max"} {
		got, ok := db.ResolveMember("Excel.Application", name)
		if !ok || got.ReturnType != "Double" || len(got.Parameters) != 30 || got.Parameters[0].Optional || !got.Parameters[29].Optional {
			t.Fatalf("Application.%s = %+v, %v", name, got, ok)
		}
		for i, parameter := range got.Parameters {
			wantName := fmt.Sprintf("Arg%d", i+1)
			wantOptional := i > 0
			if parameter.Name != wantName || parameter.Type != "Variant" || parameter.Optional != wantOptional {
				t.Fatalf("Application.%s parameter %d = %+v; want %s Variant optional=%t", name, i+1, parameter, wantName, wantOptional)
			}
		}
	}
	if got, ok := db.ResolveMember("Excel.Application", "Match"); !ok || len(got.Parameters) != 3 || !got.Parameters[2].Optional {
		t.Fatalf("Application.Match parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Application", "VLookup"); !ok || len(got.Parameters) != 4 || !got.Parameters[3].Optional {
		t.Fatalf("Application.VLookup parameters = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Excel.Application", "XLookup"); !ok || len(got.Parameters) != 6 || !got.Parameters[3].Optional || !got.Parameters[4].Optional || !got.Parameters[5].Optional {
		t.Fatalf("Application.XLookup parameters = %+v, %v", got, ok)
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
	if got, ok := db.ResolveMember("Scripting.Dictionary", "CompareMode"); !ok || got.ReturnType != "Scripting.CompareMethod" || got.ReadOnly {
		t.Fatalf("Dictionary.CompareMode = %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("MSForms.Controls", "Item"); !ok || got.ReturnType != "MSForms.Control" || len(got.Parameters) != 1 {
		t.Fatalf("Controls.Item = %+v, %v", got, ok)
	}
	for _, name := range []string{"Sum", "Average", "CountA"} {
		if got, ok := db.ResolveMember("Excel.WorksheetFunction", name); !ok || len(got.Parameters) != 30 || !got.Parameters[29].Optional {
			t.Fatalf("WorksheetFunction.%s parameters = %+v, %v", name, got, ok)
		}
	}
	if got, ok := db.ResolveMember("Excel.WorksheetFunction", "XLookup"); !ok || got.ReturnType != "Variant" || len(got.Parameters) != 6 || !got.Parameters[3].Optional || !got.Parameters[4].Optional || !got.Parameters[5].Optional {
		t.Fatalf("WorksheetFunction.XLookup parameters = %+v, %v", got, ok)
	}
}

func TestBuiltinVBAStandardLibraryCoverage(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	expectedFunctions := []string{
		"Abs", "AppActivate", "Array", "Asc", "AscB", "AscW", "Atn", "Beep",
		"CBool", "CCur", "CDate", "CDbl", "CDec", "Choose", "Chr", "ChrB", "ChrW",
		"CInt", "CLng", "Command", "Cos", "CreateObject", "CStr", "CurDir", "CVar", "CVDate",
		"Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "DDB",
		"Dir", "DoEvents", "Environ", "EOF", "Error", "Exp", "FileAttr", "FileCopy",
		"FileDateTime", "FileLen", "Filter", "Fix", "Format", "FormatCurrency",
		"FormatDateTime", "FormatNumber", "FormatPercent", "FreeFile", "FV",
		"GetAllSettings", "GetAttr", "GetObject", "GetSetting", "Hex", "Hour",
		"IIf", "IMEStatus", "InStr", "InStrRev", "InputBox", "Int", "IPmt", "IRR",
		"IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull", "IsNumeric", "IsObject",
		"Join", "Kill", "LBound", "LCase", "Left", "LeftB", "Len", "LenB", "Loc", "LOF",
		"Log", "LTrim", "Mid", "MidB", "Minute", "MIRR", "MkDir", "Month", "MonthName",
		"MsgBox", "NPer", "Now", "NPV", "Oct", "Partition", "Pmt", "PPmt", "PV",
		"QBColor", "Randomize", "Rate", "Replace", "Reset", "RGB", "Right", "RightB",
		"RmDir", "Rnd", "Round", "RTrim", "SaveSetting", "Second", "Seek", "SendKeys",
		"SetAttr", "Sgn", "Shell", "Sin", "SLN", "Space", "Split", "Sqr", "Str",
		"StrComp", "StrConv", "String", "Switch", "SYD", "Tan", "Time", "Timer",
		"TimeSerial", "TimeValue", "Trim", "TypeName", "UBound", "UCase", "VarType",
		"Weekday", "WeekdayName", "Year",
	}
	for _, name := range expectedFunctions {
		if _, ok := db.ResolveMember("VBA.Global", name); !ok {
			t.Fatalf("VBA.Global.%s missing from built-in DB", name)
		}
	}

	expectedConstants := []string{
		"vbCr", "vbCrLf", "vbLf", "vbNewLine", "vbTab", "vbBack", "vbFormFeed",
		"vbVerticalTab", "vbNullChar", "vbNullString", "vbObjectError",
		"vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel", "vbYesNo",
		"vbRetryCancel", "vbCritical", "vbQuestion", "vbExclamation", "vbInformation",
		"vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", "vbDefaultButton4",
		"vbApplicationModal", "vbSystemModal", "vbOK", "vbCancel", "vbAbort", "vbRetry",
		"vbIgnore", "vbYes", "vbNo", "vbUseCompareOption", "vbBinaryCompare",
		"vbTextCompare", "vbDatabaseCompare", "vbGeneralDate", "vbLongDate",
		"vbShortDate", "vbLongTime", "vbShortTime", "vbUseDefault", "vbTrue", "vbFalse",
		"vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble",
		"vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
		"vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbLongLong",
		"vbUserDefinedType", "vbArray", "vbNormal", "vbReadOnly", "vbHidden",
		"vbSystem", "vbVolume", "vbDirectory", "vbArchive", "vbAlias", "vbHide",
		"vbNormalFocus", "vbMinimizedFocus", "vbMaximizedFocus", "vbNormalNoFocus",
		"vbMinimizedNoFocus",
	}
	for _, name := range expectedConstants {
		if _, ok := db.ResolveConstant(name); !ok {
			t.Fatalf("%s missing from built-in constants", name)
		}
	}
}

func TestBuiltinArrayUsesOptionalParamArray(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	array, ok := db.ResolveMember("VBA.Global", "Array")
	if !ok || len(array.Parameters) != 1 {
		t.Fatalf("VBA.Global.Array = %+v, %v", array, ok)
	}
	if param := array.Parameters[0]; !param.Optional || !param.ParamArray {
		t.Fatalf("Array parameter = %+v, want optional ParamArray", param)
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

func TestResolveConstantsPreservesAmbiguousEnumCandidates(t *testing.T) {
	db := New()
	if err := db.MergeJSON([]byte(`{
  "constants": [
    {"name":"Ready","library":"ZLib","enum_group":"StateZ","value":"1"},
    {"name":"ready","library":"zlib","enum_group":"statez","value":"2"},
    {"name":"Ready","library":"ALib","enum_group":"StateA"}
  ]
}`)); err != nil {
		t.Fatal(err)
	}
	candidates := db.ResolveConstants("ready")
	if len(candidates) != 2 || candidates[0].Library != "ALib" || candidates[1].Library != "ZLib" {
		t.Fatalf("ResolveConstants = %#v, want deterministic candidates", candidates)
	}
	if winner, ok := db.ResolveConstant("ready"); !ok || (winner.Library != "ALib" && winner.Library != "ZLib") {
		t.Fatalf("ResolveConstant compatibility winner = %#v, %v", winner, ok)
	}
	all := db.AllConstantsList()
	if len(all) != 2 {
		t.Fatalf("AllConstantsList = %#v, want duplicate identity folded", all)
	}
	for _, constant := range all {
		if strings.EqualFold(constant.Library, "zlib") && strings.EqualFold(constant.EnumGroup, "statez") && constant.Value != "1" {
			t.Fatalf("duplicate constant winner = %#v, want deterministic first candidate", constant)
		}
	}
}

func TestBuiltinScriptingCompareMethodMetadataDoesNotReplaceVBAConstants(t *testing.T) {
	db, err := LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}

	if typ, ok := db.ResolveType("Scripting.CompareMethod"); !ok || typ.Kind != "enum" || typ.Library != "Scripting" {
		t.Fatalf("ResolveType(Scripting.CompareMethod) = %+v, %v", typ, ok)
	}
	for scriptingName, vbaName := range map[string]string{
		"BinaryCompare":   "vbBinaryCompare",
		"TextCompare":     "vbTextCompare",
		"DatabaseCompare": "vbDatabaseCompare",
	} {
		scriptingConstant, ok := db.ResolveConstant(scriptingName)
		if !ok || scriptingConstant.Library != "Scripting" || scriptingConstant.EnumGroup != "CompareMethod" {
			t.Fatalf("ResolveConstant(%s) = %+v, %v", scriptingName, scriptingConstant, ok)
		}
		vbaConstant, ok := db.ResolveConstant(vbaName)
		if !ok || vbaConstant.Library != "VBA" || vbaConstant.EnumGroup != "VbCompareMethod" {
			t.Fatalf("ResolveConstant(%s) = %+v, %v", vbaName, vbaConstant, ok)
		}
		if scriptingConstant.Value != vbaConstant.Value {
			t.Fatalf("%s value = %s, %s value = %s", scriptingName, scriptingConstant.Value, vbaName, vbaConstant.Value)
		}
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
	if !hasGlobal(db.GlobalsList(), "Err", "VBA.ErrObject") {
		t.Fatalf("Err global missing: %+v", db.GlobalsList())
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

func TestLoadFilesMergesTypesMembersAndOverridesMembers(t *testing.T) {
	dir := t.TempDir()
	base := filepath.Join(dir, "base.json")
	overlay := filepath.Join(dir, "overlay.json")
	if err := os.WriteFile(base, []byte(`{
  "types": [
    {
      "name": "Excel.Range",
      "library": "Excel",
      "kind": "class",
      "properties": [
        { "name": "Value", "return_type": "Variant" }
      ],
      "methods": [
        { "name": "Find", "return_type": "Object" }
      ]
    }
  ],
  "constants": [
    { "name": "xlUp", "library": "Excel", "enum_group": "XlDirection" }
  ]
}`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(overlay, []byte(`{
  "types": [
    {
      "name": "Excel.Range",
      "source": "xlflow",
      "properties": [
        { "name": "Font", "return_type": "Excel.Font" }
      ],
      "methods": [
        { "name": "Find", "return_type": "Excel.Range", "parameters": [{ "name": "What", "type": "Variant" }] }
      ]
    }
  ],
  "constants": [
    { "name": "xlUp", "library": "Excel", "enum_group": "XlDirection", "summary": "overlay" }
  ]
}`), 0o644); err != nil {
		t.Fatal(err)
	}

	db, err := LoadFiles(base, overlay)
	if err != nil {
		t.Fatal(err)
	}
	if got, ok := db.ResolveMember("Range", "Value"); !ok || got.ReturnType != "Variant" {
		t.Fatalf("base property missing after merge: %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Range", "Font"); !ok || got.ReturnType != "Excel.Font" {
		t.Fatalf("overlay property missing after merge: %+v, %v", got, ok)
	}
	if got, ok := db.ResolveMember("Range", "Find"); !ok || got.ReturnType != "Excel.Range" || len(got.Parameters) != 1 {
		t.Fatalf("overlay method should replace base method: %+v, %v", got, ok)
	}
	if c, ok := db.ResolveConstant("xlUp"); !ok || c.Summary != "overlay" {
		t.Fatalf("constant overlay missing: %+v, %v", c, ok)
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
