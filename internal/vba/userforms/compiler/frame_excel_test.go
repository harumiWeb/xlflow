package compiler

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// TestGenerateExcelFrameArtifact only writes files. Excel execution belongs to
// the developer-only frame harness, never ordinary tests or CI.
func TestGenerateExcelFrameArtifact(t *testing.T) {
	input, output, expected := os.Getenv("XLFLOW_FRAME_WORKBOOK"), os.Getenv("XLFLOW_FRAME_OUTPUT"), os.Getenv("XLFLOW_FRAME_EXPECTED")
	if input == "" || output == "" || expected == "" {
		t.Skip("set XLFLOW_FRAME_WORKBOOK/OUTPUT/EXPECTED for the local gate")
	}
	for _, path := range []string{output, expected} {
		if strings.EqualFold(filepath.Clean(path), filepath.Clean(input)) {
			t.Fatal("outputs must not overwrite input")
		}
		if _, err := os.Stat(path); !os.IsNotExist(err) {
			t.Fatalf("output exists or cannot be checked: %s", path)
		}
	}
	if strings.EqualFold(filepath.Clean(output), filepath.Clean(expected)) {
		t.Fatal("outputs must be distinct")
	}
	reader, err := zip.OpenReader(input)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := reader.Close(); err != nil {
			t.Error(err)
		}
	})
	var data []byte
	for _, entry := range reader.File {
		if entry.Name == "xl/vbaProject.bin" {
			data = readZipEntry(t, entry)
		}
	}
	project, err := vbaproject.Read(data)
	if err != nil {
		t.Fatal(err)
	}
	var base *oforms.Form
	index := -1
	for i, form := range project.Forms {
		if form.Name == "FrameTopologyForm" {
			base, index = form, i
		}
	}
	if base == nil {
		t.Fatal("baseline FrameTopologyForm missing")
		return
	}
	t.Logf("VBFrame source: %s", base.DesignerSource.Text)
	for _, level := range base.Levels {
		t.Logf("Excel baseline level=%s nextID=%d shapeCookie=%d", level.Path, level.Record.Values["NextAvailableID"], level.Record.Values["ShapeCookie"])
		for _, control := range level.Controls {
			t.Logf("site=%s id=%d cache=%d bytes=%d", control.Name, control.ID, control.CLSIDCacheIndex, control.ObjectStreamSize)
		}
	}
	desired := spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", CoordinateSystem: "parent-relative", Form: spec.FormSpecForm{Name: base.Name, Caption: new("issue884-frame-topology")}, Controls: []spec.FormSpecControl{
		{Name: "RootCommon", Type: "Label", Caption: new("root-common"), Left: new(12.0), Top: new(10.0), Width: new(90.0), Height: new(18.0)},
		{Name: "EmptyFrame", Type: "Frame", Caption: new("empty-frame"), Left: new(12.0), Top: new(40.0), Width: new(100.0), Height: new(58.0)},
		{Name: "ParentFrame", Type: "Frame", Caption: new("parent-frame"), Left: new(130.0), Top: new(40.0), Width: new(180.0), Height: new(136.0), Controls: []spec.FormSpecControl{
			{Name: "SiblingText", Type: "TextBox", Text: new("sibling-text"), Left: new(10.0), Top: new(22.0), Width: new(84.0), Height: new(18.0)},
			{Name: "NestedFrame", Type: "Frame", Caption: new("nested-frame"), Left: new(12.0), Top: new(54.0), Width: new(144.0), Height: new(64.0), Controls: []spec.FormSpecControl{
				{Name: "NestedLabel", Type: "Label", Caption: new("nested-label"), Left: new(8.0), Top: new(20.0), Width: new(108.0), Height: new(18.0)},
			}},
		}},
		{Name: "RootText", Type: "TextBox", Text: new("root-text"), Left: new(12.0), Top: new(202.0), Width: new(180.0), Height: new(20.0)},
	}}
	var generated *oforms.Form
	if os.Getenv("XLFLOW_FRAME_MODE") == "template" {
		// Reorder, reparent, delete, replace a Frame subtree and add a new Frame.
		desired.Controls = []spec.FormSpecControl{
			{Name: "ParentFrame", Type: "Frame", TabIndex: new(0), Controls: []spec.FormSpecControl{
				{Name: "NestedFrame", Type: "Frame", TabIndex: new(0), Controls: []spec.FormSpecControl{{Name: "RootText", Type: "TextBox", TabIndex: new(0)}}},
				{Name: "SiblingText", Type: "TextBox", TabIndex: new(1)},
			}},
			{Name: "RootCommon", Type: "Label", TabIndex: new(1)},
			{Name: "EmptyFrame", Type: "Label", Caption: new("replaced-frame")},
			{Name: "AddedFrame", Type: "Frame", Controls: []spec.FormSpecControl{{Name: "AddedLabel", Type: "Label", Caption: new("added")}}},
		}
		generated, err = CompileTemplate(base, desired, project.Props.CodePage)
		if experiment := os.Getenv("XLFLOW_FRAME_CASE"); experiment != "" {
			desired, err = projection.Project(base)
			if err != nil {
				t.Fatal(err)
			}
			switch experiment {
			case "noop":
			case "reorder":
				desired.Controls[0].ZIndex = new(2)
				desired.Controls[2].ZIndex = new(0)
			case "move":
				desired.Controls[6].ParentID = desired.Controls[4].ID
				desired.Controls[6].ZIndex = new(1)
			case "remove":
				desired.Controls = append(desired.Controls[:5], desired.Controls[6:]...)
			case "add":
				desired.Controls = append(desired.Controls, spec.FormSpecControl{Name: "DiffLabel", Type: "Label", ID: "new-label", ParentID: desired.Controls[4].ID, ZIndex: new(1)})
			default:
				t.Fatalf("unknown experiment %q", experiment)
			}
			generated, err = CompileTemplate(base, desired, project.Props.CodePage)
		}
	} else {
		desired.Form.Build = &spec.FormSpecBuildForm{ClientWidth: new(324.0), ClientHeight: new(240.0)}
		generated, err = CompileNew(desired, project.Props.CodePage)
	}
	if err != nil {
		t.Fatal(err)
	}
	project.Forms[index] = generated
	// Invoke a known standard-module entrypoint and assert the entire runtime
	// hierarchy. Parent filtering avoids COM Controls' descendant enumeration.
	for i := range project.Modules {
		if project.Modules[i].Name == "Main" {
			project.Modules[i].Source = `Attribute VB_Name = "Main"
Option Explicit
Private Function DumpFrameControls(ByVal controls As Object, ByVal parentName As String) As String
    Dim control As Object, index As Long
    For Each control In controls
        If CStr(control.Parent.Name) = parentName Then
            DumpFrameControls = DumpFrameControls & parentName & "/" & CStr(index) & ":" & control.Name & ":" & TypeName(control) & ";"
            If TypeName(control) = "Frame" Then DumpFrameControls = DumpFrameControls & DumpFrameControls(control.Controls, control.Name)
            index = index + 1
        End If
    Next control
End Function
Public Sub RunFrameSentinel()
    On Error GoTo Failed
    Dim form As Object
    Set form = VBA.UserForms.Add("FrameTopologyForm")
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-884-ok|" & DumpFrameControls(form.Controls, form.Name)
    Unload form
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-884-failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
`
		}
	}
	data, err = vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for _, entry := range reader.File {
		if entry.Name != "xl/vbaProject.bin" {
			if err := writer.Copy(entry); err != nil {
				t.Fatal(err)
			}
			continue
		}
		stream, err := writer.CreateHeader(&entry.FileHeader)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := stream.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(output, buffer.Bytes(), 0o600); err != nil {
		t.Fatal(err)
	}
	result, err := projection.Project(generated)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(expected, append(encoded, '\n'), 0o600); err != nil {
		t.Fatal(err)
	}
}
