package intel

import (
	"context"
	"path/filepath"
	"strings"
	"testing"
)

func TestIssue774OutlookEarlyBoundTypesResolve(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "ReproOutlook.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "ReproOutlook"
Option Explicit
Public Sub ReproVBA229()
    Dim outlookObj As Outlook.Application
    Dim mailItemObj As Outlook.MailItem
    Set outlookObj = CreateObject("Outlook.Application")
    Set mailItemObj = outlookObj.CreateItem(0)
    mailItemObj.Subject = "xlflow VBA229 reproduction"
End Sub
`,
	}

	if diagnostics := diagnosticsByCode(analyzer.LocalTypeNameDiagnosticsContext(context.Background(), doc), "VBA229"); len(diagnostics) != 0 {
		t.Fatalf("valid Outlook early-bound types should not produce VBA229: %+v", diagnostics)
	}
	if diagnostics := analyzer.Diagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("the Outlook VBA229 reproduction should be diagnostic-free: %+v", diagnostics)
	}

	appointmentDoc := doc
	appointmentDoc.Source = `Attribute VB_Name = "ReproOutlook"
Option Explicit
Public Sub ReproVBA229()
    Dim outlookObj As Outlook.Application
    Dim appointment As Outlook.AppointmentItem
    Set outlookObj = CreateObject("Outlook.Application")
    Set appointment = outlookObj.CreateItem(1)
    appointment.Subject = "xlflow Outlook appointment"
End Sub
`
	appointmentDiagnostics := analyzer.Diagnostics(appointmentDoc)
	if diagnostics := diagnosticsByCode(appointmentDiagnostics, "VBA229"); len(diagnostics) != 0 {
		t.Fatalf("valid Outlook appointment type should not produce VBA229: %+v", diagnostics)
	}
	if diagnostics := diagnosticsByCode(appointmentDiagnostics, "VB038"); len(diagnostics) != 0 {
		t.Fatalf("CreateItem's Object return should remain assignable to Outlook item types: %+v", diagnostics)
	}

	shadowDoc := Document{
		Path:       filepath.Join(t.TempDir(), "ShadowedMailItem.bas"),
		ModuleKind: "standard",
		Source: `Option Explicit
Public Sub Probe()
    Dim mailItem As String
    Debug.Print mailItem
End Sub
`,
	}
	offset := strings.Index(shadowDoc.Source, "Debug.Print mailItem") + len("Debug.Print ")
	if got, ok := analyzer.resolveDocumentExpressionTypeAt(shadowDoc, "mailItem", offset); !ok || got != "String" {
		t.Fatalf("local MailItem shadowing = %q, %v; want String, true", got, ok)
	}
	if got, ok := analyzer.typeDiagnosticBaseType(shadowDoc, "mailItem", offset); !ok || got != "String" {
		t.Fatalf("local MailItem diagnostic shadowing = %q, %v; want String, true", got, ok)
	}
}
