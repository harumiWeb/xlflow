package userforms

import "testing"

func TestParseExtractsNestedMSFormsControls(t *testing.T) {
	source := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Caption = "Customer"
   Begin MSForms.TextBox txtName
      Height = 18
   End
   Begin Forms.CommandButton.1 cmdOK
      Caption = "OK"
   End
End
Attribute VB_Name = "CustomerForm"
`
	form := Parse(source)
	if form.Name != "CustomerForm" {
		t.Fatalf("form name = %q, want CustomerForm", form.Name)
	}
	if len(form.Controls) != 2 {
		t.Fatalf("controls = %d, want 2: %+v", len(form.Controls), form.Controls)
	}
	if form.Controls[0].Name != "txtName" || form.Controls[0].Type != "MSForms.TextBox" {
		t.Fatalf("unexpected first control: %+v", form.Controls[0])
	}
	if form.Controls[1].Name != "cmdOK" || form.Controls[1].Type != "MSForms.CommandButton" {
		t.Fatalf("unexpected second control: %+v", form.Controls[1])
	}
}
