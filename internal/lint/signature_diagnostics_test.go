package lint

import (
	"context"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestProcedureSignatureDiagnostics(t *testing.T) {
	source := `Option Explicit
Type Payload
    Value As Long
End Type
Sub Bad(Optional first As Long = 1, required As Long)
End Sub
Sub UDT(ByVal value As Payload, Optional other As Payload)
End Sub
Sub Args(ParamArray values() As String)
End Sub
Sub OptionalArgs(Optional first As Long = 1, ParamArray rest() As Variant)
End Sub
Sub TooMany(` + manyParameters(61) + `)
End Sub
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB048"); len(got) != 6 {
		t.Fatalf("VB048 = %+v, want six structural signature findings", got)
	}
	wantKinds := map[string]bool{
		"required_after_optional": true, "udt_byval": true, "udt_optional": true,
		"paramarray_shape": true, "paramarray_after_optional": true, "parameter_limit": true,
	}
	for _, issue := range issuesByCode(issues, "VB048") {
		if !wantKinds[issue.Kind] {
			t.Fatalf("unexpected VB048 kind %q: %+v", issue.Kind, issues)
		}
	}
	if got := issuesByCode(issues, "VB049"); len(got) != 0 {
		t.Fatalf("unexpected property findings: %+v", got)
	}
}

func TestProcedureParameterLimitAccepts60AndRejects61(t *testing.T) {
	for _, tc := range []struct {
		count int
		want  int
	}{
		{count: 60, want: 0},
		{count: 61, want: 1},
	} {
		source := "Sub Limit(" + manyParameters(tc.count) + ")\nEnd Sub\n"
		issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.cls", []byte(source))
		if err != nil {
			t.Fatal(err)
		}
		got := 0
		for _, issue := range issues {
			if issue.Kind == "parameter_limit" {
				got++
			}
		}
		if got != tc.want {
			t.Fatalf("parameter count %d produced %d limit findings, want %d: %+v", tc.count, got, tc.want, issues)
		}
	}
}

func TestSignatureDiagnosticsPreflightFindingsCannotBeSuppressed(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Sub Run(Optional first As Long = 1, ByVal required As Long) ' xlflow:disable-next-line VB048
End Sub
`)
	result, err := (Linter{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(result.Issues, "VB048"); len(got) != 1 {
		t.Fatalf("VB048 should remain unsuppressed: issues=%+v warnings=%+v", result.Issues, result.Warnings)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VB048") {
		t.Fatalf("expected unsupported suppression warning, got %+v", result.Warnings)
	}
}

func TestPropertySignatureDiagnostics(t *testing.T) {
	source := `Option Explicit
Property Get Value(ByVal index As Long) As String
End Property
Property Let Value(ByRef index As String, ByVal value As Long)
End Property
Property Set Broken(ByVal value As Long)
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB049")
	if len(got) != 3 {
		t.Fatalf("VB049 = %+v, want three property contract findings", got)
	}
	wantKinds := map[string]bool{"property_index_shape": true, "property_value_type": true, "set_value_type": true}
	for _, issue := range got {
		if !wantKinds[issue.Kind] {
			t.Fatalf("unexpected VB049 kind %q: %+v", issue.Kind, got)
		}
	}
}

func TestPropertyIndexAllowsImplicitVariantAgainstExplicitVariant(t *testing.T) {
	source := `Property Get Item(index) As String
End Property
Property Let Item(ByRef index As Variant, ByVal value As String)
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB049"); len(got) != 0 {
		t.Fatalf("implicit Variant index must match an explicit Variant index: %+v", got)
	}
}

func TestOptionalDefaultIdentifierInfIsUnknown(t *testing.T) {
	source := `Sub Defaults(Optional value As String = Inf, Optional nanValue As String = NaN)
End Sub
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB048"); len(got) != 0 {
		t.Fatalf("Inf/NaN references should remain unknown: %+v", got)
	}
}

func TestPropertySetterValueModifiers(t *testing.T) {
	source := `Property Let OptionalValue(Optional value As String)
End Property
Property Set ParamArrayValue(ParamArray value() As Variant)
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB049")
	if len(got) != 2 {
		t.Fatalf("VB049 = %+v, want setter value modifier findings", got)
	}
	if got[0].Kind != "setter_value_optional" || got[1].Kind != "setter_value_paramarray" {
		t.Fatalf("unexpected setter modifier findings: %+v", got)
	}
}

func TestPropertyClassObjectTypeUsesResolvedObjectCategory(t *testing.T) {
	source := `Property Get Ref() As Thing
End Property
Property Set Ref(ByVal value As Thing)
End Property
Property Let Broken(ByVal value As Thing)
End Property
`
	issues, err := (Linter{ObjectTypeDeclarations: map[string]int{"thing": 1}}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issuesByCode(issues, "VB049") {
		if issue.Kind == "set_value_type" {
			t.Fatalf("resolved class object was rejected by Property Set: %+v", issue)
		}
	}
	got := 0
	for _, issue := range issuesByCode(issues, "VB049") {
		if issue.Kind == "let_value_type" {
			got++
		}
	}
	if got != 1 {
		t.Fatalf("resolved class object Property Let findings = %d, want one: %+v", got, issues)
	}
}

func TestPropertySetterAllowsOptionalIndexBeforeValue(t *testing.T) {
	source := `Property Set Picture(Optional ByVal appearance As Long, ByVal value As Object)
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB048"); len(got) != 0 {
		t.Fatalf("VB048 should allow an Optional property index before the setter value: %+v", got)
	}
	if got := issuesByCode(issues, "VB049"); len(got) != 0 {
		t.Fatalf("VB049 should allow an Optional property index before the setter value: %+v", got)
	}
}

func TestPropertySetAllowsVariantObjectValuePair(t *testing.T) {
	source := `Property Get Item(ByVal key As Variant) As Variant
End Property
Property Let Item(ByVal key As Variant, ByVal value As Variant)
End Property
Property Set Item(ByVal key As Variant, ByVal value As Object)
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB049"); len(got) != 0 {
		t.Fatalf("VB049 should allow the Variant/Object Property Set pair: %+v", got)
	}
}

func TestPropertyCrossAccessorFindingPointsAtLaterAccessor(t *testing.T) {
	source := `Property Let Reverse(ByVal value As Long)
End Property
Property Get Reverse() As String
End Property
`
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Thing.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB049")
	if len(got) != 1 || got[0].Kind != "property_value_type" || got[0].Line != 3 {
		t.Fatalf("VB049 = %+v, want later getter name range", got)
	}
}

func TestSignatureDiagnosticsFailOpenOnParserRecovery(t *testing.T) {
	source := "Sub Broken(Optional value As Long = Foo(\nEnd Sub\n"
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB048"); len(got) != 0 {
		t.Fatalf("VB048 should fail open on parser recovery: %+v", got)
	}
	if got := issuesByCode(issues, "VB049"); len(got) != 0 {
		t.Fatalf("VB049 should fail open on parser recovery: %+v", got)
	}
}

func TestOptionalDefaultExpressionIsUnknown(t *testing.T) {
	source := "Sub Defaults(Optional computed As Long = 1 + 2, Optional reference As Long = SomeConst, Optional call As Long = Foo())\nEnd Sub\n"
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB048"); len(got) != 1 || got[0].Kind != "optional_default_nonconstant" {
		t.Fatalf("VB048 = %+v, want only callable default finding", got)
	}
}

func TestOptionalDefaultLiteralTypeMismatch(t *testing.T) {
	source := "Sub Defaults(Optional text As Long = \"bad\", Optional flag As Boolean = 1)\nEnd Sub\n"
	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB048")
	if len(got) != 2 {
		t.Fatalf("VB048 = %+v, want two definite literal mismatches", got)
	}
	for _, issue := range got {
		if issue.Kind != "optional_default_type" {
			t.Fatalf("unexpected default kind: %+v", got)
		}
	}
}

func TestOptionalDefaultNumericLiteralForms(t *testing.T) {
	for _, value := range []string{"1&", "&H10", "&O10", "1%", "1!", "1#", "1@", "1^", "- &H10"} {
		if !isNumericLiteral(value) {
			t.Fatalf("%q should be recognized as a VBA numeric literal", value)
		}
		if !optionalDefaultTypeMismatch(value, "String", signatureTypeIntrinsic) {
			t.Fatalf("%q should mismatch a String Optional parameter", value)
		}
		if !optionalDefaultTypeMismatch(value, "Boolean", signatureTypeIntrinsic) {
			t.Fatalf("%q should mismatch a Boolean Optional parameter", value)
		}
	}

	issues, err := (Linter{}).LintSourceContext(context.Background(), "Main.bas", []byte("Sub Defaults(Optional text As String = 1&)\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB048")
	if len(got) != 1 || got[0].Kind != "optional_default_type" {
		t.Fatalf("numeric Optional default should produce VB048 optional_default_type: %+v", issues)
	}
}

func TestOptionalDefaultStringSuffixIsUnknown(t *testing.T) {
	if isNumericLiteral("1$") {
		t.Fatal("String suffix must not be treated as numeric")
	}
}

func TestOptionalDefaultDateLiteralHandling(t *testing.T) {
	const value = "#1/1/2024#"
	if optionalDefaultTypeMismatch(value, "Date", signatureTypeIntrinsic) {
		t.Fatal("date literal should match a Date Optional parameter")
	}
	if !optionalDefaultTypeMismatch(value, "Long", signatureTypeIntrinsic) {
		t.Fatal("date literal should mismatch a non-Date Optional parameter")
	}
}

func manyParameters(n int) string {
	params := ""
	for i := 1; i <= n; i++ {
		if i > 1 {
			params += ", "
		}
		params += "ByVal p" + intString(i) + " As Long"
	}
	return params
}
