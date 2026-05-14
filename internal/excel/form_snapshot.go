package excel

import "github.com/harumiWeb/xlflow/internal/excel/forms"

type FormSnapshotOutput = forms.SnapshotOutput
type FormSpec = forms.FormSpec
type FormSpecForm = forms.FormSpecForm
type FormSpecControl = forms.FormSpecControl
type FormSpecWarning = forms.FormSpecWarning

func ResolveFormSnapshotOutput(root, outPath string) (FormSnapshotOutput, error) {
	return forms.ResolveSnapshotOutput(root, outPath)
}

func WriteFormSnapshot(output FormSnapshotOutput, spec FormSpec) error {
	return forms.WriteSnapshot(output, spec)
}

func FormSpecFromInspectSnapshot(snapshot any) (FormSpec, error) {
	return forms.FormSpecFromInspectSnapshot(snapshot)
}
