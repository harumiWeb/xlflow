package compiler

import (
	"bytes"
	"encoding/hex"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const excelPictureFixture = "testdata/pictures-excel-authored"

func readExcelPictureForm(t *testing.T, file string) *oforms.Form {
	t.Helper()
	body, err := os.ReadFile(filepath.Join(excelPictureFixture, file))
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	names := oforms.DiscoverForms(container)
	if len(names) != 1 {
		t.Fatalf("DiscoverForms = %q, want one Excel-authored fixture form", names)
	}
	form, err := oforms.ReadForm(container, names[0], 932)
	if err != nil {
		t.Fatal(err)
	}
	return form
}

func authoredControl(form *oforms.Form, name string) (*oforms.Control, *oforms.Level) {
	for _, level := range form.Levels {
		for _, control := range level.Controls {
			if control.Name == name {
				return control, level
			}
		}
	}
	return nil, nil
}

func formSpecControlIndex(t *testing.T, snapshot spec.FormSpec, name string) int {
	t.Helper()
	for index, control := range snapshot.Controls {
		if control.Name == name {
			return index
		}
	}
	t.Fatalf("projected control %q is missing", name)
	return -1
}

func TestExcelAuthoredBMPAndJPEGLoadPicturePersistAsValidatedBMP(t *testing.T) {
	environmentBody, err := os.ReadFile(filepath.Join(excelPictureFixture, "environment.json"))
	if err != nil {
		t.Fatal(err)
	}
	var environment struct {
		BeforeSave []struct {
			Name             string `json:"name"`
			Type             int    `json:"type"`
			Width            int    `json:"width"`
			Height           int    `json:"height"`
			PictureSizeMode  int    `json:"pictureSizeMode"`
			PictureAlignment int    `json:"pictureAlignment"`
		} `json:"beforeSave"`
		Reopened []struct {
			Name             string `json:"name"`
			Type             int    `json:"type"`
			Width            int    `json:"width"`
			Height           int    `json:"height"`
			PictureSizeMode  int    `json:"pictureSizeMode"`
			PictureAlignment int    `json:"pictureAlignment"`
		} `json:"reopened"`
	}
	if err := json.Unmarshal(environmentBody, &environment); err != nil {
		t.Fatal(err)
	}
	for _, observations := range [][]struct {
		Name             string `json:"name"`
		Type             int    `json:"type"`
		Width            int    `json:"width"`
		Height           int    `json:"height"`
		PictureSizeMode  int    `json:"pictureSizeMode"`
		PictureAlignment int    `json:"pictureAlignment"`
	}{environment.BeforeSave, environment.Reopened} {
		for _, observation := range observations {
			if observation.Type != 1 || observation.Width != 508 || observation.Height != 339 || observation.PictureSizeMode != 3 || observation.PictureAlignment != 3 {
				t.Fatalf("Excel environment observation = %#v, want Type=1, HIMETRIC=508x339, mode/alignment=3/3", observation)
			}
		}
	}
	for _, fixtureName := range []string{"baseline.bin", "normalized.bin"} {
		t.Run(fixtureName, func(t *testing.T) {
			form := readExcelPictureForm(t, fixtureName)
			projected, err := projection.Project(form)
			if err != nil {
				t.Fatal(err)
			}
			for _, name := range []string{"ImageBmp", "ImageJpeg"} {
				control, level := authoredControl(form, name)
				if control == nil || level == nil || control.Record == nil {
					t.Fatalf("fixture control %s was not found", name)
				}
				raw := control.Record.Pictures["Picture"]
				if len(raw) < 24 || hex.EncodeToString(raw[:16]) != "0452e30b918fce119de300aa004bb851" || !bytes.Equal(raw[16:20], []byte{0x6c, 0x74, 0, 0}) {
					t.Fatalf("%s has an unexpected GuidAndPicture preamble: %x", name, raw[:min(len(raw), 24)])
				}
				asset, err := picture.Decode(raw)
				if err != nil {
					t.Fatalf("decode Excel-authored %s picture: %v", name, err)
				}
				if asset.Format != "bmp" || asset.Width != 24 || asset.Height != 16 || !bytes.Equal(asset.Data, raw[24:]) {
					t.Fatalf("%s asset = %#v, want persisted BMP 24x16", name, asset)
				}
				if control.Record.Values["Picture"] != 0xffff || control.Record.Mask&(1<<10) == 0 {
					t.Fatalf("%s picture marker = %#x, mask = %#x", name, control.Record.Values["Picture"], control.Record.Mask)
				}
				if name == "ImageJpeg" && !strings.Contains(level.Path, "/i") {
					t.Fatalf("ImageJpeg is not nested in a container level: %q", level.Path)
				}
				var snapshotControl *spec.FormSpecControl
				for i := range projected.Controls {
					if projected.Controls[i].Name == name {
						snapshotControl = &projected.Controls[i]
						break
					}
				}
				if snapshotControl == nil || snapshotControl.Properties["pictureAlignment"] != 3 || snapshotControl.Properties["pictureSizeMode"] != 3 {
					t.Fatalf("projected %s display properties = %#v, want 3/3", name, snapshotControl)
				}
				if slices.Contains(snapshotControl.Unsupported, "pictureAlignment") || slices.Contains(snapshotControl.Unsupported, "pictureSizeMode") || !slices.Contains(snapshotControl.Unsupported, "picture") {
					t.Fatalf("projected %s unsupported markers = %q", name, snapshotControl.Unsupported)
				}
			}
		})
	}
}

func TestCompileNewImagePictureAndDisplayProperties(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(excelPictureFixture, "logo.jpg"))
	if err != nil {
		t.Fatal(err)
	}
	input := newSpec()
	input.Controls = []spec.FormSpecControl{{
		ID: "image1", Name: "Image1", Type: "Image",
		Properties: map[string]any{"pictureAlignment": float64(3), "pictureSizeMode": float64(3)},
		Picture:    &spec.FormSpecPicture{Path: "assets/logo.jpg", Data: bytes.Clone(data)},
	}}
	form, err := CompileNew(input, 1252)
	if err != nil {
		t.Fatal(err)
	}
	control := form.Controls[0]
	asset, err := picture.Decode(control.Record.Pictures["Picture"])
	if err != nil {
		t.Fatal(err)
	}
	if asset.Format != "jpeg" || !bytes.Equal(asset.Data, data) || control.Record.Values["PictureAlignment"] != 3 || control.Record.Values["PictureSizeMode"] != 3 {
		t.Fatalf("generated Image record = %#v, asset=%#v", control.Record, asset)
	}
	projected, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if projected.Controls[0].Properties["pictureAlignment"] != 3 || projected.Controls[0].Properties["pictureSizeMode"] != 3 {
		t.Fatalf("generated display properties = %#v", projected.Controls[0].Properties)
	}
	input.Controls[0].Picture.Data[0] ^= 0xff
	if !bytes.Equal(control.Record.Pictures["Picture"][24:], data) {
		t.Fatal("generated picture aliases authoring Data")
	}
}

func TestCompileNewRejectsUnsupportedPictureNoOpAndAcceptsExplicitRemove(t *testing.T) {
	input := newSpec()
	input.Controls = []spec.FormSpecControl{{
		ID: "image1", Name: "Image1", Type: "Image", Unsupported: []string{"picture"},
	}}
	form, err := CompileNew(input, 1252)
	detail, ok := errors.AsType[*Error](err)
	if form != nil || err == nil || !ok || detail.Code != GenerationUnsupported {
		t.Fatalf("CompileNew = %v, %v; want unsupported picture metadata", form, err)
	}
	input.Controls[0].Picture = &spec.FormSpecPicture{Remove: true}
	form, err = CompileNew(input, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if form.Controls[0].Record.Mask&(1<<10) != 0 || len(form.Controls[0].Record.Pictures) != 0 {
		t.Fatalf("explicit remove generated picture data: %#v", form.Controls[0].Record)
	}
}

func TestCompileEditsReplaceAndRemoveAuthoredImagePreservingOtherStreams(t *testing.T) {
	base := readExcelPictureForm(t, "baseline.bin")
	beforeBytes, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(base)
	if err != nil {
		t.Fatal(err)
	}
	index := formSpecControlIndex(t, before, "ImageBmp")
	jpeg, err := os.ReadFile(filepath.Join(excelPictureFixture, "logo.jpg"))
	if err != nil {
		t.Fatal(err)
	}
	after := snapshotCopy(t, before)
	after.Controls[index].Picture = &spec.FormSpecPicture{Path: `assets\logo.jpg`, Data: bytes.Clone(jpeg)}
	replaced, err := CompileEdits(base, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	updated, _ := authoredControl(replaced, "ImageBmp")
	asset, err := picture.Decode(updated.Record.Pictures["Picture"])
	if err != nil || asset.Format != "jpeg" || !bytes.Equal(asset.Data, jpeg) {
		t.Fatalf("replacement picture = %#v, err=%v", asset, err)
	}
	untouchedBefore, _ := authoredControl(base, "ImageJpeg")
	untouchedAfter, _ := authoredControl(replaced, "ImageJpeg")
	if !bytes.Equal(untouchedBefore.Record.Pictures["Picture"], untouchedAfter.Record.Pictures["Picture"]) {
		t.Fatal("editing root Image changed the nested Image picture")
	}
	unchanged, err := oforms.SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(beforeBytes, unchanged) {
		t.Fatalf("CompileEdits mutated base: err=%v", err)
	}

	remove := snapshotCopy(t, before)
	remove.Controls[index].Picture = &spec.FormSpecPicture{Remove: true}
	cleared, err := CompileEdits(base, before, remove, 932)
	if err != nil {
		t.Fatal(err)
	}
	removed, _ := authoredControl(cleared, "ImageBmp")
	if removed.Record.Mask&(1<<10) != 0 {
		t.Fatalf("Picture mask bit remains set: %#x", removed.Record.Mask)
	}
	if _, ok := removed.Record.Pictures["Picture"]; ok {
		t.Fatal("removed Picture bytes remain in the decoded record")
	}
	stillThere, _ := authoredControl(cleared, "ImageJpeg")
	if !bytes.Equal(untouchedBefore.Record.Pictures["Picture"], stillThere.Record.Pictures["Picture"]) {
		t.Fatal("removing root Image changed the nested Image picture")
	}
}

func TestCompileTemplatePictureActionsAndMarkerReconciliation(t *testing.T) {
	base := readExcelPictureForm(t, "baseline.bin")
	snapshot, err := projection.Project(base)
	if err != nil {
		t.Fatal(err)
	}
	desired := templateInput(t, base, snapshot)
	index := formSpecControlIndex(t, snapshot, "ImageBmp")
	bmp, err := os.ReadFile(filepath.Join(excelPictureFixture, "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	caption := "Issue 912 picture push verification"
	desired.Form.Caption = &caption
	desired.Form.Build = &spec.FormSpecBuildForm{Caption: &caption}
	desired.Controls[index].Picture = &spec.FormSpecPicture{Path: "assets/logo.bmp", Data: bytes.Clone(bmp)}
	beforeInput, err := cloneFormSpec(desired)
	if err != nil {
		t.Fatal(err)
	}
	replaced, err := CompileTemplate(base, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(beforeInput, desired) {
		t.Fatal("CompileTemplate mutated picture action or source bytes")
	}
	rootImage, _ := authoredControl(replaced, "ImageBmp")
	asset, err := picture.Decode(rootImage.Record.Pictures["Picture"])
	if err != nil || asset.Format != "bmp" || !bytes.Equal(asset.Data, bmp) {
		t.Fatalf("template replacement = %#v, err=%v", asset, err)
	}
	nestedBeforeReplacement, _ := authoredControl(base, "ImageJpeg")
	nestedAfterReplacement, _ := authoredControl(replaced, "ImageJpeg")
	if !bytes.Equal(nestedBeforeReplacement.Record.Pictures["Picture"], nestedAfterReplacement.Record.Pictures["Picture"]) {
		t.Fatal("template replacement changed unrelated nested picture bytes")
	}

	removeInput := templateInput(t, base, snapshot)
	removeInput.Controls[index].Picture = &spec.FormSpecPicture{Remove: true}
	cleared, err := CompileTemplate(base, removeInput, 932)
	if err != nil {
		t.Fatal(err)
	}
	removed, _ := authoredControl(cleared, "ImageBmp")
	if removed.Record.Mask&(1<<10) != 0 || removed.Record.Pictures["Picture"] != nil {
		t.Fatalf("template remove retained Picture resource: %#v", removed.Record)
	}
	otherBefore, _ := authoredControl(base, "ImageJpeg")
	otherAfter, _ := authoredControl(cleared, "ImageJpeg")
	if !bytes.Equal(otherBefore.Record.Pictures["Picture"], otherAfter.Record.Pictures["Picture"]) {
		t.Fatal("template remove changed unrelated nested picture bytes")
	}

	warnings := []spec.FormSpecWarning{
		{Code: "unsupported_properties", Control: "ImageBmp", Message: "Unsupported Designer properties were preserved only in the binary model: mouseIcon, picture."},
		{Code: "unsupported_properties", Control: "ImageJpeg", Message: "Unsupported Designer properties were preserved only in the binary model: picture."},
	}
	remaining := reconcileTemplatePictureWarnings(warnings, desired.Controls)
	if len(remaining) != 2 || remaining[0].Message != "Unsupported Designer properties were preserved only in the binary model: mouseIcon." || remaining[1] != warnings[1] {
		t.Fatalf("picture warning reconciliation = %#v", remaining)
	}
}

func TestCloneFormSpecPreservesPictureDataWithoutAliasing(t *testing.T) {
	input := spec.FormSpec{Controls: []spec.FormSpecControl{{
		Type: "Frame", Name: "Frame", Controls: []spec.FormSpecControl{{
			Type: "Image", Name: "Image", Picture: &spec.FormSpecPicture{Path: "image.bmp", Data: []byte{1, 2, 3}},
		}},
	}}}
	cloned, err := cloneFormSpec(input)
	if err != nil {
		t.Fatal(err)
	}
	data := cloned.Controls[0].Controls[0].Picture.Data
	if !bytes.Equal(data, []byte{1, 2, 3}) {
		t.Fatalf("cloned picture data = %v", data)
	}
	data[0] = 9
	if input.Controls[0].Controls[0].Picture.Data[0] != 1 {
		t.Fatal("cloneFormSpec shares picture bytes with its input")
	}
}
