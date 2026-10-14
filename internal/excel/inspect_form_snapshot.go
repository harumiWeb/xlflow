package excel

import (
	"archive/zip"
	"fmt"
	"io"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

func projectPersistedFormDimensions(snapshotPath, formName, basis string, formsPayload *any) error {
	if basis != "designer" && basis != "both" {
		return nil
	}
	if formsPayload == nil {
		return fmt.Errorf("inspect-form response has no forms payload")
	}

	size, err := readRootDisplayedSize(snapshotPath, formName)
	if err != nil {
		return err
	}
	forms, ok := (*formsPayload).(map[string]any)
	if !ok {
		return fmt.Errorf("inspect-form response forms payload has type %T", *formsPayload)
	}
	var designer map[string]any
	if basis == "both" {
		designer, ok = forms["designer"].(map[string]any)
	} else {
		// A single-basis designer response uses the form snapshot directly.
		designer = forms
		ok = true
	}
	if !ok || designer == nil {
		return fmt.Errorf("inspect-form response has no designer snapshot for %q", formName)
	}
	snapshotName, ok := designer["name"].(string)
	if !ok || !strings.EqualFold(snapshotName, formName) {
		return fmt.Errorf("inspect-form designer snapshot is for %q, want %q", snapshotName, formName)
	}

	if observed, exists := designer["client_width"]; exists {
		designer["observed_client_width"] = observed
	}
	if observed, exists := designer["client_height"]; exists {
		designer["observed_client_height"] = observed
	}
	designer["client_width"] = float64(size.Width) * 72 / 2540
	designer["client_height"] = float64(size.Height) * 72 / 2540
	return nil
}

func readRootDisplayedSize(snapshotPath, formName string) (oforms.Size, error) {
	workbook, err := zip.OpenReader(snapshotPath)
	if err != nil {
		return oforms.Size{}, fmt.Errorf("open private workbook snapshot: %w", err)
	}
	defer func() { _ = workbook.Close() }()

	var projectEntry *zip.File
	for _, entry := range workbook.File {
		if entry.Name != "xl/vbaProject.bin" {
			continue
		}
		if projectEntry != nil {
			return oforms.Size{}, fmt.Errorf("private workbook snapshot contains duplicate xl/vbaProject.bin entries")
		}
		projectEntry = entry
	}
	if projectEntry == nil {
		return oforms.Size{}, fmt.Errorf("private workbook snapshot is missing xl/vbaProject.bin")
	}
	stream, err := projectEntry.Open()
	if err != nil {
		return oforms.Size{}, fmt.Errorf("open private workbook VBA project: %w", err)
	}
	projectData, readErr := io.ReadAll(stream)
	closeErr := stream.Close()
	if readErr != nil {
		return oforms.Size{}, fmt.Errorf("read private workbook VBA project: %w", readErr)
	}
	if closeErr != nil {
		return oforms.Size{}, fmt.Errorf("close private workbook VBA project stream: %w", closeErr)
	}

	container, err := cfb.Open(projectData)
	if err != nil {
		return oforms.Size{}, fmt.Errorf("parse private workbook VBA project: %w", err)
	}
	dirCompressed, ok := container.Stream("VBA/dir")
	if !ok {
		return oforms.Size{}, fmt.Errorf("private workbook VBA project is missing VBA/dir")
	}
	dirPlain, err := ovba.Decompress(dirCompressed)
	if err != nil {
		return oforms.Size{}, fmt.Errorf("decompress private workbook VBA directory: %w", err)
	}
	dirInfo, err := ovba.ParseDir(dirPlain)
	if err != nil {
		return oforms.Size{}, fmt.Errorf("parse private workbook VBA directory: %w", err)
	}
	formStorages := oforms.DiscoverForms(container)
	formStorageIndex := slices.IndexFunc(formStorages, func(storage string) bool {
		return strings.EqualFold(storage, formName)
	})
	if formStorageIndex < 0 {
		return oforms.Size{}, fmt.Errorf("persisted UserForm %q was not found", formName)
	}
	formStorage := formStorages[formStorageIndex]
	form, err := oforms.ReadForm(container, formStorage, dirInfo.CodePage)
	if err != nil {
		return oforms.Size{}, fmt.Errorf("read persisted UserForm %q: %w", formStorage, err)
	}
	if len(form.Levels) == 0 || form.Levels[0] == nil || form.Levels[0].Record == nil {
		return oforms.Size{}, fmt.Errorf("persisted UserForm %q has no root MS-OFORMS record", formStorage)
	}
	size, ok := form.Levels[0].Record.Sizes["DisplayedSize"]
	if !ok {
		return oforms.Size{}, fmt.Errorf("persisted UserForm %q has no root DisplayedSize", formStorage)
	}
	if size.Width < 0 || size.Height < 0 {
		return oforms.Size{}, fmt.Errorf("persisted UserForm %q has negative root DisplayedSize", formStorage)
	}
	return size, nil
}
