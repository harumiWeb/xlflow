package workbookuse

import "errors"

var ErrUnsupportedPlatform = errors.New("workbook-use detection is unavailable on this platform")

type State struct {
	OpenInExcel bool
	ExcelPIDs   []uint32
}
