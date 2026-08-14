package analyze

import (
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func projectConstantValues(files []parsedFile, typeDB *vbadb.DB) map[string]constexpr.Value {
	documents := make([]lint.ConstantValueDocument, 0, len(files))
	for _, file := range files {
		documents = append(documents, lint.ConstantValueDocument{Source: string(file.Source), IR: &file.IR})
	}
	return lint.ProjectConstantValues(documents, typeDB)
}
