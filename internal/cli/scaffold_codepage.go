package cli

import (
	"archive/zip"
	"fmt"
	"io"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

// seedScaffoldCodePage records the created/imported project's encoding once.
// UTF-8 source files and template projects keep their own encoding contracts.
func seedScaffoldCodePage(root string) error {
	cfg, err := config.Load(root)
	if err != nil {
		return err
	}
	codePage, err := workbookCodePage(workbookArgPath(root, cfg.Excel.Path))
	if err != nil {
		return fmt.Errorf("read scaffold VBA codepage: %w", err)
	}
	return config.UpdateBlankCodePage(filepath.Join(root, config.FileName), codePage)
}

func workbookCodePage(path string) (uint16, error) {
	archive, err := zip.OpenReader(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = archive.Close() }()
	for _, part := range archive.File {
		if !strings.EqualFold(part.Name, "xl/vbaProject.bin") {
			continue
		}
		stream, err := part.Open()
		if err != nil {
			return 0, err
		}
		body, readErr := io.ReadAll(stream)
		closeErr := stream.Close()
		if readErr != nil {
			return 0, readErr
		}
		if closeErr != nil {
			return 0, closeErr
		}
		project, err := cfb.Open(body)
		if err != nil {
			return 0, err
		}
		dir, ok := project.Stream("VBA/dir")
		if !ok {
			return 0, fmt.Errorf("VBA/dir stream is missing")
		}
		plain, err := ovba.Decompress(dir)
		if err != nil {
			return 0, err
		}
		info, err := ovba.ParseDir(plain)
		if err != nil {
			return 0, err
		}
		if info.CodePage == 0 {
			return 0, fmt.Errorf("VBA project codepage is missing")
		}
		return info.CodePage, nil
	}
	return 0, fmt.Errorf("xl/vbaProject.bin is missing")
}
