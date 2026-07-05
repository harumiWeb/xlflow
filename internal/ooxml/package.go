package ooxml

import (
	"archive/zip"
	"fmt"
	"io"
	"path"
	"strings"
)

type Package struct {
	rc    *zip.ReadCloser
	files map[string]*zip.File
}

func Open(filePath string) (*Package, error) {
	rc, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	p := &Package{rc: rc, files: map[string]*zip.File{}}
	for _, f := range rc.File {
		p.files[cleanPartName(f.Name)] = f
	}
	return p, nil
}

func (p *Package) Close() error {
	if p == nil || p.rc == nil {
		return nil
	}
	return p.rc.Close()
}

func (p *Package) openPart(name string) (io.ReadCloser, error) {
	f := p.files[cleanPartName(name)]
	if f == nil {
		return nil, fmt.Errorf("OOXML part not found: %s", name)
	}
	return f.Open()
}

func cleanPartName(name string) string {
	name = strings.ReplaceAll(name, "\\", "/")
	name = strings.TrimPrefix(name, "/")
	return path.Clean(name)
}
