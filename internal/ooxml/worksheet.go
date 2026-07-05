package ooxml

import (
	"encoding/xml"
	"io"
	"strings"
)

type Formula struct {
	Cell        string
	Type        string
	Ref         string
	SharedIndex string
	Text        string
}

func (p *Package) ReadWorksheetFormulas(partPath string) (formulas []Formula, err error) {
	rc, err := p.openPart(partPath)
	if err != nil {
		return nil, err
	}
	defer func() {
		if closeErr := rc.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	decoder := xml.NewDecoder(rc)
	currentCell := ""
	inFormula := false
	var formula Formula
	var text strings.Builder
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "c":
				currentCell = ""
				for _, attr := range t.Attr {
					if attr.Name.Local == "r" {
						currentCell = attr.Value
						break
					}
				}
			case "f":
				inFormula = true
				formula = Formula{Cell: currentCell}
				text.Reset()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "t":
						formula.Type = attr.Value
					case "ref":
						formula.Ref = attr.Value
					case "si":
						formula.SharedIndex = attr.Value
					}
				}
			}
		case xml.CharData:
			if inFormula {
				text.Write([]byte(t))
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "f":
				formula.Text = ensureFormulaPrefix(strings.TrimSpace(text.String()))
				formulas = append(formulas, formula)
				inFormula = false
			case "c":
				currentCell = ""
			}
		}
	}
	return formulas, nil
}
